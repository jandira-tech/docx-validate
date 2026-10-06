/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Platform-agnostic XSD validator interface backed by `libxml2-wasm`.
 *
 * Part 1 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §3.
 *
 * The goal: every place that does XSD validation today (currently
 * `BaseValidator` calling `libxmljs2` directly) goes through this interface.
 * PR B reworks `BaseValidator` to take an injected `XsdValidator`; PR C wires
 * the four top-level classes (`Validate` / `Repair` / `Normalize` / `Measure`)
 * on top.
 *
 * Why WASM: `libxmljs2` ships native bindings and only runs in Node. The WASM
 * port (https://www.npmjs.com/package/libxml2-wasm) wraps the same C library
 * and runs in both Node and the browser — that's what unlocks the
 * one-bundle-two-runtimes architecture from the spec.
 */

import type { ValidationIssue } from "./types";

/**
 * Platform-agnostic XSD validator contract.
 *
 * `validate(xml, schemaPath)` parses the XML, loads the XSD from disk (Node
 * side; browser consumers will eventually inject a different provider), and
 * returns a `ValidationIssue[]` matching the shape used everywhere else in
 * the library. An empty array means "passes XSD". Each issue has severity
 * `error` plus a stable `code` (`xsd-validation-failed`, `xml-parse-error`,
 * or `xsd-schema-load-failed` for an unloadable schema) so tests can assert
 * on issues without matching free-form error prose.
 */
export type XsdValidator = {
    readonly validate: (xml: string, schemaPath: string) => Promise<ValidationIssue[]>;
};

let memoizedValidator: Promise<XsdValidator> | undefined;

/**
 * Returns the default WASM-backed validator. Memoised — one WASM init per
 * process. Power users (e.g. tests injecting a fake validator, or jubarte-first
 * supplying a different engine) can build their own `XsdValidator`-shaped
 * object and pass it through the constructors documented in the spec.
 */
export const createXsdValidator = (): Promise<XsdValidator> => {
    if (!memoizedValidator) {
        memoizedValidator = buildWasmValidator();
    }
    return memoizedValidator;
};

/**
 * Reset the memoised factory. Tests only — production code should never call
 * this. Used to force a fresh WASM init between independent test cases.
 */
export const _resetXsdValidatorMemo = (): void => {
    memoizedValidator = undefined;
};

/**
 * One-time registration of Node-side fs input providers so libxml2-wasm can
 * resolve relative `<xs:import schemaLocation="../mce/mc.xsd"/>` references
 * inside the bundled OOXML schemas. Without this, `XsdValidator.fromDoc()`
 * throws on the first unresolved import.
 *
 * Browser consumers will need a buffer-backed input provider; that's PR C's
 * concern — for PR A the validator is Node-only by design.
 */
let fsProvidersRegistered = false;
const ensureFsProviders = async (): Promise<void> => {
    if (fsProvidersRegistered) {
        return;
    }
    const { xmlRegisterFsInputProviders } = await import("libxml2-wasm/lib/nodejs.mjs");
    xmlRegisterFsInputProviders();
    fsProvidersRegistered = true;
};

const buildWasmValidator = async (): Promise<XsdValidator> => {
    const { XmlDocument, XsdValidator: WasmXsdValidator, XmlValidateError } = await import("libxml2-wasm");
    const { readFile } = await import("node:fs/promises");
    await ensureFsProviders();

    // Cache of compiled schema validators, mirroring the libxmljs2-era
    // `_xsdCache` in validators/base.ts: the OOXML schema bundle is ~1.1 MB
    // and every XML part in a package reuses the same handful of schemas, so
    // recompiling per file would dominate runtime. The cache stores the
    // in-flight load PROMISE: concurrent first-loads of the same schema then
    // compile exactly one validator instead of racing to overwrite the entry
    // (each race loser would hold a compiled validator that is never
    // disposed — a wasm-heap leak). A failed load evicts its entry so the
    // next call retries fresh. Entries live for the process lifetime (a
    // handful of schemas; each holds its compiled form in the wasm heap) and
    // are dropped wholesale by `_resetXsdValidatorMemo`, which discards this
    // closure.
    const validatorCache = new Map<string, Promise<InstanceType<typeof WasmXsdValidator>>>();

    const loadValidator = (schemaPath: string): Promise<InstanceType<typeof WasmXsdValidator>> => {
        const cached = validatorCache.get(schemaPath);
        if (cached) {
            return cached;
        }
        const load = (async () => {
            // Read schema file directly, parse with fromString. fsInputProviders
            // resolves any relative <xs:import schemaLocation="..."/> references
            // encountered during the parse, allowing OOXML schemas with imports
            // to load cleanly when they're all present on disk. The document base
            // URL makes those imports resolve against the schema file's
            // directory, not the process cwd.
            const schemaSource = await readFile(schemaPath, "utf-8");
            const schemaDoc = XmlDocument.fromString(schemaSource, { url: schemaPath });
            try {
                return WasmXsdValidator.fromDoc(schemaDoc);
            } finally {
                // Dispose even when compilation throws — an undisposed source
                // doc stays resident in the wasm heap.
                schemaDoc.dispose();
            }
        })().catch((err: unknown) => {
            validatorCache.delete(schemaPath);
            throw err;
        });
        validatorCache.set(schemaPath, load);
        return load;
    };

    const validate = async (xml: string, schemaPath: string): Promise<ValidationIssue[]> => {
        let validator;
        try {
            validator = await loadValidator(schemaPath);
        } catch (loadErr) {
            // Load failures are errors, not info: a validator must not claim
            // a document is valid against a schema it could not load. The
            // known-noisy bundled case (opc-coreProperties.xsd's remote
            // Dublin Core import) is string-filtered one layer up in
            // `validators/base.ts` (IGNORED_VALIDATION_ERRORS), mirroring the
            // Python original's message-based suppression.
            const message = loadErr instanceof Error ? loadErr.message : String(loadErr);
            return [
                {
                    severity: "error",
                    code: "xsd-schema-load-failed",
                    message: `Schema load failed (${schemaPath}): ${message}`,
                    path: schemaPath,
                },
            ];
        }

        let xmlDoc;
        try {
            xmlDoc = XmlDocument.fromString(xml);
        } catch (parseErr) {
            const message = parseErr instanceof Error ? parseErr.message : String(parseErr);
            return [{ severity: "error", code: "xml-parse-error", message }];
        }

        try {
            validator.validate(xmlDoc);
            return [];
        } catch (err) {
            if (err instanceof XmlValidateError) {
                return err.details.map((d) => ({
                    severity: "error" as const,
                    code: "xsd-validation-failed",
                    message: d.message,
                    ...(d.file !== undefined ? { path: d.file } : {}),
                    ...(d.line !== undefined ? { line: d.line } : {}),
                }));
            }
            const message = err instanceof Error ? err.message : String(err);
            return [{ severity: "error", code: "xsd-validation-failed", message }];
        } finally {
            xmlDoc.dispose();
        }
    };

    return { validate };
};
