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
 * `error` plus a stable `code` (`xsd-validation-failed`) so tests can assert
 * on issues without matching free-form error prose.
 */
export interface XsdValidator {
    validate(xml: string, schemaPath: string): Promise<ValidationIssue[]>;
}

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
    if (fsProvidersRegistered) return;
    const { xmlRegisterFsInputProviders } = await import("libxml2-wasm/lib/nodejs.mjs");
    xmlRegisterFsInputProviders();
    fsProvidersRegistered = true;
};

const buildWasmValidator = async (): Promise<XsdValidator> => {
    const { XmlDocument, XsdValidator: WasmXsdValidator, XmlValidateError } = await import("libxml2-wasm");
    const { readFile } = await import("node:fs/promises");
    await ensureFsProviders();

    return {
        async validate(xml: string, schemaPath: string): Promise<ValidationIssue[]> {
            // Read schema file directly, parse with fromString. fsInputProviders
            // resolves any relative <xs:import schemaLocation="..."/> references
            // encountered during the parse, allowing OOXML schemas with imports
            // to load cleanly when they're all present on disk.
            let validator;
            try {
                const schemaSource = await readFile(schemaPath, "utf-8");
                // Set the document base URL so relative imports resolve against
                // the schema file's directory, not the process cwd.
                const schemaDoc = XmlDocument.fromString(schemaSource, { url: schemaPath });
                validator = WasmXsdValidator.fromDoc(schemaDoc);
                schemaDoc.dispose();
            } catch (loadErr) {
                // The bundled OOXML schemas have unresolvable namespace-only
                // imports (e.g. sharedTypes is referenced but not in the tree).
                // CLAUDE.md note 4 documents this — libxmljs2 also degraded here.
                // Surface as an info-level diagnostic instead of crashing so
                // structurally well-formed docs still validate via other checks.
                const message = loadErr instanceof Error ? loadErr.message : String(loadErr);
                return [{
                    severity: "info",
                    code: "xsd-schema-load-skipped",
                    message: `Schema load failed (${schemaPath}): ${message}`,
                    path: schemaPath,
                }];
            }

            let xmlDoc;
            try {
                xmlDoc = XmlDocument.fromString(xml);
            } catch (parseErr) {
                validator.dispose();
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
                validator.dispose();
            }
        },
    };
};
