// cspell:words Ctxt DTDLOAD NOENT
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

import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

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

/**
 * Counts how the shipped engine handled a part. `statusOnlyPasses` is a
 * libxml2 status of 0 with no structured error object. `messageCollections`
 * is a non-zero status (or an internal failure) that still built messages.
 * Tests read this off the real `validate()` path.
 */
export const xsdValidateStats = {
    statusOnlyPasses: 0,
    messageCollections: 0,
};

// The optional status probe uses private fields from the pinned WASM release.
// Check them before entering C; fall back to the public API if they change.
const wasmPointer = (value: unknown): number | undefined => {
    if (typeof value !== "object" || value === null || !("_ptr" in value)) {
        return undefined;
    }
    const pointer = value._ptr;
    return typeof pointer === "number" && Number.isSafeInteger(pointer) && pointer > 0 ? pointer : undefined;
};

// Function-table callbacks live as long as the WASM module. Reuse this entry
// across factory resets instead of allocating an entry for each new factory.
let silentSchemaErrors: number | undefined;

/**
 * `xmlSchemaValidateDoc` status. A no-op structured callback stops libxml2's
 * default stderr writer. The callback does not build an error object.
 */
const schemaValidateStatus = (
    schema: unknown,
    doc: unknown,
    newCtxt: (schemaPtr: number) => number,
    freeCtxt: (ctx: number) => void,
    validateDoc: (ctx: number, docPtr: number) => number,
    silence: (ctx: number, fn: number, data: number) => void,
    silentFn: number,
): number | undefined => {
    const schemaPtr = wasmPointer(schema);
    const docPtr = wasmPointer(doc);
    if (schemaPtr === undefined || docPtr === undefined) {
        return undefined;
    }
    const ctx = newCtxt(schemaPtr);
    if (!ctx) {
        return -1;
    }
    try {
        // Default libxml2 handler writes schema errors to stderr. A no-op
        // keeps the status call quiet and does not build an error object.
        silence(ctx, silentFn, 0);
        return validateDoc(ctx, docPtr);
    } finally {
        freeCtxt(ctx);
    }
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
 * One-time registration of a RESTRICTED Node-side fs input provider so
 * libxml2-wasm can resolve relative `<xs:import schemaLocation="../mce/mc.xsd"/>`
 * references inside the bundled OOXML schemas. Without an input provider,
 * `XsdValidator.fromDoc()` throws on the first unresolved import.
 *
 * Restriction: `open` hands out a file descriptor ONLY for files under an
 * allowed schema root (see {@link allowSchemaResourceRoot}) and `match` claims
 * every request, so libxml2 never falls back to its own loader. Untrusted
 * OOXML therefore cannot make libxml2 open `file:///etc/passwd` (or anything
 * else outside the configured schemas directory) through a DTD or a schema
 * import; libxml2 reports the refusal as a load failure.
 *
 * Browser consumers will need a buffer-backed input provider; that's PR C's
 * concern — for PR A the validator is Node-only by design.
 */

/** The bundled OOXML schema tree shipped in the package. */
const bundledSchemasDir = (): string => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    // Source tree: src/lib/ → src/scripts/office/schemas. Built bundle:
    // dist/index.mjs → dist/schemas (copied there by vite.config.ts).
    const candidates = [path.resolve(here, "..", "scripts", "office", "schemas"), path.resolve(here, "schemas")];
    return candidates.find((candidate) => existsSync(candidate)) ?? (candidates[0] as string);
};

/** Absolute directories libxml2 may read XSD resources from. */
const schemaRoots = new Set<string>([bundledSchemasDir()]);

/** The allowed schema roots, sorted. Exported for diagnostics and tests. */
export const schemaResourceRoots = (): string[] => [...schemaRoots].sort();

/**
 * Allow libxml2 to read XSD resources from `root` (and anything under it).
 * `BaseSchemaValidator` calls this with its configured `schemasDir` so a
 * caller-supplied schema tree keeps working; the bundled tree is always
 * allowed.
 */
export const allowSchemaResourceRoot = (root: string): void => {
    schemaRoots.add(path.resolve(root));
};

/** Drop every explicitly registered root, keeping the bundled default. Tests only. */
export const _resetSchemaResourceRoots = (): void => {
    schemaRoots.clear();
    schemaRoots.add(bundledSchemasDir());
};

/** `file:` URL → filesystem path; anything else → `null` (network, data:, …). */
const toLocalPath = (filename: string): string | null => {
    if (!filename.startsWith("file:")) {
        return path.isAbsolute(filename) ? filename : null;
    }
    try {
        return fileURLToPath(filename);
    } catch {
        return null;
    }
};

const isInsideRoot = (root: string, target: string): boolean => {
    const relative = path.relative(root, target);
    return relative === "" || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative));
};

let fsProvidersRegistered = false;
const ensureFsProviders = async (): Promise<void> => {
    if (fsProvidersRegistered) {
        return;
    }
    const { fsInputProviders } = await import("libxml2-wasm/lib/nodejs.mjs");
    const { xmlRegisterInputProvider } = await import("libxml2-wasm/lib/libxml2.mjs");
    // `match` claims EVERY request so libxml2 always routes through us and
    // never falls back to its own default loader; the allow-list is enforced
    // in `open`, where refusing means "no file descriptor", not "try someone
    // else". The file mechanics themselves are libxml2-wasm's own.
    xmlRegisterInputProvider({
        match: () => true,
        open: (filename: string) => {
            const local = toLocalPath(filename);
            return local !== null && schemaResourceRoots().some((root) => isInsideRoot(root, local))
                ? fsInputProviders.open(filename)
                : undefined;
        },
        read: fsInputProviders.read,
        close: fsInputProviders.close,
    });
    fsProvidersRegistered = true;
};

const buildWasmValidator = async (): Promise<XsdValidator> => {
    const { ParseOption, XmlDocument, XsdValidator: WasmXsdValidator, XmlValidateError } = await import("libxml2-wasm");
    // Instance documents: never touch the network, never resolve an external
    // entity or external DTD. (No NOENT/DTDLOAD, so entities stay unresolved.)
    // libxml2's ParseOption is a bit-flag field, so combining flags requires a
    // bitwise OR.
    // oxlint-disable-next-line no-bitwise
    const INSTANCE_PARSE_OPTIONS = ParseOption.XML_PARSE_NONET | ParseOption.XML_PARSE_NO_XXE;
    // Schema documents: the network is refused, but local imports MUST resolve
    // (through the restricted input provider above), so NO_XXE is deliberately
    // NOT set here — it would refuse the OOXML import closure as well.
    const SCHEMA_PARSE_OPTIONS = ParseOption.XML_PARSE_NONET;
    const libxml2 = await import("libxml2-wasm/lib/libxml2.mjs");
    const { xmlSchemaFreeValidCtxt, xmlSchemaNewValidCtxt, xmlSchemaSetValidStructuredErrors, xmlSchemaValidateDoc } = libxml2;
    // addFunction is exported by the ESM build and missing from the .d.mts.
    const { addFunction } = libxml2 as unknown as { addFunction: (fn: () => void, signature: string) => number };
    const silentFn = typeof addFunction === "function" ? (silentSchemaErrors ??= addFunction(() => undefined, "vii")) : undefined;
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
            // Read schema file directly, parse with fromString. The registered
            // input provider resolves any relative
            // <xs:import schemaLocation="..."/> references encountered during
            // the parse — but only from an allowed schema root — allowing OOXML
            // schemas with imports to load cleanly offline when they are all
            // present on disk. The document base URL makes those imports resolve
            // against the schema file's directory, not the process cwd.
            const schemaSource = await readFile(schemaPath, "utf-8");
            const schemaDoc = XmlDocument.fromString(schemaSource, { url: schemaPath, option: SCHEMA_PARSE_OPTIONS });
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
            // a document is valid against a schema it could not load.
            // opc-coreProperties.xsd imports the bundled dublincore copies
            // (schemas/dublincore/). IGNORED_VALIDATION_ERRORS in
            // validators/base.ts is empty, so a load failure is reported
            // here and is never suppressed (#1209).
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
            xmlDoc = XmlDocument.fromString(xml, { option: INSTANCE_PARSE_OPTIONS });
        } catch (parseErr) {
            const message = parseErr instanceof Error ? parseErr.message : String(parseErr);
            return [{ severity: "error", code: "xml-parse-error", message }];
        }

        try {
            // Status 0 is a pass: the integer is enough, and building
            // XmlValidateError would throw away an empty detail list.
            // Status > 0 still collects messages — the ignore list and
            // isWordBlockingXsdIssue both match text the integer does not carry.
            const status =
                silentFn === undefined
                    ? undefined
                    : schemaValidateStatus(
                          validator,
                          xmlDoc,
                          xmlSchemaNewValidCtxt,
                          xmlSchemaFreeValidCtxt,
                          xmlSchemaValidateDoc,
                          xmlSchemaSetValidStructuredErrors,
                          silentFn,
                      );
            if (status === 0) {
                xsdValidateStats.statusOnlyPasses += 1;
                return [];
            }
            xsdValidateStats.messageCollections += 1;
            if (status !== undefined && status < 0) {
                return [
                    { severity: "error", code: "xsd-validation-failed", message: "XSD engine failed: invalid input or internal error" },
                ];
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
                return [{ severity: "error", code: "xsd-validation-failed", message: `XSD engine failed: ${message}` }];
            }
        } finally {
            xmlDoc.dispose();
        }
    };

    return { validate };
};
