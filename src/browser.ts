/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Browser entry — full DOCX validation, entirely client-side.
//
// Unzips the .docx in memory (JSZip), validates every part through the same
// DOCXSchemaValidator the Node build uses, but backed by MemoryPartFS and the
// WASM XSD engine (libxml2-wasm, wasm inlined). No `node:fs`, no native binding,
// no temp dir. The schema graph is supplied as an in-memory bundle keyed by path
// relative to the schema root (the WASM engine resolves xs:include/import from
// it). Verdicts match the Node validator.

import { MemoryPartFS } from "./lib/part-fs";
import { createBrowserEngine } from "./lib/xsd-engine/index";
import { DEFAULT_PROFILE, type Profile, type ValidationResult } from "./lib/types";
import { DOCXSchemaValidator } from "./scripts/office/validators/docx";

export { MemoryPartFS } from "./lib/part-fs";
export { createBrowserEngine } from "./lib/xsd-engine/index";
export type { Profile, ValidationResult, ValidationIssue } from "./lib/types";

export interface ValidateDocxBrowserOptions {
    /** OOXML schema set: path-relative-to-root → XSD text. Required (no disk). */
    schemaBundle: Record<string, string>;
    /** Validation profile. Defaults to the library default ("lenient"). */
    profile?: Profile;
}

export type DocxBytes = ArrayBuffer | Uint8Array | Blob;

async function toUint8Array(input: DocxBytes): Promise<Uint8Array> {
    if (input instanceof Uint8Array) return input;
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    if (typeof Blob !== "undefined" && input instanceof Blob) return new Uint8Array(await input.arrayBuffer());
    throw new TypeError("validateDocx: expected ArrayBuffer, Uint8Array, or Blob");
}

/**
 * Validate a `.docx` fully in the browser. Returns the same structured
 * {@link ValidationResult} as the Node validator.
 */
export async function validateDocx(input: DocxBytes, options: ValidateDocxBrowserOptions): Promise<ValidationResult> {
    const { default: JSZip } = await import("jszip");
    const bytes = await toUint8Array(input);
    const zip = await JSZip.loadAsync(bytes);

    const parts: Array<[string, Uint8Array]> = [];
    for (const name of Object.keys(zip.files)) {
        const entry = zip.files[name];
        if (entry.dir) continue;
        parts.push([name, await entry.async("uint8array")]);
    }

    const xsdEngine = createBrowserEngine({ schemaBundle: options.schemaBundle });
    await xsdEngine.init();

    const validator = new DOCXSchemaValidator({
        partFS: new MemoryPartFS(parts),
        // The WASM engine resolves schemas from `schemaBundle`, so schemasDir is
        // only a label; pass a sentinel so the constructor never probes disk.
        schemasDir: "/__schemas__",
        schemaBundle: options.schemaBundle,
        xsdEngine,
        profile: options.profile ?? DEFAULT_PROFILE,
    });

    return validator.validate();
}
