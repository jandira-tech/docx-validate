/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

// In-memory DOCX repair — no disk, no temp dir. Runs the repair suite over a
// set of parts held in a `MemoryPartFS`. Repair parses/serializes with
// `@xmldom/xmldom` only (no native `libxmljs2`), so the *repair logic* itself
// is pure JS.
//
// Browser status: the repair LOGIC is disk-free and native-free, but the
// validator class it lives on (`base.ts`) still statically imports `node:fs`
// (DiskPartFS default, schema reads) and `libxmljs2` (XSD validation). Fully
// bundling this for the browser requires decoupling those static imports from
// the repair import graph — tracked separately. In Node this runs with no
// per-part disk I/O.

import type { Profile } from "../../lib/types";
import { MemoryPartFS } from "../../lib/part-fs";
import { DOCXSchemaValidator } from "./validators/docx";

export interface RepairInMemoryResult {
    /** Repaired parts as `[relativePath, bytes]`, ready to repack into a zip. */
    parts: Array<[string, Buffer]>;
    /** Number of repairs applied (0 ⇒ nothing changed). */
    repairs: number;
}

/**
 * Repair an unpacked DOCX held entirely in memory.
 *
 * @param parts   `[relativePath, content]` pairs (e.g. from a JSZip unpack).
 * @param options `profile` selects the repair/validation profile (default
 *                `"word-valid"`).
 */
export async function repairDocxInMemory(
    parts: Iterable<[string, string | Buffer]>,
    options: { profile?: Profile } = {},
): Promise<RepairInMemoryResult> {
    const partFS = new MemoryPartFS(parts);
    const validator = new DOCXSchemaValidator({
        partFS,
        profile: options.profile ?? "word-valid",
        // Repair never reads XSD schemas; pass an explicit (unused) dir so the
        // constructor does not probe the filesystem via defaultSchemasDir().
        schemasDir: "__unused_for_repair__",
    });
    const repairs = await validator.repair();
    return { parts: partFS.entries(), repairs };
}
