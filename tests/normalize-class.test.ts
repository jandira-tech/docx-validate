/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";
import { Normalize } from "../src/normalize";

const WORKING_FIXTURE = path.resolve(
    __dirname,
    "fixtures",
    "working",
    "sample-document.afterword-repaired-word-repaired.docx",
);

describe("Normalize class (PR C task C.3)", () => {
    it("returns a NormalizeResult shape (bytes + changed + passes)", async () => {
        const bytes = readFileSync(WORKING_FIXTURE);
        const result = await new Normalize().run(new Uint8Array(bytes));
        expect(result.bytes).toBeInstanceOf(Uint8Array);
        expect(result.bytes.byteLength).toBeGreaterThan(0);
        expect(typeof result.changed).toBe("boolean");
        expect(typeof result.passes.mergeRuns).toBe("number");
    }, 30_000);

    it("is content-idempotent — second pass reports changed=false (mergeRuns count = 0)", async () => {
        // Content-level idempotency: once mergeRuns has collapsed sibling
        // runs on first call, the second call finds nothing to merge.
        // BYTE-level idempotency is NOT asserted here — JSZip's pack is
        // non-deterministic (CLAUDE.md note 3) so a.bytes !== b.bytes
        // even when the unpacked XML content is identical.
        const bytes = readFileSync(WORKING_FIXTURE);
        await new Normalize().run(new Uint8Array(bytes));
        const a = await new Normalize().run(new Uint8Array(bytes));
        const b = await new Normalize().run(a.bytes);
        expect(b.changed).toBe(false);
        expect(b.passes.mergeRuns).toBe(0);
    }, 60_000);
});
