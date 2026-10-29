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
import { Repair } from "../src/repair";

const WORKING_FIXTURE = path.resolve(__dirname, "fixtures", "working", "sample-document.afterword-repaired-word-repaired.docx");

describe("Repair class (PR C task C.2)", () => {
    it("instantiates with no options", () => {
        expect(new Repair()).toBeInstanceOf(Repair);
    });

    it("returns a RepairResult shape (bytes + repairs + diagnostics)", async () => {
        const bytes = readFileSync(WORKING_FIXTURE);
        const result = await new Repair().run(new Uint8Array(bytes));
        expect(result.bytes).toBeInstanceOf(Uint8Array);
        expect(result.bytes.byteLength).toBeGreaterThan(0);
        expect(typeof result.repairs).toBe("number");
        expect(Array.isArray(result.diagnostics)).toBe(true);
    }, 30_000);
});
