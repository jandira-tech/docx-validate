/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-cell-ends-with-table.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-cell-ends-with-table-repaired.docx");

describe("word-valid cell that ends with a table", { timeout: 30000 }, () => {
    it("rejects a table cell whose last child is w:tbl", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(
            result.issues.some((issue) => issue.severity === "error" && issue.code === "word-cell-ends-with-table"),
        ).toBe(true);
    });

    it("does not reject the document once that cell ends with w:p", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-cell-ends-with-table")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient or strict", async () => {
        for (const profile of ["lenient", "strict"] as const) {
            const result = await validate(BROKEN, { profile });
            expect(result.issues.some((issue) => issue.code === "word-cell-ends-with-table")).toBe(false);
        }
    });
});
