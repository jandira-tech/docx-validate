/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-row-without-cell.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-row-without-cell-repaired.docx");

describe("word-valid row without a cell", { timeout: 30000 }, () => {
    it("rejects a table row that contains no w:tc", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-row-without-cell")).toBe(true);
    });

    it("does not reject the document once that row contains a cell", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-row-without-cell")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "word-row-without-cell")).toBe(false);
    });
});
