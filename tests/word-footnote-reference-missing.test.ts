/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-footnote-reference-missing.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-footnote-reference-present.docx");

describe("word-valid dangling footnote reference", { timeout: 30000 }, () => {
    it("rejects a footnote reference whose id is not a footnote id", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-footnote-reference-missing")).toBe(true);
    });

    it("does not reject the document once that reference is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-footnote-reference-missing")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-footnote-reference-missing")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-footnote-reference-missing")).toBe(true);
    });
});
