/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-drawing-extent-negative.docx");
const CX_ONLY = path.join(HERE, "fixtures", "broken", "package.word-drawing-extent-negative-cx.docx");
const CY_ONLY = path.join(HERE, "fixtures", "broken", "package.word-drawing-extent-negative-cy.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-drawing-extent-positive.docx");
const CODE = "word-drawing-extent-negative";

describe("word-valid negative drawing extent", { timeout: 30000 }, () => {
    it("rejects a document whose drawing extents are negative", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error" && issue.code === CODE);
        expect(errors).toHaveLength(8);
    });

    it("rejects a document with one negative cx", async () => {
        const result = await validate(CX_ONLY, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error" && issue.code === CODE);
        expect(errors).toHaveLength(1);
        expect(errors[0]?.message).toContain('cx="-9525"');
    });

    it("rejects a document with one negative cy", async () => {
        const result = await validate(CY_ONLY, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error" && issue.code === CODE);
        expect(errors).toHaveLength(1);
        expect(errors[0]?.message).toContain('cy="-12065"');
    });

    it("does not reject the document once every extent is positive", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === CODE)).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === CODE)).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
    });
});
