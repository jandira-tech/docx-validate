/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-anchor-open.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-anchor-open-repaired.docx");

describe("word-valid anchor attributes and position choice", { timeout: 30000 }, () => {
    it("rejects a wp:anchor missing locked, layoutInCell, or allowOverlap, and a position with both children", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error");
        const missing = errors.filter((issue) => issue.code === "word-anchor-missing-attribute");
        expect(missing.some((issue) => issue.message.includes("locked"))).toBe(true);
        expect(missing.some((issue) => issue.message.includes("layoutInCell"))).toBe(true);
        expect(missing.some((issue) => issue.message.includes("allowOverlap"))).toBe(true);
        const both = errors.filter((issue) => issue.code === "word-position-offset-and-align");
        expect(both.some((issue) => issue.message.includes("wp:positionH"))).toBe(true);
        expect(both.some((issue) => issue.message.includes("wp:positionV"))).toBe(true);
    });

    it("does not reject the document once those attributes are present and each position has one child", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-anchor-missing-attribute")).toBe(false);
        expect(result.issues.some((issue) => issue.code === "word-position-offset-and-align")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the causes on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-anchor-missing-attribute")).toBe(false);
        expect(result.issues.some((issue) => issue.code === "word-position-offset-and-align")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the causes on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-anchor-missing-attribute")).toBe(true);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-position-offset-and-align")).toBe(true);
    });
});
