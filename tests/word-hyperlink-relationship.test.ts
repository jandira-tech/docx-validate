/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-hyperlink-relationship.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-hyperlink-relationship-repaired.docx");

describe("word-valid hyperlink relationship", () => {
    it("rejects a w:hyperlink whose r:id is not an external hyperlink relationship", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-hyperlink-relationship")).toBe(true);
    });

    it("does not reject the same hyperlink once that relationship is external", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
    });
});
