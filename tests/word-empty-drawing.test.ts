/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-empty-drawing.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-empty-drawing-removed.docx");

describe("word-valid empty drawing", { timeout: 30000 }, () => {
    it("rejects an empty w:drawing inside a markup-compatibility choice", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(
            result.issues.some((issue) => issue.severity === "error" && issue.code === "word-empty-drawing"),
        ).toBe(true);
    });

    it("does not reject the document once that empty drawing is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-empty-drawing")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient or strict", async () => {
        for (const profile of ["lenient", "strict"] as const) {
            const result = await validate(BROKEN, { profile });
            expect(result.issues.some((issue) => issue.code === "word-empty-drawing")).toBe(false);
        }
    });
});
