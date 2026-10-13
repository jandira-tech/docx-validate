/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-duplicate-theme.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-duplicate-theme-removed.docx");

describe("word-valid duplicate theme relationship", () => {
    it("rejects a document relationships part with two theme relationships", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(
            result.issues.some(
                (issue) => issue.severity === "error" && issue.code === "word-duplicate-theme-relationship",
            ),
        ).toBe(true);
    });

    it("does not reject the document once only one theme relationship remains", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-theme-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient or strict", async () => {
        for (const profile of ["lenient", "strict"] as const) {
            const result = await validate(BROKEN, { profile });
            expect(result.issues.some((issue) => issue.code === "word-duplicate-theme-relationship")).toBe(false);
        }
    });
});
