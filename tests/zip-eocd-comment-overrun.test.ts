/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.zip-eocd-comment-overrun.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.zip-eocd-comment-declared.docx");

describe("word-valid ZIP comment length past the end of the file", () => {
    it("rejects a comment length the file does not contain", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "zip-eocd-comment-overrun")).toBe(true);
    });

    it("does not reject the same package once the comment length matches the bytes present", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "zip-eocd-comment-overrun")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "zip-eocd-comment-overrun")).toBe(false);
    });
});
