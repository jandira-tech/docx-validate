/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-sdt-id-empty.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-sdt-id-empty-repaired.docx");

describe("word-valid empty structured-document-tag id", { timeout: 30000 }, () => {
    it("rejects a w:sdt whose w:id w:val is empty", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error");
        expect(errors.some((issue) => issue.code === "word-sdt-id-empty")).toBe(true);
    });

    it("does not reject the document once that id is a number", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-sdt-id-empty")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient or strict", async () => {
        for (const profile of ["lenient", "strict"] as const) {
            const result = await validate(BROKEN, { profile });
            expect(result.issues.some((issue) => issue.code === "word-sdt-id-empty")).toBe(false);
        }
    });
});
