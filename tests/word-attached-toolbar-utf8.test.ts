/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-attached-toolbar-utf8.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-attached-toolbar-utf8-repaired.docx");

describe("word-valid attached toolbar UTF-8 replacement", { timeout: 30000 }, () => {
    it("rejects a toolbarData target that contains EF BF BD", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error");
        expect(errors.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(true);
    });

    it("does not reject the toolbar once those sequences are gone", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-attached-toolbar-utf8")).toBe(true);
    });
});
