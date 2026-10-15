/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-duplicate-settings.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-duplicate-settings-removed.docx");

describe("word-valid duplicate settings relationship", { timeout: 30000 }, () => {
    it("rejects a document relationships part with two settings relationships", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-duplicate-settings-relationship")).toBe(
            true,
        );
    });

    it("does not reject the document once only one settings relationship remains", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-settings-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-settings-relationship")).toBe(false);
    });
});
