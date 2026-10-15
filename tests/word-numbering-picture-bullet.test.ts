/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-numbering-picture-bullet.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-numbering-picture-bullet-removed.docx");

describe("word-valid numbering picture bullet", { timeout: 30000 }, () => {
    it("rejects a numbering level whose picture bullet id has no w:numPicBullet", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-numbering-picture-bullet")).toBe(true);
    });

    it("does not reject the document once that picture bullet id is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-numbering-picture-bullet")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "word-numbering-picture-bullet")).toBe(false);
    });
});
