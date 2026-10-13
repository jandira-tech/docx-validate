/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-undeclared-part.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-undeclared-part-removed.docx");

describe("word-valid part with no content type", { timeout: 30000 }, () => {
    it("rejects a package part that [Content_Types].xml does not declare", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(
            result.issues.some((issue) => issue.severity === "error" && issue.code === "ct-uncovered-part"),
        ).toBe(true);
    });

    it("does not reject the document once that part is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "ct-uncovered-part")).toBe(false);
        expect(result.valid).toBe(true);
    });
});
