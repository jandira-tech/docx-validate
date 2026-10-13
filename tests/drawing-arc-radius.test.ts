/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.drawing-arc-radius.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.drawing-arc-radius-integer.docx");

describe("word-valid non-integer DrawingML arc radius", () => {
    it("rejects an a:arcTo radius Word cannot open", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(
            result.issues.some((issue) => issue.severity === "error" && issue.code === "word-drawing-arc-radius"),
        ).toBe(true);
    });

    it("does not reject the same document once every arc radius is an integer", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-drawing-arc-radius")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient or strict", async () => {
        for (const profile of ["lenient", "strict"] as const) {
            const result = await validate(BROKEN, { profile });
            expect(result.issues.some((issue) => issue.code === "word-drawing-arc-radius")).toBe(false);
        }
    });
});
