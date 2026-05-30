/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * PR C Task C.1: assert the new `Validate` class.
 *
 * Uses `validate-class.test.ts` (not `validate.test.ts`) to avoid colliding
 * with the existing `tests/validate.test.ts` that covers the legacy
 * path-taking `validate()` function.
 */

import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";
import { Validate } from "../src/validate";

const WORKING_DIR = path.resolve(__dirname, "fixtures", "working");
const BROKEN_DIR = path.resolve(__dirname, "fixtures", "broken");

const WORKING_FIXTURE = path.join(WORKING_DIR, "sample-document.afterword-repaired-word-repaired.docx");
const BROKEN_FIXTURE = path.join(BROKEN_DIR, "sample-document.broken-tables.docx");

describe("Validate class (PR C task C.1)", () => {
    it("instantiates with no options", () => {
        const v = new Validate();
        expect(v).toBeInstanceOf(Validate);
    });

    it("accepts xsdValidator + schemasDir + profile options without throwing", () => {
        const v = new Validate({
            schemasDir: "/dev/null/non-existent",
            profile: "strict",
        });
        expect(v).toBeInstanceOf(Validate);
    });

    it("run() returns a ValidationResult shape (valid + issues array)", async () => {
        const bytes = readFileSync(WORKING_FIXTURE);
        const result = await new Validate().run(new Uint8Array(bytes));
        expect(typeof result.valid).toBe("boolean");
        expect(Array.isArray(result.issues)).toBe(true);
    }, 30_000);

    it("run() flags a broken-by-design fixture as invalid", async () => {
        const bytes = readFileSync(BROKEN_FIXTURE);
        const result = await new Validate().run(new Uint8Array(bytes));
        expect(result.valid).toBe(false);
        expect(result.issues.length).toBeGreaterThan(0);
    }, 30_000);
});
