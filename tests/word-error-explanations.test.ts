/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 */

import { describe, expect, it } from "vitest";
import { WORD_ERROR_EXPLANATIONS, explainWordError } from "../src/scripts/office/validators/word-error-explanations";

describe("word error explanations", () => {
    it("explains the unconditional Word-blocking codes", () => {
        // The codes isWordBlockingIssue always treats as Word-blocking must each
        // carry a user-facing explanation so the CLI can tell the user why Word fails.
        const alwaysBlocking = [
            "ignorable-undeclared",
            "word-math-spre-body",
            "word-math-parse",
            "word-content-type-invalid",
            "word-drawing-scalar-whitespace",
            "id-durable-overflow",
            "comment-thread-commentid-paraid-orphan",
            "comment-thread-durableid-orphan",
        ];
        for (const code of alwaysBlocking) {
            expect(WORD_ERROR_EXPLANATIONS[code], `missing explanation for ${code}`).toBeTruthy();
        }
    });

    it("explainWordError returns the mapped string or undefined", () => {
        expect(explainWordError("ignorable-undeclared")).toContain("mc:Ignorable");
        expect(explainWordError("not-a-real-code")).toBeUndefined();
        expect(explainWordError(undefined)).toBeUndefined();
    });
});
