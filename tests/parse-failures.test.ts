/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { withoutConsequentialParseFailures } from "../src/lib/parse-failures";
import type { ValidationIssue } from "../src/lib/types";

const parse: ValidationIssue = {
    code: "comment-marker-parse",
    severity: "error",
    path: "word/document.xml",
    message: "Error parsing XML",
};

describe("withoutConsequentialParseFailures", () => {
    it("suppresses only parse restatements for rejected parts, without mutating the findings", () => {
        const kept: ValidationIssue[] = [
            { ...parse, path: "word/comments.xml" },
            { ...parse, path: undefined },
            { ...parse, code: "comment-orphan-start" },
            { ...parse, code: undefined },
        ];
        const issues = [parse, ...kept, { ...parse, code: "comment-thread-parse" }];
        expect(withoutConsequentialParseFailures(issues, new Set(["word/document.xml"]))).toEqual(kept);
        expect(issues).toHaveLength(6);
    });

    it("retains native parse diagnostics when XML diagnostics accepted every part", () => {
        expect(withoutConsequentialParseFailures([parse], new Set())).toEqual([parse]);
    });
});
