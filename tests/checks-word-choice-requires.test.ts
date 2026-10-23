/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordChoiceRequiresUndeclaredCheck } from "../src/scripts/office/checks/word-choice-requires";

describe("word-choice-requires-undeclared", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                '<mc:AlternateContent><mc:Choice Requires="future"><w:p xmlns:future="urn:future"/></mc:Choice><mc:Fallback xmlns:future="urn:future"><w:p/></mc:Fallback></mc:AlternateContent>',
            ),
        });
        expect(await wordChoiceRequiresUndeclaredCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-choice-requires-undeclared",
                path: "word/document.xml",
                message:
                    "mc:Choice Requires lists undeclared prefix future. Microsoft Word offers to recover the file. Declare the prefix on the Choice element or an ancestor.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                '<mc:AlternateContent><mc:Choice xmlns:future="urn:future" Requires="future xml w"/><mc:Choice/><mc:Choice Requires=" "/></mc:AlternateContent>',
            ),
        });
        expect(await wordChoiceRequiresUndeclaredCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordChoiceRequiresUndeclaredCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordChoiceRequiresUndeclaredCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
