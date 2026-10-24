/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordTextContainsParagraphCheck } from "../src/scripts/office/checks/word-text-paragraph";

describe("word-text-contains-paragraph", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<w:p><w:r><w:t><w:p/></w:t></w:r></w:p>") });
        expect(await wordTextContainsParagraphCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-text-contains-paragraph",
                path: "word/document.xml",
                message:
                    "A <w:t> element contains 1 <w:p> element(s). Microsoft Word cannot open the file. Move the paragraph out of the text run.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<w:p><w:r><w:t>text</w:t></w:r></w:p>") });
        expect(await wordTextContainsParagraphCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordTextContainsParagraphCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordTextContainsParagraphCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
