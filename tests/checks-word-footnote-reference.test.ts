/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordFootnoteReferenceMissingCheck } from "../src/scripts/office/checks/word-footnote-reference";

describe("word-footnote-reference-missing", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document('<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'),
        });
        expect(await wordFootnoteReferenceMissingCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-footnote-reference-missing",
                path: "word/document.xml",
                message:
                    "A footnote reference uses id 2, and this package has no w:footnote with that id. Microsoft Word offers to recover the file. Remove the w:footnoteReference, or add the matching footnote.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:footnoteReference w:id="2"/><w:footnoteReference/>`),
            "word/footnotes.xml": document(`<w:footnote w:id="2"/><w:footnote w:id="2"/><w:footnote/>`),
        });
        expect(await wordFootnoteReferenceMissingCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordFootnoteReferenceMissingCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordFootnoteReferenceMissingCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
