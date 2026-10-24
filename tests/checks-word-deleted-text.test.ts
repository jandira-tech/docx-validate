/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordTextboxDeletedTextCheck, wordMoveFromDeletedTextCheck } from "../src/scripts/office/checks/word-deleted-text";

describe("word-textbox-deleted-text", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:del><w:txbxContent><w:p><w:r><w:delText>x</w:delText></w:r></w:p></w:txbxContent></w:del>"),
        });
        expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-textbox-deleted-text",
                path: "word/document.xml",
                message:
                    "A text box contains deleted text. Microsoft Word offers to recover the file. Remove the w:delText or w:delInstrText from the text box.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                "<w:p><w:r><w:delText>x</w:delText><w:delInstrText>y</w:delInstrText></w:r></w:p><w:txbxContent><w:del><w:p><w:r><w:delText>x</w:delText><w:delInstrText>y</w:delInstrText></w:r></w:p></w:del></w:txbxContent>",
            ),
        });
        expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordTextboxDeletedTextCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordTextboxDeletedTextCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-move-from-deleted-text", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:p><w:moveFrom><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:p>"),
        });
        expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-move-from-deleted-text",
                path: "word/document.xml",
                message:
                    "A move source contains deleted text. Microsoft Word offers to recover the file. Use w:t inside w:moveFrom, not w:delText.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:p><w:moveFrom><w:r><w:t>x</w:t></w:r></w:moveFrom><w:r><w:delText>y</w:delText></w:r></w:p>"),
        });
        expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordMoveFromDeletedTextCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordMoveFromDeletedTextCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("distinguishes deletion within a text box from an outside deletion and ordinary body text", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(
            `<w:txbxContent><w:p><w:r><w:delInstrText>x</w:delInstrText></w:r></w:p></w:txbxContent><w:p><w:r><w:delText>body</w:delText></w:r></w:p>`,
        ),
    });
    expect(await wordTextboxDeletedTextCheck.run(ctx)).toHaveLength(1);
    expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([]);
});
it("retains the historical tolerance for fallback text boxes", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(
            `<mc:AlternateContent><mc:Choice Requires="wps" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:p/></mc:Choice><mc:Fallback><w:txbxContent><w:delText>x</w:delText></w:txbxContent></mc:Fallback></mc:AlternateContent>`,
        ),
    });
    expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([]);
});
