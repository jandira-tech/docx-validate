/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, W, WS, rels, relation } from "./support/word-rule-test-xml";
import { wordFootnoteReferenceMissingCheck, wordEndnoteReferenceMissingCheck } from "../src/scripts/office/checks/word-footnote-reference";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";

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
            "_rels/.rels": rels(relation("rId0", "officeDocument", "", "word/document.xml")),
            "word/document.xml": document(`<w:footnoteReference w:id="2"/><w:footnoteReference/>`),
            "word/_rels/document.xml.rels": rels(relation("rId1", "footnotes", "", "footnotes.xml")),
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

describe("endnote references (#784)", () => {
    it.each([W, WS])("ignores an endnote reference inside a footnote under %s (session 16)", async (namespace) => {
        const ctx = ValidationContext.fromEntries({
            "word/footnotes.xml": `<w:footnotes xmlns:w="${namespace}"><w:footnote w:id="1"><w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:footnote></w:footnotes>`,
        });
        expect(await wordEndnoteReferenceMissingCheck.run(ctx)).toEqual([]);
    });

    it("registers the endnote check with the footnote prerequisites and group", () => {
        const check = createDefaultCheckRegistry().get("word-endnote-reference-missing");
        expect(check).toBe(wordEndnoteReferenceMissingCheck);
        expect(check?.requires).toEqual(["xml-syntax"]);
        expect(check?.groups).toEqual(["story"]);
        expect(check?.formats).toEqual(["docx"]);
    });

    it.each([W, WS])("checks missing and present ids under %s", async (namespace) => {
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": rels(relation("rId0", "officeDocument", "", "word/document.xml")),
            "word/document.xml": document(
                '<w:p><w:r><w:endnoteReference w:id="1"/><w:endnoteReference w:id="5"/><w:endnoteReference/></w:r></w:p>',
                namespace,
            ),
            "word/_rels/document.xml.rels": rels(relation("rId1", "endnotes", "", "endnotes.xml")),
            "word/endnotes.xml": `<w:endnotes xmlns:w="${namespace}"><w:endnote w:id="5"/><w:endnote w:id="5"/><w:endnote/></w:endnotes>`,
        });
        expect(await wordEndnoteReferenceMissingCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-endnote-reference-missing",
                path: "word/document.xml",
                message:
                    "An endnote reference uses id 1, and this package has no w:endnote with that id. Microsoft Word offers to recover the file. Remove the w:endnoteReference, or add the matching endnote.",
            },
        ]);
    });

    it("does not let a footnote satisfy an endnote reference and skips malformed parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": rels(relation("rId0", "officeDocument", "", "word/document.xml")),
            "word/document.xml": document('<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>'),
            "word/_rels/document.xml.rels": rels(relation("rId1", "endnotes", "", "endnotes.xml")),
            "word/footnotes.xml": document('<w:footnote w:id="1"/>'),
            "word/endnotes.xml": "<broken",
        });
        expect(await wordEndnoteReferenceMissingCheck.run(ctx)).toEqual([
            expect.objectContaining({ code: "word-endnote-reference-missing" }),
        ]);
        expect(await wordEndnoteReferenceMissingCheck.run(ValidationContext.fromEntries({ "word/document.xml": "<broken" }))).toEqual([]);
    });
});
