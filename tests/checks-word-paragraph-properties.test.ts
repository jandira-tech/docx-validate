/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import {
    PARAGRAPH_PROPERTIES_PARENTS,
    wordParagraphPropertiesMisplacedCheck as check,
} from "../src/scripts/office/checks/word-paragraph-properties";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const document = (body: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

const styles = (body: string) => `<?xml version="1.0"?><w:styles xmlns:w="${W}">${body}</w:styles>`;

describe("word-paragraph-properties-misplaced", () => {
    it("reports a pPr nested inside pPr, as in 5Errors.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:pPr><w:pPr/><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr></w:p>`),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:pPr sits directly inside w:pPr;.*w:p\b.*w:style/),
            }),
        ]);
    });

    it("accepts the Word-tolerated duplicate pPr children of one paragraph", async () => {
        const { readFile } = await import("node:fs/promises");
        const fixture = await ValidationContext.fromArchive(
            new Uint8Array(await readFile("tests/fixtures/working/word-tolerated-duplicate-ppr.docx")),
        );
        expect(await check.run(fixture)).toEqual([]);
        const doubled = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:pPr><w:numPr/></w:pPr><w:pPr><w:spacing/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`),
        });
        expect(await check.run(doubled)).toEqual([]);
    });

    it("accepts every schema-admitted pPr parent", async () => {
        const parents = [...PARAGRAPH_PROPERTIES_PARENTS].filter((local) => local !== "pPrChange" && local !== "pPrDefault");
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:pPr/></w:p>`),
            "word/numbering.xml": `<?xml version="1.0"?><w:numbering xmlns:w="${W}"><w:lvl><w:pPr/></w:lvl></w:numbering>`,
            "word/styles.xml": styles(
                `<w:style w:type="paragraph"><w:pPr/></w:style><w:tblStylePr><w:pPr/></w:tblStylePr>` +
                    `<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
        expect(parents.length).toBeGreaterThanOrEqual(4);
    });

    it("accepts a pPr wrapped by a tracked-changes snapshot", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:pPr><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr></w:p>`),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("leaves foreign wrappers, part roots, and unparseable parts alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<x:wrap xmlns:x="urn:x"><w:pPr/></x:wrap>`),
            "word/root.xml": `<w:pPr xmlns:w="${W}"/>`,
            "word/broken.xml": "<w:pPr",
        });
        expect(ctx.xml("word/broken.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
