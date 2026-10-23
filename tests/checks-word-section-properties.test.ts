/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import {
    SECTION_PROPERTIES_PARENTS,
    wordSectionPropertiesMisplacedCheck as check,
} from "../src/scripts/office/checks/word-section-properties";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const document = (body: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

describe("word-section-properties-misplaced", () => {
    it("reports a sectPr directly inside a paragraph, outside its pPr", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:p><w:r><w:t>a</w:t></w:r><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr></w:p>` +
                    `<w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:sectPr sits directly inside w:p;.*w:body.*w:pPr/),
            }),
        ]);
    });

    it("reports a sectPr inside a table cell", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:sectPr/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([expect.objectContaining({ code: check.id, message: expect.stringMatching(/inside w:tc;/) })]);
    });

    it("accepts the legal placements: last in the body, inside pPr, inside a sectPrChange snapshot, and in a glossary docPartBody", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr></w:pPr></w:p>` +
                    `<w:p><w:pPr><w:sectPr><w:sectPrChange w:id="1"><w:sectPr/></w:sectPrChange></w:sectPr></w:pPr></w:p>` +
                    `<w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr>`,
            ),
            "word/glossary/document.xml": `<?xml version="1.0"?><w:glossaryDocument xmlns:w="${W}"><w:docPart><w:docPartBody><w:p/><w:sectPr/></w:docPartBody></w:docPart></w:glossaryDocument>`,
        });
        expect(SECTION_PROPERTIES_PARENTS.has("docPartBody")).toBe(true);
        expect(await check.run(ctx)).toEqual([]);
    });

    it("accepts the Word-tolerated out-of-order sectPr children control", async () => {
        const { readFile } = await import("node:fs/promises");
        const ctx = await ValidationContext.fromArchive(
            new Uint8Array(await readFile("tests/fixtures/working/word-tolerated-misplaced-pgsz.docx")),
        );
        expect(await check.run(ctx)).toEqual([]);
    });

    it("leaves foreign wrappers, part roots, and unparseable parts alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<x:wrap xmlns:x="urn:x"><w:sectPr/></x:wrap>`),
            "word/root.xml": `<w:sectPr xmlns:w="${W}"/>`,
            "word/broken.xml": "<w:sectPr",
        });
        expect(ctx.xml("word/broken.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
