/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import {
    TABLE_BORDERS_PARENTS,
    TABLE_CELL_SPACING_PARENTS,
    TBLPRBASE_ORDER,
    wordTablePropertiesMisplacedCheck as check,
} from "../src/scripts/office/checks/word-table-properties";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const document = (body: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

const styles = (body: string) => `<?xml version="1.0"?><w:styles xmlns:w="${W}">${body}</w:styles>`;

describe("word-table-properties-misplaced: tblBorders", () => {
    it("reports tblBorders ordered after tblLayout inside tblPr, as in the probed table_tests_v2.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr><w:tblW/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top/></w:tblBorders></w:tblPr>` +
                    `<w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:tblBorders follows w:tblLayout inside w:tblPr;/),
            }),
        ]);
    });

    it("reports tblBorders ordered after tblCellMar and tblBorders under a parent that cannot hold it", async () => {
        const afterCellMar = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr><w:tblCellMar/><w:tblBorders/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await check.run(afterCellMar)).toEqual([
            expect.objectContaining({ code: check.id, message: expect.stringMatching(/follows w:tblCellMar/) }),
        ]);
        const wrongParent = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:tbl><w:tblBorders/></w:tbl>`),
        });
        expect(await check.run(wrongParent)).toEqual([
            expect.objectContaining({
                code: check.id,
                message: expect.stringMatching(/w:tblBorders sits directly inside w:tbl;.*w:tblPr/),
            }),
        ]);
    });

    it("tolerates tblBorders under tcPr in a table-style conditional, as the Word-opened jubarte lane specimens carry it", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/styles.xml": styles(
                `<w:style w:type="table"><w:tblPr><w:tblBorders><w:top/></w:tblBorders></w:tblPr>` +
                    `<w:tblStylePr w:type="firstRow"><w:tcPr><w:tblBorders><w:bottom w:val="single"/></w:tblBorders></w:tcPr></w:tblStylePr></w:style>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
        expect(TABLE_BORDERS_PARENTS.has("tcPr")).toBe(true);
    });

    it("tolerates table-property order quirks in styles parts, as the Word-opened 1993 control carries them", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/styles.xml": styles(`<w:style w:type="table"><w:tblPr><w:tblCellMar/><w:tblBorders/></w:tblPr></w:style>`),
        });
        expect(await check.run(ctx)).toEqual([]);
        const header = ValidationContext.fromEntries({
            "word/header1.xml": `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:tbl><w:tblPr><w:tblLayout/><w:tblBorders/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:hdr>`,
        });
        expect(await check.run(header)).toEqual([
            expect.objectContaining({ code: check.id, path: "word/header1.xml", message: expect.stringMatching(/follows w:tblLayout/) }),
        ]);
    });

    it("accepts ordered tblBorders inside tblPr and tblPrEx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr><w:tblBorders><w:top/></w:tblBorders><w:tblLayout/><w:tblCellMar/></w:tblPr>` +
                    `<w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
                    `<w:tbl><w:tblPrEx><w:tblBorders/></w:tblPrEx><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });
});

describe("word-table-properties-misplaced: tblCellSpacing", () => {
    it("reports tblCellSpacing ordered after tblBorders inside tblPr, as in word-rejects-table-cellspacing.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr><w:tblW/><w:tblBorders><w:top/></w:tblBorders><w:tblCellSpacing w:w="20" w:type="dxa"/></w:tblPr>` +
                    `<w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                message: expect.stringMatching(/w:tblCellSpacing follows w:tblBorders inside w:tblPr;/),
            }),
        ]);
    });

    it("reports tblCellSpacing under a parent that cannot hold it, as the sectPr walk test builds it", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:sectPr><w:pgSz/><w:tblCellSpacing w:w="10" w:type="dxa"/></w:sectPr>`),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: check.id, message: expect.stringMatching(/w:tblCellSpacing sits directly inside w:sectPr;/) }),
        ]);
    });

    it("accepts ordered tblCellSpacing in tblPr and tblPrEx, and any position inside trPr", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr><w:tblCellSpacing w:w="20" w:type="dxa"/><w:tblBorders><w:top/></w:tblBorders></w:tblPr>` +
                    `<w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:trPr><w:trHeight/><w:tblCellSpacing w:w="20" w:type="dxa"/></w:trPr>` +
                    `<w:tc><w:p/></w:tc></w:tr></w:tbl>` +
                    `<w:tbl><w:tblPrEx><w:tblCellSpacing/><w:tblBorders/></w:tblPrEx><w:tblGrid><w:gridCol/></w:tblGrid>` +
                    `<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(TABLE_CELL_SPACING_PARENTS.has("trPr")).toBe(true);
        expect(await check.run(ctx)).toEqual([]);
        expect(TBLPRBASE_ORDER.indexOf("tblCellSpacing")).toBeLessThan(TBLPRBASE_ORDER.indexOf("tblBorders"));
        expect(TBLPRBASE_ORDER.indexOf("tblBorders")).toBeLessThan(TBLPRBASE_ORDER.indexOf("tblLayout"));
    });

    it("leaves foreign wrappers, part roots, and unparseable parts alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<x:wrap xmlns:x="urn:x"><w:tblBorders/></x:wrap>`),
            "word/root.xml": `<w:tblBorders xmlns:w="${W}"/>`,
            "word/broken.xml": "<w:tblBorders",
        });
        expect(ctx.xml("word/broken.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
