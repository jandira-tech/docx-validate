/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-cell-margin-order`: `w:tblCellMar` / `w:tcMar` are `xsd:sequence`
 * types in `ISO-IEC29500-4_2016/wml.xsd` (`CT_TblCellMar` at :2458,
 * `CT_TcMar` at :2299). A margin child written after a later sibling is
 * schema-invalid; Microsoft Word reads margins by name and opens the file, so
 * the finding is an advisory.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { wordCellMarginOrderCheck } from "../src/scripts/office/checks/cell-margin-order";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";

/** The transitional order Word itself writes: no `start`/`end`. */
const legal = `<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>`;
const misordered = `<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>`;

const document = (body: string, ns = W): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`;

const table = (margins: string): string =>
    `<w:tbl><w:tblPr><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;

const cell = (margins: string): string => `<w:tc><w:tcPr><w:tcMar>${margins}</w:tcMar></w:tcPr><w:p/></w:tc>`;

const styles = (margins: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}">` +
    `<w:style w:type="table" w:styleId="MyOwnTableStyle"><w:tblPr><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr></w:style></w:styles>`;

const ctx = (parts: Record<string, string>): ValidationContext => ValidationContext.fromEntries(parts);
const findings = async (parts: Record<string, string>): Promise<readonly ValidationIssue[]> =>
    await wordCellMarginOrderCheck.run(ctx(parts));

describe("word-cell-margin-order", () => {
    it("reports w:bottom after w:right inside w:tblCellMar in a document story", async () => {
        const issues = await findings({ "word/document.xml": document(table(misordered)) });
        expect(issues.map((issue) => [issue.severity, issue.code, issue.path])).toEqual([
            ["warning", "word-cell-margin-order", "word/document.xml"],
        ]);
        expect(issues[0]?.message).toContain("w:bottom");
        expect(issues[0]?.message).toContain("top, start, left, bottom, end, right");
    });

    it("reports the same misorder inside w:tcPr/w:tcMar", async () => {
        const issues = await findings({
            "word/document.xml": document(
                `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr>${cell(misordered)}</w:tr></w:tbl>`,
            ),
        });
        expect(issues.map((issue) => [issue.code, issue.path])).toEqual([["word-cell-margin-order", "word/document.xml"]]);
        expect(issues[0]?.message).toContain("w:tcMar");
    });

    it("reports the misorder in word/styles.xml", async () => {
        const issues = await findings({ "word/styles.xml": styles(misordered) });
        expect(issues.map((issue) => [issue.code, issue.path])).toEqual([["word-cell-margin-order", "word/styles.xml"]]);
    });

    it("leaves a correctly ordered margin set accepted", async () => {
        expect(await findings({ "word/document.xml": document(table(legal)) })).toEqual([]);
    });

    it("leaves a margin set with children omitted accepted", async () => {
        const partial = `<w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>`;
        expect(await findings({ "word/document.xml": document(table(partial)) })).toEqual([]);
    });

    it("reports every occurrence separately instead of collapsing them", async () => {
        const issues = await findings({ "word/document.xml": document(table(misordered) + table(misordered)) });
        expect(issues).toHaveLength(2);
    });

    it("ignores a margin element whose parent cannot carry table margins", async () => {
        const orphan = `<w:tblCellMar>${misordered}</w:tblCellMar>`;
        expect(await findings({ "word/document.xml": document(orphan) })).toEqual([]);
    });

    it("matches the ISO Strict WordprocessingML namespace too", async () => {
        const issues = await findings({ "word/document.xml": document(table(misordered), WS) });
        expect(issues.map((issue) => issue.path)).toEqual(["word/document.xml"]);
    });

    it("skips a part that does not parse and keeps scanning the rest", async () => {
        const issues = await findings({
            "word/document.xml": `<w:document xmlns:w="${W}"><w:body>`,
            "word/styles.xml": styles(misordered),
        });
        expect(issues.map((issue) => issue.path)).toEqual(["word/styles.xml"]);
    });
});
