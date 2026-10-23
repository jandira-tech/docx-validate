/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-table-properties-required`: `CT_Tbl` (ISO-IEC29500-4_2016/wml.xsd:2532)
 * is a sequence of optional range markup, a required `w:tblPr`, a required
 * `w:tblGrid`, then the rows. Word fills in table defaults when `w:tblPr` is
 * absent, so a header table written as `w:tbl > w:tblGrid > w:tr` opens fine and
 * the finding is an advisory.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { wordTablePropertiesRequiredCheck } from "../src/scripts/office/checks/table-properties-required";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const CELL = `<w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
const ROW = `<w:tr>${CELL}</w:tr>`;
const GRID = `<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>`;
/** `CT_TblPr` has only optional children, so an empty `w:tblPr` is valid. */
const TBL_PR = `<w:tblPr/>`;
const LEADING_MARKUP = `<w:bookmarkStart w:id="1" w:name="t"/><w:bookmarkEnd w:id="1"/>`;

const table = (...children: readonly string[]): string => `<w:tbl>${children.join("")}</w:tbl>`;
const body = (content: string, ns = W): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${ns}"><w:body>${content}</w:body></w:document>`;
const header = (content: string, ns = W): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${ns}">${content}</w:hdr>`;

const findings = async (parts: Record<string, string>): Promise<readonly ValidationIssue[]> =>
    await wordTablePropertiesRequiredCheck.run(ValidationContext.fromEntries(parts));

describe("word-table-properties-required", () => {
    it("reports a body table that starts at w:tblGrid, naming the missing w:tblPr", async () => {
        const issues = await findings({ "word/document.xml": body(table(GRID, ROW)) });
        expect(issues.map((issue) => [issue.severity, issue.code, issue.path])).toEqual([
            ["warning", "word-table-properties-required", "word/document.xml"],
        ]);
        expect(issues[0]?.message).toContain("w:tblPr");
        expect(issues[0]?.message).toContain("w:tblGrid");
    });

    it("reports a header table that starts at w:tblGrid", async () => {
        const issues = await findings({ "word/header1.xml": header(table(GRID, ROW)) });
        expect(issues.map((issue) => [issue.code, issue.path])).toEqual([["word-table-properties-required", "word/header1.xml"]]);
    });

    it("names the missing w:tblGrid", async () => {
        const issues = await findings({ "word/document.xml": body(table(TBL_PR, ROW)) });
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("w:tblGrid");
    });

    it("reports w:tblPr written after w:tblGrid", async () => {
        const issues = await findings({ "word/document.xml": body(table(GRID, TBL_PR, ROW)) });
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("w:tblPr");
    });

    it("reports a duplicate w:tblPr", async () => {
        const issues = await findings({ "word/document.xml": body(table(TBL_PR, TBL_PR, GRID, ROW)) });
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("w:tblPr");
    });

    it("reports a w:tblGrid written after the rows", async () => {
        const issues = await findings({ "word/document.xml": body(table(TBL_PR, ROW, GRID)) });
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("w:tblGrid");
    });

    it("reports a duplicate w:tblGrid", async () => {
        const issues = await findings({ "word/document.xml": body(table(TBL_PR, GRID, ROW, GRID)) });
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("w:tblGrid");
    });

    it("reports an empty w:tbl as missing both required children", async () => {
        const issues = await findings({ "word/document.xml": body(table()) });
        expect(issues).toHaveLength(2);
        expect(issues.map((issue) => issue.message.split(". ")[0])).toEqual(["w:tbl: w:tblPr is missing", "w:tbl: w:tblGrid is missing"]);
    });

    it("accepts a well-formed table in the body and in a header", async () => {
        const issues = await findings({
            "word/document.xml": body(table(TBL_PR, GRID, ROW)),
            "word/header1.xml": header(table(TBL_PR, GRID, ROW)),
        });
        expect(issues).toEqual([]);
    });

    it("accepts leading range markup before w:tblPr", async () => {
        const issues = await findings({ "word/document.xml": body(table(LEADING_MARKUP, TBL_PR, GRID, ROW)) });
        expect(issues).toEqual([]);
    });

    it("accepts several rows", async () => {
        const issues = await findings({ "word/document.xml": body(table(TBL_PR, GRID, ROW, ROW)) });
        expect(issues).toEqual([]);
    });

    it("ignores children from a foreign namespace", async () => {
        const foreign = `<w:tbl xmlns:x="urn:x"><x:hint/><w:tblPr/>${GRID}${ROW}</w:tbl>`;
        const issues = await findings({ "word/document.xml": body(foreign) });
        expect(issues).toEqual([]);
    });

    it("reports every defective table separately and checks nested tables", async () => {
        const nested = table(TBL_PR, GRID, `<w:tr><w:tc><w:tbl>${GRID}${ROW}</w:tbl></w:tc></w:tr>`);
        const issues = await findings({ "word/document.xml": body(table(GRID, ROW) + nested) });
        expect(issues).toHaveLength(2);
    });

    it("matches the ISO Strict WordprocessingML namespace too", async () => {
        const issues = await findings({ "word/document.xml": body(table(GRID, ROW), WS) });
        expect(issues.map((issue) => issue.path)).toEqual(["word/document.xml"]);
    });

    it("skips a part that does not parse without throwing", async () => {
        const issues = await findings({
            "word/document.xml": `<w:document xmlns:w="${W}"><w:body>`,
            "word/header1.xml": header(table(GRID, ROW)),
        });
        expect(issues.map((issue) => issue.path)).toEqual(["word/header1.xml"]);
    });
});
