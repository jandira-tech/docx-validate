/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-table-properties-required` against real packages: one synthetic
 * single-defect specimen built from the clean minimal package, and the jubarte
 * 0.13.0 lane redlines the check was written for (#212, #217, #218).
 *
 * All 24 lane redlines were opened and exported by Microsoft Word, so a missing
 * `w:tblPr` is an advisory — `word-valid` stays valid and only `strict`
 * rejects. That is a historical observation bound to the sha256 in
 * `tests/lane/jubarte-0.13.0/manifest.csv`, not a fresh Word verification.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { wordTablePropertiesRequiredCheck } from "../src/scripts/office/checks/table-properties-required";
import { validate } from "../src/scripts/office/validate";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const LANE = path.join(path.dirname(fileURLToPath(import.meta.url)), "lane", "jubarte-0.13.0");

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const ROW = `<w:tr w14:paraId="00000003" w14:textId="00000003"><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p w14:paraId="00000004" w14:textId="00000004"/></w:tc></w:tr>`;
const GRID = `<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>`;

/** `CT_Tbl` with both required children present — the control. */
const legalTable = `<w:tbl><w:tblPr/>${GRID}${ROW}</w:tbl>`;
/** The lane shape: `w:tblGrid` first, no `w:tblPr`. */
const gridOnlyTable = `<w:tbl>${GRID}${ROW}</w:tbl>`;

/** The minimal package plus a header part, so the defect sits where the lane puts it. */
const packageWithHeaderTable = (table: string): Readonly<Record<string, string>> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    ),
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="rId2" Type="${R}/header" Target="header1.xml"/></Relationships>`,
    ),
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:mc=`, `xmlns:r="${R}" xmlns:mc=`).replace(
        "<w:pgSz",
        `<w:headerReference w:type="default" r:id="rId2"/><w:pgSz`,
    ),
    "word/header1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:w14="${W14}">${table}</w:hdr>`,
});

const buildWithHeaderTable = (table: string) => buildDocx(packageWithHeaderTable(table));

const validateBytes = async (bytes: Uint8Array, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });

const advisories = (issues: readonly { code?: string }[]) => issues.filter((issue) => issue.code === "word-table-properties-required");

const laneFindings = async (file: string) => {
    const { context, openError } = await contextFromArchive(readFileSync(path.join(LANE, file)));
    expect(openError).toBeUndefined();
    return wordTablePropertiesRequiredCheck.run(context);
};

describe("word-table-properties-required: single-defect package", { timeout: 180000 }, () => {
    it("keeps word-valid valid while carrying the advisory", async () => {
        const result = await validateBytes(await buildWithHeaderTable(gridOnlyTable), "word-valid");
        expect(advisories(result.issues)).toEqual([expect.objectContaining({ severity: "warning", path: "word/header1.xml" })]);
        expect(result.valid).toBe(true);
    });

    it("rejects the same package under strict, and says why", async () => {
        const result = await validateBytes(await buildWithHeaderTable(gridOnlyTable), "strict");
        expect(result.valid).toBe(false);
        expect(advisories(result.issues)).toHaveLength(1);
        // The rejection has to be this defect's schema error, not fixture noise.
        const schemaError = result.issues.find((issue) => issue.code === "xsd-error" && issue.path === "word/header1.xml");
        expect(schemaError?.message).toContain("tblGrid");
    });

    it("does not report the advisory under lenient", async () => {
        const result = await validateBytes(await buildWithHeaderTable(gridOnlyTable), "lenient");
        expect(advisories(result.issues)).toEqual([]);
    });

    it("keeps the well-formed header table clean under every profile", async () => {
        const bytes = await buildWithHeaderTable(legalTable);
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            const result = await validateBytes(bytes, profile);
            expect(advisories(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        }
    });
});

/**
 * Counts recounted from the lane XML and cross-checked against the filed
 * issues: each specimen has exactly one `w:tbl` in the named header part, with
 * no `w:tblPr` anywhere in it, so one finding per package.
 */
const LANE_COUNTS: readonly { file: string; part: string; findings: number; issue: string }[] = [
    { file: "k11_1.redline.docx", part: "word/header2.xml", findings: 1, issue: "#217" },
    { file: "k11_2.redline.docx", part: "word/header2.xml", findings: 1, issue: "#218" },
    { file: "k13_1.redline.docx", part: "word/header4.xml", findings: 1, issue: "#212" },
];

describe("word-table-properties-required: jubarte 0.13.0 lane specimens", { timeout: 180000 }, () => {
    for (const { file, part, findings, issue } of LANE_COUNTS) {
        it(`${file} reports ${findings} finding in ${part} (${issue})`, async () => {
            const reported = await laneFindings(file);
            expect(reported.map((issue_) => issue_.path)).toEqual([part]);
            expect(reported).toHaveLength(findings);
            expect(reported[0]?.message).toContain("w:tblPr");
        });
    }

    it("leaves a specimen whose header tables are well formed with no finding", async () => {
        // k11_2 carries the defect only in header2.xml; the body tables of the
        // same package are complete, which is the "remains accepted" control.
        const reported = await laneFindings("k11_2.redline.docx");
        expect(reported.every((issue_) => issue_.path === "word/header2.xml")).toBe(true);
    });
});
