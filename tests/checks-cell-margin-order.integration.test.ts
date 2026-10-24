/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-cell-margin-order` against real packages: one synthetic single-defect
 * specimen built from the clean minimal package, and the jubarte 0.13.0 lane
 * redlines the check was written for (#211, #216, #222, #223, #227, #228,
 * #231, #232, #234, #236, #241, #246).
 *
 * All 24 lane redlines were opened and exported by Microsoft Word, so every
 * finding here is an advisory — `word-valid` stays valid and only `strict`
 * rejects. That is a historical observation bound to the sha256 in
 * `tests/lane/jubarte-0.13.0/manifest.csv`, not a fresh Word verification.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { wordCellMarginOrderCheck } from "../src/scripts/office/checks/cell-margin-order";
import { validate } from "../src/scripts/office/validate";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const LANE = path.join(path.dirname(fileURLToPath(import.meta.url)), "lane", "jubarte-0.13.0");

/** `CT_TblCellMar` order Word writes (transitional: no `start`/`end`). */
const LEGAL = `<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>`;
const MISORDERED = `<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>`;

// w14:paraId / w14:textId keep the `paraid-missing-element` rule satisfied, so
// the margin order is the only difference between specimen and control.
const tableWith = (margins: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
    `<w:tr w14:paraId="00000003" w14:textId="00000003"><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr>` +
    `<w:p w14:paraId="00000004" w14:textId="00000004"/></w:tc></w:tr></w:tbl>` +
    `<w:p w14:paraId="00000002" w14:textId="00000002"/>`;

const withTable = (margins: string): Promise<Uint8Array> =>
    buildDocx({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${tableWith(margins)}<w:sectPr>`),
    });

const validateBytes = async (bytes: Uint8Array, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });

const advisories = (issues: readonly { code?: string }[]) => issues.filter((issue) => issue.code === "word-cell-margin-order");

const laneFindings = async (file: string) => {
    const { context, openError } = await contextFromArchive(readFileSync(path.join(LANE, file)));
    expect(openError).toBeUndefined();
    return wordCellMarginOrderCheck.run(context);
};

const countsByPart = (issues: readonly { path?: string }[]): Record<string, number> => {
    const counts: Record<string, number> = {};
    for (const issue of issues) {
        const key = issue.path ?? "";
        counts[key] = (counts[key] ?? 0) + 1;
    }
    return counts;
};

describe("word-cell-margin-order: single-defect package", { timeout: 180000 }, () => {
    it("keeps word-valid valid while carrying the advisory", async () => {
        const result = await validateBytes(await withTable(MISORDERED), "word-valid");
        expect(advisories(result.issues)).toEqual([expect.objectContaining({ severity: "warning", path: "word/document.xml" })]);
        expect(result.valid).toBe(true);
    });

    it("rejects the same package under strict, and says why", async () => {
        const result = await validateBytes(await withTable(MISORDERED), "strict");
        expect(result.valid).toBe(false);
        expect(advisories(result.issues).length).toBeGreaterThan(0);
        // The rejection has to be this defect's schema error, not fixture noise.
        const schemaError = result.issues.find((issue) => issue.code === "xsd-error" && issue.path === "word/document.xml");
        expect(schemaError?.message).toContain("}bottom'");
    });

    it("does not report the advisory under lenient", async () => {
        const result = await validateBytes(await withTable(MISORDERED), "lenient");
        expect(advisories(result.issues)).toEqual([]);
    });

    it("keeps the correctly ordered control clean under every profile", async () => {
        const bytes = await withTable(LEGAL);
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            const result = await validateBytes(bytes, profile);
            expect(advisories(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        }
    });
});

/**
 * Counts recounted from the lane XML and cross-checked against the filed
 * issues: 31 misordered `w:tblCellMar` in `word/document.xml` and 1 in
 * `word/styles.xml` per specimen (the 31 `w:tblPrEx` sets jubarte writes as
 * `top, bottom` are legal).
 */
const LANE_COUNTS: readonly { file: string; document: number; styles: number; issues: string }[] = [
    { file: "k08_1.redline.docx", document: 31, styles: 1, issues: "#222, #223" },
    { file: "k08_2.redline.docx", document: 31, styles: 1, issues: "#234, #246" },
    { file: "k09_1.redline.docx", document: 31, styles: 1, issues: "#241, #236" },
    { file: "k09_2.redline.docx", document: 31, styles: 1, issues: "#228, #227" },
];

describe("word-cell-margin-order: jubarte 0.13.0 lane specimens", { timeout: 180000 }, () => {
    for (const { file, document, styles, issues } of LANE_COUNTS) {
        it(`${file} reports ${document} in word/document.xml and ${styles} in word/styles.xml (${issues})`, async () => {
            const findings = await laneFindings(file);
            expect(countsByPart(findings)).toEqual({ "word/document.xml": document, "word/styles.xml": styles });
        });
    }
});
