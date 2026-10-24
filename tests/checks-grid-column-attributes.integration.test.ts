/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Integration tests for the grid-column attribute check (Task 8 family 8d;
 * lane issues #209, #219, #224, #226, #233, #240).
 *
 * (a) A synthetic single-defect package: the same minimal table, once with
 *     `w:type="dxa"` on the grid column (defect) and once with `w:w` alone
 *     (control). Word-valid stays valid and carries the advisory, strict
 *     fails, lenient never reports the code.
 * (b) The jubarte 0.13.0 lane specimens the issues were filed from. All 24
 *     lane redlines opened and exported cleanly in Microsoft Word during the
 *     lane run, so the defect is schema-invalid but historically
 *     Word-tolerated; the counts are bound to the file hashes, not a new Word
 *     verification.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile, ValidationIssue } from "../src/lib/types";
import { wordGridColumnAttributesCheck } from "../src/scripts/office/checks/grid-column-attributes";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

/** One minimal table whose grid column carries the given attributes. */
const tableWithGridCol = (gridCol: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    `<w:tblGrid>${gridCol}</w:tblGrid>` +
    `<w:tr w14:paraId="00000002" w14:textId="00000002"><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr>` +
    `<w:p w14:paraId="00000003" w14:textId="00000003"/></w:tc></w:tr></w:tbl>`;

const partsWith = (gridCol: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${tableWithGridCol(gridCol)}<w:sectPr>`),
});

const validateBytes = async (bytes: Uint8Array, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });

const advisories = (issues: readonly ValidationIssue[]): ValidationIssue[] =>
    issues.filter((issue) => issue.code === "word-grid-column-attributes");

describe("grid-column-attributes integration: synthetic single-defect package", { timeout: 120000 }, () => {
    it("the control (w:w alone) is clean under every profile", async () => {
        const control = await buildDocx(partsWith(`<w:gridCol w:w="100"/>`));
        for (const profile of ["lenient", "word-valid", "strict"] as const) {
            const result = await validateBytes(control, profile);
            expect(advisories(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        }
    });

    it("word-valid stays valid and carries the warning; strict fails; lenient is silent", async () => {
        const defect = await buildDocx(partsWith(`<w:gridCol w:w="100" w:type="dxa"/>`));

        const wordValid = await validateBytes(defect, "word-valid");
        expect(wordValid.valid).toBe(true);
        expect(advisories(wordValid.issues)).toHaveLength(1);
        expect(advisories(wordValid.issues)[0]).toMatchObject({
            severity: "warning",
            code: "word-grid-column-attributes",
            path: "word/document.xml",
        });
        // The width stays visible in the finding, so no information is lost.
        expect(advisories(wordValid.issues)[0]!.message).toContain(`w:type="dxa"`);
        expect(advisories(wordValid.issues)[0]!.message).toContain(`w:w="100"`);

        const strict = await validateBytes(defect, "strict");
        expect(strict.valid).toBe(false);
        expect(advisories(strict.issues)).toHaveLength(1);

        const lenient = await validateBytes(defect, "lenient");
        expect(advisories(lenient.issues)).toEqual([]);
    });
});

describe("grid-column-attributes integration: jubarte 0.13.0 lane specimens", () => {
    /**
     * k07_1 (#209) and k07_2 (#226): 38 grid columns carry `w:type` in
     * word/document.xml. In k07_2 seven of the 38 sit inside
     * `w:tblGridChange/w:tblGrid`, which uses the same `CT_TblGridCol` type.
     */
    const SPECIMENS: readonly { file: string; issue: number; findings: number }[] = [
        { file: "k07_1.redline.docx", issue: 209, findings: 38 },
        { file: "k07_2.redline.docx", issue: 226, findings: 38 },
    ];

    for (const { file, issue, findings: count } of SPECIMENS) {
        it(`${file} (#${issue}): ${count} findings in word/document.xml`, async () => {
            const { context } = await contextFromArchive(new Uint8Array(readFileSync(path.join(LANE, file))));
            const issues = await wordGridColumnAttributesCheck.run(context);
            expect(issues).toHaveLength(count);
            for (const finding of issues) {
                expect(finding.code).toBe("word-grid-column-attributes");
                expect(finding.severity).toBe("warning");
                expect(finding.path).toBe("word/document.xml");
                expect(finding.message).toContain("w:type");
                expect(finding.message).toContain(`w:w="`); // the width stays visible
            }
            // k07_2 also carries 14 clean grid columns (w:w only); they must
            // stay unreported, leaving exactly the count the issue states.
        });
    }
});
