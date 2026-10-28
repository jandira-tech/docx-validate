/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Integration tests for the shading colour value-space check (Task 8 family
 * 8c; lane issues #210, #215, #220, #221, #225, #229, #238, #242).
 *
 * (a) A synthetic single-defect package: the same minimal table, once with
 *     `w:fill=""` (defect) and once with the optional attribute omitted
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
import { wordShadingColorCheck } from "../src/scripts/office/checks/shading-color";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

/** One minimal table row whose cell carries the given `w:shd`. */
const tableWithShading = (shd: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
    `<w:tr w14:paraId="00000002" w14:textId="00000002"><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/>${shd}</w:tcPr>` +
    `<w:p w14:paraId="00000003" w14:textId="00000003"/></w:tc></w:tr></w:tbl>`;

const partsWith = (shd: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${tableWithShading(shd)}<w:sectPr>`),
});

const validateBytes = async (bytes: Uint8Array, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });

const advisories = (issues: readonly ValidationIssue[]): ValidationIssue[] => issues.filter((issue) => issue.code === "word-shading-color");

describe("shading-color integration: synthetic single-defect package", { timeout: 120000 }, () => {
    it("the control (colour omitted) is clean under every profile", async () => {
        const control = await buildDocx(partsWith(`<w:shd w:val="clear"/>`));
        for (const profile of ["lenient", "word-valid", "strict"] as const) {
            const result = await validateBytes(control, profile);
            expect(advisories(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        }
    });

    it("word-valid stays valid and carries the warning; strict fails; lenient is silent", async () => {
        const defect = await buildDocx(partsWith(`<w:shd w:val="clear" w:fill=""/>`));

        const wordValid = await validateBytes(defect, "word-valid");
        expect(wordValid.valid).toBe(true);
        expect(advisories(wordValid.issues)).toHaveLength(1);
        expect(advisories(wordValid.issues)[0]).toMatchObject({
            severity: "warning",
            code: "word-shading-color",
            path: "word/document.xml",
        });
        expect(advisories(wordValid.issues)[0]!.message).toContain(`w:fill=""`);

        const strict = await validateBytes(defect, "strict");
        expect(strict.valid).toBe(false);
        expect(advisories(strict.issues)).toHaveLength(1);

        const lenient = await validateBytes(defect, "lenient");
        expect(advisories(lenient.issues)).toEqual([]);
    });
});

/** `parts` with its main document stored as `main`, in place of the part there, and no main-document sidecar (#1122). */
const mainStoredAs = (parts: Readonly<Record<string, string>>, main: string): Record<string, string> => {
    const { "word/document.xml": document, "word/_rels/document.xml.rels": _sidecar, ...rest } = parts;
    return {
        ...rest,
        "[Content_Types].xml": rest["[Content_Types].xml"]!.replace(new RegExp(`<Override PartName="/${main}"[^>]*/>`), "").replace(
            'PartName="/word/document.xml"',
            `PartName="/${main}"`,
        ),
        "_rels/.rels": rest["_rels/.rels"]!.replace('Target="word/document.xml"', `Target="${main}"`),
        [main]: document!,
    };
};

// #1122 (review 1117 F1): a main document stored as word/styles.xml or
// word/numbering.xml is a read story part and a styles/numbering part; its
// finding is reported once, in a full run as in the check.
describe("shading-color integration: a main document stored under a styles or numbering name (#1122)", { timeout: 120000 }, () => {
    it.each(["word/styles.xml", "word/numbering.xml"])(
        "a full word-valid run reports the main document stored as %s once",
        async (main) => {
            const bytes = await buildDocx(mainStoredAs(partsWith(`<w:shd w:val="clear" w:fill=""/>`), main));
            expect(advisories((await validateBytes(bytes, "word-valid")).issues).map((issue) => issue.path)).toEqual([main]);
        },
    );
});

describe("shading-color integration: jubarte 0.13.0 lane specimens", () => {
    /** k06_1 (#221) and k06_2 (#215): 60 empty `w:fill` per document part. */
    const SPECIMENS: readonly { file: string; issue: number; documentFindings: number }[] = [
        { file: "k06_1.redline.docx", issue: 221, documentFindings: 60 },
        { file: "k06_2.redline.docx", issue: 215, documentFindings: 60 },
    ];

    for (const { file, issue, documentFindings } of SPECIMENS) {
        it(`${file} (#${issue}): ${documentFindings} findings in word/document.xml, none in word/styles.xml`, async () => {
            const { context } = await contextFromArchive(new Uint8Array(readFileSync(path.join(LANE, file))));
            const issues = await wordShadingColorCheck.run(context);
            expect(issues).toHaveLength(documentFindings);
            for (const finding of issues) {
                expect(finding.code).toBe("word-shading-color");
                expect(finding.severity).toBe("warning");
                expect(finding.path).toBe("word/document.xml");
                expect(finding.message).toContain(`w:fill=""`);
            }
            // k06_2 also carries one fully valid w:shd in styles.xml
            // (w:color="auto" w:fill="E1DFDD"): it must stay unreported.
            expect(issues.some((finding) => finding.path === "word/styles.xml")).toBe(false);
        });
    }
});
