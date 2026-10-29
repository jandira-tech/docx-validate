/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Plan Task 7, final step: complete Word-valid validation does not compile
 * the XSD bundle. Every Word-blocking structure is named by a structural
 * check (`WORD_RULE_CHECKS` and the non-XSD validators), so word-valid never
 * asks libxml2 for a schema. Lenient and strict still walk the schemas.
 *
 * Replaces the transitional "word-valid XSD yes/no walk" suite: its
 * non-XSD cases are kept here, its walk-stopping cases have nothing left to
 * test.
 */
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import * as xsdMod from "../src/lib/xsd-validator";
import { validate, type ValidateRunResult } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const BROKEN = path.join(HERE, "fixtures", "broken");
const WORKING = path.join(HERE, "fixtures", "working");
const EXTERNAL = path.join(HERE, "fixtures", "external");
const INVALID = path.join(REPO, "word-invalid-fixtures");
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");
const PGSZ = path.join(WORKING, "word-tolerated-misplaced-pgsz.docx");

// [fixture, word-valid verdict]. Rejections are Word OPEN_ERROR or
// unreadable-content specimens; the controls are files Word opens cleanly
// despite schema defects.
const CASES: readonly (readonly [string, boolean])[] = [
    [path.join(BROKEN, "word-rejects-table-cellspacing.docx"), false],
    [path.join(BROKEN, "word-rejects-misplaced-run.docx"), false],
    [path.join(INVALID, "OPEN_ERROR/tests__fixtures__broken__header-with-rid-no-sidecar.docx"), false],
    [path.join(EXTERNAL, "open-xml-sdk/UnknownElement.docx"), false],
    [path.join(EXTERNAL, "open-xml-sdk/5Errors.docx"), false],
    [path.join(EXTERNAL, "superdoc/super-editor/hyperlink-exported.docx"), false],
    [path.join(EXTERNAL, "superdoc/super-editor/invalid-underline.docx"), false],
    [path.join(EXTERNAL, "docx-parser-converter/tagged_tests/table_tests_v2.docx"), false],
    [PGSZ, true],
    [path.join(WORKING, "word-tolerated-duplicate-ppr.docx"), true],
    [path.join(WORKING, "word-clean-strict01.docx"), true],
    [path.join(WORKING, "word-tolerated-broken-media-rel.docx"), true],
    [path.join(LANE, "k07_1.redline.docx"), true],
];

const SPRE_INSIDE_PARAGRAPH = `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:sPre><m:sPrePr><m:ctrlPr><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/></m:ctrlPr></m:sPrePr><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup><m:e><m:r><m:t>Z</m:t></m:r></m:e></m:sPre></m:oMath></m:oMathPara>`;

const errorCodes = (result: ValidateRunResult): string[] =>
    result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "");

const expectNoXsd = (result: ValidateRunResult): void => {
    expect(result.xsdWalk?.started ?? false).toBe(false);
    expect(result.issues.filter((issue) => issue.code?.startsWith("xsd-"))).toEqual([]);
};

afterEach(() => {
    vi.restoreAllMocks();
});

describe("word-valid never compiles the XSD bundle", { timeout: 120_000 }, () => {
    for (const [fixture, expected] of CASES) {
        it(`${path.relative(HERE, fixture)} is ${expected ? "valid" : "invalid"} without an XSD engine`, async () => {
            const factory = vi.spyOn(xsdMod, "createXsdValidator");
            const result = await validate(fixture, { profile: "word-valid" });
            expect(factory).not.toHaveBeenCalled();
            expectNoXsd(result);
            expect(result.valid, JSON.stringify(result.issues.filter((issue) => issue.severity === "error"))).toBe(expected);
        });
    }

    it("needs no schema directory at all", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(PGSZ, { profile: "word-valid", schemasDir: dir });
            expectNoXsd(result);
            expect(result.valid).toBe(true);
        });
    });

    for (const profile of ["lenient", "strict"] as const) {
        it(`${profile} still compiles the schemas and walks every eligible part`, async () => {
            const factory = vi.spyOn(xsdMod, "createXsdValidator");
            const result = await validate(path.join(BROKEN, "word-rejects-table-cellspacing.docx"), { profile });
            expect(factory).toHaveBeenCalled();
            expect(result.xsdWalk?.started).toBe(true);
            expect(result.xsdWalk?.parts.length).toBe(result.xsdWalk?.eligible);
            expect(result.issues.some((issue) => issue.code === "xsd-error" && issue.message.length > 0)).toBe(true);
            expect(result.valid).toBe(false);
        });
    }
});

describe("word-valid keeps its non-XSD rejections", { timeout: 120_000 }, () => {
    it("a broken comment thread is a comment-thread code", async () => {
        const result = await validate(path.join(INVALID, "OPEN_ERROR/tests__fixtures__broken__sample-document.empty-rels.docx"), {
            profile: "word-valid",
        });
        expect(result.valid).toBe(false);
        expect(errorCodes(result).some((code) => code.startsWith("comment-thread"))).toBe(true);
        expectNoXsd(result);
    });

    it("body-level m:sPre is word-math-spre-body", async () => {
        const result = await validate(path.join(INVALID, "OPEN_ERROR/tests__fixtures__math-body-level.docx"), { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain("word-math-spre-body");
        expectNoXsd(result);
    });

    it("m:sPre inside a w:p does not produce word-math-spre-body", async () => {
        await withTempDir(async (dir) => {
            const zip = await JSZip.loadAsync(await readFile(PGSZ));
            const xml = await zip.file("word/document.xml")!.async("string");
            const at = xml.indexOf("</w:p>");
            expect(at).toBeGreaterThan(0);
            zip.file("word/document.xml", `${xml.slice(0, at)}${SPRE_INSIDE_PARAGRAPH}${xml.slice(at)}`);
            const docxPath = path.join(dir, "spre-in-paragraph.docx");
            await writeFile(docxPath, await zip.generateAsync({ type: "nodebuffer" }));
            const result = await validate(docxPath, { profile: "word-valid" });
            expect(errorCodes(result)).not.toContain("word-math-spre-body");
        });
    });

    it("a broken customXml relationship is rels-broken", async () => {
        const file = path.join(INVALID, "UNREADABLE_CONTENT/tests__fixtures__external__superdoc__layout-engine__shapes__vectors.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain("rels-broken");
        expectNoXsd(result);
    });

    it("an OLE package fails as package-open-failed", async () => {
        await withTempDir(async (dir) => {
            const bad = path.join(dir, "encrypted.docx");
            await writeFile(bad, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00, 0x00, 0x00]));
            const result = await validate(bad, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result)).toContain("package-open-failed");
            expectNoXsd(result);
        });
    });
});

describe("malformed markup compatibility stays diagnosed without XSD", { timeout: 120_000 }, () => {
    // An AlternateContent with no Choice.
    const malformed = () =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:sectPr>",
                `<mc:AlternateContent><mc:Fallback><w:p w14:paraId="00000002" w14:textId="00000002"/></mc:Fallback></mc:AlternateContent><w:sectPr>`,
            ),
        });

    it("word-valid rejects it as mce-malformed", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "mce-malformed.docx");
            await writeFile(file, await malformed());
            const result = await validate(file, { profile: "word-valid" });
            expectNoXsd(result);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
            );
            expect(result.valid).toBe(false);
        });
    });

    it("word-valid accepts a requirement Word understands, which strict's schema view does not", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "must-understand-w14.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace('mc:Ignorable="w14"', 'mc:MustUnderstand="w14"'),
                }),
            );
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.map((issue) => issue.code)).not.toContain("mce-must-understand");
            expect(wordValid.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === "mce-must-understand")).toEqual([
                expect.objectContaining({ severity: "error", path: "word/document.xml" }),
            ]);
        });
    });

    it("strict reports the same finding once", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "mce-malformed.docx");
            await writeFile(file, await malformed());
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.filter((issue) => issue.code === "mce-malformed")).toHaveLength(1);
            expect(result.valid).toBe(false);
        });
    });
});
