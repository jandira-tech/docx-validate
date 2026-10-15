/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word-valid yes/no walk. Every assertion calls the shipped `validate()`
 * entry. `xsdWalk` is the engine trace attached after `mergeResults`:
 * `started` means `validateAgainstXsd` ran, `parts` are the relative paths
 * actually handed to libxml2, `eligible` is how many a full walk would hand
 * over. A missing trace fails the assertion.
 */
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import * as xsdMod from "../src/lib/xsd-validator";
import { validate, type ValidateRunResult } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const BROKEN = path.join(HERE, "fixtures", "broken");
const WORKING = path.join(HERE, "fixtures", "working");
const INVALID = path.join(REPO, "word-invalid-fixtures");

const PGSZ = path.join(WORKING, "word-tolerated-misplaced-pgsz.docx");
const STRICT = path.join(WORKING, "word-clean-strict01.docx");
const MEDIA = path.join(WORKING, "word-tolerated-broken-media-rel.docx");
const CELLSPACING = path.join(BROKEN, "word-rejects-table-cellspacing.docx");
const EMPTY_RELS = path.join(INVALID, "OPEN_ERROR/tests__fixtures__broken__sample-document.empty-rels.docx");
const HEADER_XSD = path.join(INVALID, "OPEN_ERROR/tests__fixtures__broken__header-with-rid-no-sidecar.docx");
const MATH_BODY = path.join(INVALID, "OPEN_ERROR/tests__fixtures__math-body-level.docx");
const CUSTOM_XML = path.join(INVALID, "UNREADABLE_CONTENT/tests__fixtures__external__superdoc__layout-engine__shapes__vectors.docx");

const CAP_PART = "word/aa-cap.xml";
// libxml2 reports the seven missing pgMar attributes before tblCellSpacing.
// The first three strings are tolerated. The blocking error is later in the
// same part, so a cap of three must not hide it.
const CAP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sectPr>
      <w:pgSz/>
      <w:pgMar/>
      <w:cols/>
      <w:tblCellSpacing w:w="10" w:type="dxa"/>
    </w:sectPr>
  </w:body>
</w:document>
`;
const TAIL_PART = "word/zz-tail.xml";
const TAIL_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:body>
</w:document>
`;

const SPRE_INSIDE_PARAGRAPH = `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:sPre><m:sPrePr><m:ctrlPr><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/></m:ctrlPr></m:sPrePr><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup><m:e><m:r><m:t>Z</m:t></m:r></m:e></m:sPre></m:oMath></m:oMathPara>`;

interface XsdWalk {
    started: boolean;
    parts: readonly string[];
    eligible: number;
}

const norm = (p: string): string => p.split(/[\\/]/).join("/");

const walkOf = (result: ValidateRunResult): XsdWalk => {
    const walk = (result as ValidateRunResult & { xsdWalk?: XsdWalk }).xsdWalk;
    expect(walk, "validate() must attach xsdWalk after mergeResults").toBeDefined();
    return walk as XsdWalk;
};

const statsOf = (): { statusOnlyPasses: number; messageCollections: number } | undefined =>
    (xsdMod as { xsdValidateStats?: { statusOnlyPasses: number; messageCollections: number } }).xsdValidateStats;

const errorCodes = (result: ValidateRunResult): string[] =>
    result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "");

const withCapAndTail = async (src: string, dir: string): Promise<string> => {
    const zip = await JSZip.loadAsync(await readFile(src));
    zip.file(CAP_PART, CAP_XML);
    zip.file(TAIL_PART, TAIL_XML);
    const out = path.join(dir, "cap-then-block.docx");
    await writeFile(out, await zip.generateAsync({ type: "nodebuffer" }));
    return out;
};

const withTailPart = async (src: string, dir: string): Promise<string> => {
    const zip = await JSZip.loadAsync(await readFile(src));
    zip.file(TAIL_PART, TAIL_XML);
    const out = path.join(dir, "with-tail.docx");
    await writeFile(out, await zip.generateAsync({ type: "nodebuffer" }));
    return out;
};

describe("word-valid XSD yes/no walk", { timeout: 120_000 }, () => {
    it("does not downgrade an XSD engine failure to a Word-tolerated warning", async () => {
        const factory = vi.spyOn(xsdMod, "createXsdValidator").mockResolvedValue({
            validate: async () => [
                { severity: "error", code: "xsd-validation-failed", message: "XSD engine failed: invalid input or internal error" },
            ],
        });
        try {
            const result = await validate(PGSZ, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.severity === "error" && issue.message.includes("XSD engine failed:"))).toBe(true);
        } finally {
            factory.mockRestore();
        }
    });

    it("returns no on a blocking non-XSD defect before XSD starts", async () => {
        const result = await validate(EMPTY_RELS, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result).some((code) => code.startsWith("comment-thread"))).toBe(true);
        const walk = walkOf(result);
        expect(walk.started).toBe(false);
        expect(walk.parts).toEqual([]);
    });

    it("still starts XSD when the only blocking code is xsd-error", async () => {
        const result = await validate(HEADER_XSD, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain("xsd-error");
        const walk = walkOf(result);
        expect(walk.started).toBe(true);
        expect(walk.parts.length).toBeGreaterThan(0);
        expect(walk.parts.length).toBeLessThan(walk.eligible);
    });

    it("stops the word-valid XSD walk before a part that sorts after the blocking error", async () => {
        await withTempDir(async (dir) => {
            const docxPath = await withTailPart(CELLSPACING, dir);
            const before = statsOf()?.messageCollections ?? 0;
            const result = await validate(docxPath, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result)).toContain("xsd-error");
            const walk = walkOf(result);
            expect(walk.started).toBe(true);
            expect(walk.parts.map(norm)).not.toContain(TAIL_PART);
            expect(walk.parts.length).toBeGreaterThan(0);
            expect(walk.parts.length).toBeLessThan(walk.eligible);
            const stats = statsOf();
            expect(stats, "shipped validator exports xsdValidateStats").toBeDefined();
            expect(stats?.messageCollections ?? 0).toBeGreaterThan(before);
        });
    });

    it("surfaces a Word-blocking xsd-error that comes after three tolerated errors in the same part", async () => {
        await withTempDir(async (dir) => {
            const docxPath = await withCapAndTail(PGSZ, dir);
            const result = await validate(docxPath, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const blocking = result.issues.filter(
                (issue) =>
                    issue.severity === "error" &&
                    issue.code === "xsd-error" &&
                    issue.message.includes("tblCellSpacing") &&
                    issue.message.includes("not expected"),
            );
            expect(blocking.length).toBeGreaterThan(0);
            const walk = walkOf(result);
            expect(walk.started).toBe(true);
            expect(walk.parts.map(norm)).toContain(CAP_PART);
            expect(walk.parts.map(norm)).not.toContain(TAIL_PART);
            expect(walk.parts.length).toBeGreaterThan(0);
            expect(walk.parts.length).toBeLessThan(walk.eligible);
        });
    });

    it("lenient still walks every eligible part and still returns xsd-error messages", async () => {
        await withTempDir(async (dir) => {
            const docxPath = await withTailPart(CELLSPACING, dir);
            const result = await validate(docxPath, { profile: "lenient" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.code === "xsd-error" && issue.message.length > 0)).toBe(true);
            const walk = walkOf(result);
            expect(walk.started).toBe(true);
            expect(walk.parts.map(norm)).toContain(TAIL_PART);
            expect(walk.parts.length).toBe(walk.eligible);
            expect(walk.eligible).toBeGreaterThan(1);
        });
    });

    it("strict still walks every eligible part and still returns xsd-error messages", async () => {
        await withTempDir(async (dir) => {
            const docxPath = await withTailPart(CELLSPACING, dir);
            const result = await validate(docxPath, { profile: "strict" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.code === "xsd-error" && issue.message.length > 0)).toBe(true);
            const walk = walkOf(result);
            expect(walk.started).toBe(true);
            expect(walk.parts.map(norm)).toContain(TAIL_PART);
            expect(walk.parts.length).toBe(walk.eligible);
        });
    });

    it("a word-valid pass still starts XSD, finishes the walk, and counts a status-only pass", async () => {
        const before = statsOf()?.statusOnlyPasses ?? 0;
        const result = await validate(PGSZ, { profile: "word-valid" });
        expect(result.valid).toBe(true);
        const walk = walkOf(result);
        expect(walk.started).toBe(true);
        expect(walk.eligible).toBeGreaterThan(0);
        expect(walk.parts.length).toBe(walk.eligible);
        const stats = statsOf();
        expect(stats, "shipped validator exports xsdValidateStats").toBeDefined();
        expect(stats?.statusOnlyPasses ?? 0).toBeGreaterThan(before);
        const zip = await JSZip.loadAsync(await readFile(PGSZ));
        expect(zip.file("docProps/core.xml")).toBeTruthy();
        for (const needle of ["hyphenationZone", "purl.org/dc/terms", "dublincore.org", "Invalid XSD schema"]) {
            expect(
                result.issues.some((issue) => issue.severity === "error" && issue.message.includes(needle)),
                needle,
            ).toBe(false);
        }
    });

    it("a schema-load failure stays an error under word-valid", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(PGSZ, { profile: "word-valid", schemasDir: dir });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.severity === "error" && issue.message.includes("Schema load failed"))).toBe(true);
        });
    });

    it("strict OOXML records xsd-strict-skipped, does not send the strict main part to the engine, and stays valid", async () => {
        // word-clean-strict01.docx mixes purl.oclc.org strict parts (word/document.xml)
        // with transitional package parts ([Content_Types].xml, .rels, core.xml).
        // The walk still enters, skips each strict root, and may send the transitional parts.
        const result = await validate(STRICT, { profile: "word-valid" });
        expect(result.valid).toBe(true);
        expect(result.issues.some((issue) => issue.code === "xsd-strict-skipped")).toBe(true);
        const walk = walkOf(result);
        expect(walk.started).toBe(true);
        expect(walk.parts.map(norm)).not.toContain("word/document.xml");
        expect(walk.parts.length).toBe(walk.eligible);
    });

    it("body-level m:sPre is word-math-spre-body without needing XSD", async () => {
        const result = await validate(MATH_BODY, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain("word-math-spre-body");
        expect(walkOf(result).started).toBe(false);
    });

    it("m:sPre inside a w:p does not produce word-math-spre-body", async () => {
        await withTempDir(async (dir) => {
            const zip = await JSZip.loadAsync(await readFile(PGSZ));
            const entry = zip.file("word/document.xml");
            expect(entry).toBeTruthy();
            const xml = await entry!.async("string");
            const at = xml.indexOf("</w:p>");
            expect(at).toBeGreaterThan(0);
            zip.file("word/document.xml", `${xml.slice(0, at)}${SPRE_INSIDE_PARAGRAPH}${xml.slice(at)}`);
            const docxPath = path.join(dir, "spre-in-paragraph.docx");
            await writeFile(docxPath, await zip.generateAsync({ type: "nodebuffer" }));
            const result = await validate(docxPath, { profile: "word-valid" });
            expect(errorCodes(result)).not.toContain("word-math-spre-body");
        });
    });

    it("a broken media relationship stays a non-failure and still walks XSD", async () => {
        const result = await validate(MEDIA, { profile: "word-valid" });
        expect(result.valid).toBe(true);
        const walk = walkOf(result);
        expect(walk.started).toBe(true);
        expect(walk.parts.length).toBe(walk.eligible);
    });

    it("a broken customXml relationship is a no before XSD starts", async () => {
        const result = await validate(CUSTOM_XML, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain("rels-broken");
        const walk = walkOf(result);
        expect(walk.started).toBe(false);
        expect(walk.parts).toEqual([]);
    });

    it("an OLE package fails as package-open-failed without an XSD walk", async () => {
        await withTempDir(async (dir) => {
            const bad = path.join(dir, "encrypted.docx");
            await writeFile(bad, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00, 0x00, 0x00]));
            const result = await validate(bad, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result)).toContain("package-open-failed");
            const walk = walkOf(result);
            expect(walk.started).toBe(false);
            expect(walk.parts).toEqual([]);
        });
    });
});
