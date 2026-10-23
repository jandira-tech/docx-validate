/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Integration tests for `textboxContentNamespaceCheck` (issue #244,
 * jubarte lane family k12): a synthetic single-defect package through the
 * full `validate()` profiles, then the lane specimen through the named
 * check directly.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { textboxContentNamespaceCheck } from "../src/scripts/office/checks/textbox-content-namespace";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

const CHECK_ID = "word-textbox-content-namespace";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WNE = "http://schemas.microsoft.com/office/word/2006/wordml";

// One schema-shaped wp:wsp textbox whose body is the Word Strict-writer
// wne:txbxContent quirk — the k12 defect.
const drawingWith = (txbxContent: string): string =>
    `<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing>` +
    `<wp:inline xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wne="${WNE}" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="2149522" cy="1207827"/><wp:docPr id="7" name="TextBox7"/>` +
    `<a:graphic><a:graphicData uri="${WPS}">` +
    `<wp:wsp><wp:cNvCnPr/><wp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2149522" cy="1207827"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wp:spPr>` +
    `<wp:txbx>${txbxContent}</wp:txbx><wp:bodyPr wrap="square"/>` +
    `</wp:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const INNER_PARAGRAPH = `<w:p w14:paraId="00000003" w14:textId="00000003"><w:r><w:t>boxed</w:t></w:r></w:p>`;
const DEFECTIVE_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace(
    `<w:sectPr>`,
    drawingWith(`<wne:txbxContent>${INNER_PARAGRAPH}</wne:txbxContent>`) + `<w:sectPr>`,
);
const CONTROL_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace(
    `<w:sectPr>`,
    drawingWith(`<wp:txbxContent>${INNER_PARAGRAPH}</wp:txbxContent>`) + `<w:sectPr>`,
);

async function validateBytes(bytes: Uint8Array, profile: Profile) {
    return withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
}

describe("textbox-content-namespace: synthetic single-defect package", { timeout: 120000 }, () => {
    it("word-valid stays valid and carries the advisory", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "word-valid");
        expect(result.valid).toBe(true);
        const findings = result.issues.filter((issue) => issue.code === CHECK_ID);
        expect(findings).toHaveLength(1);
        expect(findings[0]).toMatchObject({ severity: "warning", path: "word/document.xml" });
        expect(findings[0]!.message).toContain(WNE);
    });

    it("strict rejects the package", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "strict");
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.code === CHECK_ID && issue.severity === "warning")).toBe(true);
    });

    it("the wp:txbxContent control stays clean under strict", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": CONTROL_DOCUMENT }), "strict");
        expect(result.valid).toBe(true);
        expect(result.issues.map((issue) => issue.code)).not.toContain(CHECK_ID);
    });

    it("lenient does not carry the check's code", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "lenient");
        expect(result.issues.map((issue) => issue.code)).not.toContain(CHECK_ID);
    });
});

describe("textbox-content-namespace: jubarte 0.13.0 lane specimens", () => {
    it("k12_1.redline.docx: five wne:txbxContent sites in document.xml, one more in header2.xml", async () => {
        const { context } = await contextFromArchive(readFileSync(path.join(LANE, "k12_1.redline.docx")));
        const issues = await textboxContentNamespaceCheck.run(context);
        for (const issue of issues) {
            expect(issue).toMatchObject({ severity: "warning", code: CHECK_ID });
            expect(issue.message).toContain(WNE);
        }
        const byPath = new Map<string, number>();
        for (const issue of issues) {
            byPath.set(issue.path ?? "", (byPath.get(issue.path ?? "") ?? 0) + 1);
        }
        // Issue #244 counts the five wp:txbx sites in document.xml; the same
        // misuse also sits in header2.xml, which the issue did not enumerate.
        expect(byPath).toEqual(
            new Map([
                ["word/document.xml", 5],
                ["word/header2.xml", 1],
            ]),
        );
    });
});
