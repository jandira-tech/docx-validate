/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The style and font-table messages name the stored part they judged —
 * the same spelling the issue's `path` carries — so a styles part stored
 * as `word/xstyles.xml` or a font table under `word/zfonts/` is named by
 * its own name, not the default one (#695). A package that keeps the
 * default names produces byte-identical messages.
 */
import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const FONT_TABLE_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";

const EMPTY_FONT_TABLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="${W}"/>`;

// A styles part missing the implied-default `Normal` and any `w:latentStyles` block.
const WITHOUT_NORMAL = MINIMAL_PARTS["word/styles.xml"]!.replace(
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`,
    "",
);
// The document references one style id no styles part defines.
const UNDEFINED_STYLE_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace(
    `<w:p w14:paraId="00000001" w14:textId="00000001">`,
    `<w:p w14:paraId="00000001" w14:textId="00000001"><w:pPr><w:pStyle w:val="Undefined1"/></w:pPr>`,
);

const relationships = (...relationship: string[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship.join("")}</Relationships>`;
const relationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const override = (part: string, type: string): string => `<Override PartName="/${part}" ContentType="${type}"/>`;
const withOverrides = (...overrides: string[]): string =>
    MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", `${overrides.join("")}</Types>`);

/**
 * The same four defects under the default part names: `word/styles.xml`
 * without `Normal` or `w:latentStyles`, an `Undefined1` reference, and an
 * empty `word/fontTable.xml` — so the pinned messages stay byte-identical.
 */
const DEFAULT_NAMES: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/fontTable.xml", FONT_TABLE_TYPE)),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "fontTable", "fontTable.xml"),
    ),
    "word/document.xml": UNDEFINED_STYLE_DOCUMENT,
    "word/styles.xml": WITHOUT_NORMAL,
    "word/fontTable.xml": EMPTY_FONT_TABLE,
};

/**
 * The same four defects on renamed parts the main document relates:
 * `word/xstyles.xml` and `word/zfonts/fontTable.xml`, beside an unrelated
 * `word/styles.xml` that defines every referenced style.
 */
const RENAMED_PARTS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/xstyles.xml", STYLES_TYPE), override("word/zfonts/fontTable.xml", FONT_TABLE_TYPE)),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "xstyles.xml"),
        relationship("rId2", "fontTable", "zfonts/fontTable.xml"),
    ),
    "word/document.xml": UNDEFINED_STYLE_DOCUMENT,
    "word/xstyles.xml": WITHOUT_NORMAL,
    "word/zfonts/fontTable.xml": EMPTY_FONT_TABLE,
};

const validated = async (
    parts: Record<string, string>,
    profile: "strict" | "lenient" | "word-valid",
): Promise<readonly ValidationIssue[]> =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        await fs.writeFile(file, await buildDocx(parts));
        return (await validate(file, { profile })).issues;
    });

const messagesOf = (issues: readonly ValidationIssue[], code: string): string[] =>
    issues.filter((i) => i.code === code).map((i) => i.message);

const PROFILES = ["strict", "lenient", "word-valid"] as const;

describe("each style message names the stored part it judged (#695)", () => {
    it.each(PROFILES)("the renamed styles and font-table parts are named (%s)", async (profile) => {
        const issues = await validated(RENAMED_PARTS, profile);
        expect(messagesOf(issues, "style-default-missing")).toEqual([
            `word/xstyles.xml is missing the implied-default style 'Normal' ` +
                `(type='paragraph'). ECMA-376 §17.7.4.4 requires a default for each of the ` +
                `four style types; missing 'TableNormal' in particular triggers Word's ` +
                `"Document Recovery — Table Properties" dialog on open.`,
        ]);
        expect(messagesOf(issues, "style-reference-undefined")).toEqual([
            `<w:pStyle w:val='Undefined1'/> references a style not defined in word/xstyles.xml. ` +
                `Word substitutes the default style and triggers a repair dialog.`,
        ]);
        expect(messagesOf(issues, "latent-styles-missing")).toEqual([
            `word/xstyles.xml has no <w:latentStyles> block. ` +
                `This is unusual for Word-generated files and indicates ` +
                `the repairer may have regenerated styles from scratch.`,
        ]);
        expect(messagesOf(issues, "font-table-empty")).toEqual([
            `word/zfonts/fontTable.xml contains zero <w:font> entries. ` +
                `This causes Word to substitute all fonts with its default, ` +
                `silently changing document appearance.`,
        ]);
    });

    it.each(PROFILES)("the default-named package's messages are unchanged (%s)", async (profile) => {
        const issues = await validated(DEFAULT_NAMES, profile);
        expect(messagesOf(issues, "style-default-missing")).toEqual([
            `word/styles.xml is missing the implied-default style 'Normal' ` +
                `(type='paragraph'). ECMA-376 §17.7.4.4 requires a default for each of the ` +
                `four style types; missing 'TableNormal' in particular triggers Word's ` +
                `"Document Recovery — Table Properties" dialog on open.`,
        ]);
        expect(messagesOf(issues, "style-reference-undefined")).toEqual([
            `<w:pStyle w:val='Undefined1'/> references a style not defined in word/styles.xml. ` +
                `Word substitutes the default style and triggers a repair dialog.`,
        ]);
        expect(messagesOf(issues, "latent-styles-missing")).toEqual([
            `word/styles.xml has no <w:latentStyles> block. ` +
                `This is unusual for Word-generated files and indicates ` +
                `the repairer may have regenerated styles from scratch.`,
        ]);
        expect(messagesOf(issues, "font-table-empty")).toEqual([
            `word/fontTable.xml contains zero <w:font> entries. ` +
                `This causes Word to substitute all fonts with its default, ` +
                `silently changing document appearance.`,
        ]);
    });
});
