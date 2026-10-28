/**
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 *
 * The style checks read the styles part the main document relates, not
 * whichever `styles.xml` sorts last (or first, for table borders), so a
 * reached glossary's own styles part is never judged for the main
 * document (#688).
 */
import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const BORDERS = ["top", "left", "bottom", "right"]
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join("");

// The main styles part: the minimal one plus `Heading1` and a bordered table style `Grid`.
const MAIN_STYLES = MINIMAL_PARTS["word/styles.xml"]!.replace(
    "</w:styles>",
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>` +
        `<w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>` +
        `<w:tblPr><w:tblBorders>${BORDERS}</w:tblBorders></w:tblPr></w:style></w:styles>`,
);
// A styles part defining neither `Heading1` nor `Grid`.
const OTHER_STYLES = MINIMAL_PARTS["word/styles.xml"]!;

// The main document uses both styles; its one cell repeats the table's borders.
const BODY =
    `<w:p w14:paraId="00000002" w14:textId="00000002"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
    `<w:tr w14:paraId="00000003" w14:textId="00000003"><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcBorders>${BORDERS}</w:tcBorders></w:tcPr>` +
    `<w:p w14:paraId="00000004" w14:textId="00000004"><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${BODY}<w:sectPr>`);

const GLOSSARY_DOCUMENT =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="x"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{00000000-0000-0000-0000-000000000001}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>x</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

const relationships = (...relationship: string[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship.join("")}</Relationships>`;
const relationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const override = (part: string, type: string): string => `<Override PartName="/${part}" ContentType="${type}"/>`;
const withOverrides = (...overrides: string[]): string =>
    MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", `${overrides.join("")}</Types>`);

/** The main package plus a reached glossary stored under `word/<folder>/`, relating its own `OTHER_STYLES`. */
const withGlossary = (folder: string, document = DOCUMENT): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override(`word/${folder}/document.xml`, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"),
        override(`word/${folder}/styles.xml`, STYLES_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId9", "glossaryDocument", `${folder}/document.xml`),
    ),
    "word/document.xml": document,
    "word/styles.xml": MAIN_STYLES,
    [`word/${folder}/document.xml`]: GLOSSARY_DOCUMENT,
    [`word/${folder}/_rels/document.xml.rels`]: relationships(relationship("rId1", "styles", "styles.xml")),
    [`word/${folder}/styles.xml`]: OTHER_STYLES,
});

/** The main styles stored as `word/mystyles.xml` and related; an unrelated `word/styles.xml` defines neither style. */
const RENAMED_MAIN_STYLES: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/mystyles.xml", STYLES_TYPE)),
    "word/_rels/document.xml.rels": relationships(relationship("rId1", "styles", "mystyles.xml")),
    "word/document.xml": DOCUMENT,
    "word/mystyles.xml": MAIN_STYLES,
    "word/styles.xml": OTHER_STYLES,
};

/** No styles relationship at all: today's rule, the `styles.xml` by name, still applies. */
const UNRELATED_STYLES: Record<string, string> = {
    ...MINIMAL_PARTS,
    "word/_rels/document.xml.rels": relationships(),
    "word/document.xml": DOCUMENT,
    "word/styles.xml": MAIN_STYLES,
};

/**
 * The main document relates its styles as `word/xstyles.bin`, stored with a
 * non-XML content type; a reached glossary at `word/glossary/` holds the
 * package's only `styles.xml` (#693).
 */
const nonXmlStyles = (document = DOCUMENT): Record<string, string> => ({
    ...Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== "word/styles.xml")),
    "[Content_Types].xml": withOverrides(
        override("word/xstyles.bin", "application/octet-stream"),
        override("word/glossary/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"),
        override("word/glossary/styles.xml", STYLES_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "xstyles.bin"),
        relationship("rId9", "glossaryDocument", "glossary/document.xml"),
    ),
    "word/document.xml": document,
    "word/xstyles.bin": MAIN_STYLES,
    "word/glossary/document.xml": GLOSSARY_DOCUMENT,
    "word/glossary/_rels/document.xml.rels": relationships(relationship("rId1", "styles", "styles.xml")),
    "word/glossary/styles.xml": OTHER_STYLES,
});

const validated = async (
    parts: Record<string, string>,
    profile: "strict" | "lenient" | "word-valid",
): Promise<readonly ValidationIssue[]> =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        await fs.writeFile(file, await buildDocx(parts));
        return (await validate(file, { profile })).issues;
    });

const codesAt = (issues: readonly ValidationIssue[], code: string): string[] =>
    issues.filter((i) => i.code === code).map((i) => i.path ?? "-");

const PROFILES = ["strict", "lenient", "word-valid"] as const;

describe("style checks read the styles part the main document relates (#688)", () => {
    it.each(PROFILES)(
        "a glossary whose styles part sorts after word/styles.xml draws no style-reference-undefined (%s)",
        async (profile) => {
            const issues = await validated(withGlossary("zglossary"), profile);
            expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
            expect(codesAt(issues, "latent-styles-missing")).toEqual(["word/styles.xml"]);
            expect(codesAt(issues, "cell-borders-redundant")).toEqual(["word/document.xml"]);
        },
    );

    it.each(PROFILES)("a glossary at word/glossary/ does not hide the main table style's borders (%s)", async (profile) => {
        const issues = await validated(withGlossary("glossary"), profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
        expect(codesAt(issues, "latent-styles-missing")).toEqual(["word/styles.xml"]);
        expect(codesAt(issues, "cell-borders-redundant")).toEqual(["word/document.xml"]);
    });

    it.each(PROFILES)("a related styles part stored under another name is the one judged (%s)", async (profile) => {
        const issues = await validated(RENAMED_MAIN_STYLES, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
        expect(codesAt(issues, "latent-styles-missing")).toEqual(["word/mystyles.xml"]);
        expect(codesAt(issues, "cell-borders-redundant")).toEqual(["word/document.xml"]);
    });

    it.each(PROFILES)("without a styles relationship, word/styles.xml is still judged by name (%s)", async (profile) => {
        const issues = await validated(UNRELATED_STYLES, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
        expect(codesAt(issues, "latent-styles-missing")).toEqual(["word/styles.xml"]);
        expect(codesAt(issues, "cell-borders-redundant")).toEqual(["word/document.xml"]);
    });

    it("a related styles part missing a referenced style is still reported, not hidden by an unrelated one that defines it", async () => {
        const parts: Record<string, string> = { ...RENAMED_MAIN_STYLES, "word/mystyles.xml": OTHER_STYLES, "word/styles.xml": MAIN_STYLES };
        const issues = await validated(parts, "strict");
        expect(codesAt(issues, "style-reference-undefined")).toEqual(["word/document.xml", "word/document.xml"]);
    });

    it("repair defines a missing well-known style in the main styles part, not the glossary's", async () => {
        await withTempDir(async (dir) => {
            const document = DOCUMENT.replace("<w:t>Title</w:t>", `<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:t>Title</w:t>`);
            for (const [name, content] of Object.entries(withGlossary("zglossary", document))) {
                await fs.mkdir(path.dirname(path.join(dir, name)), { recursive: true });
                await fs.writeFile(path.join(dir, name), content, "utf-8");
            }
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairMissingStyleDefinitions();
            expect(repairs).toBe(1);
            expect(await fs.readFile(path.join(dir, "word/styles.xml"), "utf-8")).toContain(`w:styleId="CommentReference"`);
            expect(await fs.readFile(path.join(dir, "word/zglossary/styles.xml"), "utf-8")).toBe(OTHER_STYLES);
        });
    });
});

describe("a styles relationship naming a stored non-XML part judges nothing by name (#693)", () => {
    it.each(PROFILES)("the glossary's styles part is not judged for the main document (%s)", async (profile) => {
        const issues = await validated(nonXmlStyles(), profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
        expect(issues.filter((issue) => issue.path === "word/glossary/styles.xml").map((issue) => issue.code)).toEqual([]);
    });

    it("repair writes no definition into the glossary's styles part", async () => {
        await withTempDir(async (dir) => {
            const document = DOCUMENT.replace("<w:t>Title</w:t>", `<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:t>Title</w:t>`);
            for (const [name, content] of Object.entries(nonXmlStyles(document))) {
                await fs.mkdir(path.dirname(path.join(dir, name)), { recursive: true });
                await fs.writeFile(path.join(dir, name), content, "utf-8");
            }
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairMissingStyleDefinitions();
            expect(repairs).toBe(0);
            expect(await fs.readFile(path.join(dir, "word/glossary/styles.xml"), "utf-8")).toBe(OTHER_STYLES);
            expect(await fs.readFile(path.join(dir, "word/xstyles.bin"), "utf-8")).toBe(MAIN_STYLES);
        });
    });
});
