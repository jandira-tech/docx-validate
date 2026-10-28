/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Each story root's style references are judged against the styles part
 * that root relates (#698): a reached glossary document, and the parts
 * only the glossary's own relationships reach, against the glossary's
 * `styles` part — or not judged when the glossary relates none — never
 * against the main document's. A part both roots reach stays with the
 * main document, and the repair writes a glossary's missing well-known
 * style into the glossary's styles part only.
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
const GLOSSARY_DOCUMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
const COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const BORDERS = ["top", "left", "bottom", "right"]
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join("");

// The main styles part: the minimal one plus `Heading1` — never `GlossaryStyle` or `Grid`.
const MAIN_STYLES = MINIMAL_PARTS["word/styles.xml"]!.replace(
    "</w:styles>",
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>`,
);
// The glossary's own styles part: `GlossaryStyle` and the bordered table style `Grid`.
const GLOSSARY_STYLES = MINIMAL_PARTS["word/styles.xml"]!.replace(
    "</w:styles>",
    `<w:style w:type="paragraph" w:styleId="GlossaryStyle"><w:name w:val="glossary"/><w:basedOn w:val="Normal"/></w:style>` +
        `<w:style w:type="table" w:styleId="Grid"><w:name w:val="Glossary Grid"/><w:basedOn w:val="TableNormal"/>` +
        `<w:tblPr><w:tblBorders>${BORDERS}</w:tblBorders></w:tblPr></w:style></w:styles>`,
);

const glossaryDocument = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="x"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{00000000-0000-0000-0000-000000000001}"/></w:docPartPr>` +
    `<w:docPartBody>${body}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

const relationships = (...relationship: string[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship.join("")}</Relationships>`;
const relationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const override = (part: string, type: string): string => `<Override PartName="/${part}" ContentType="${type}"/>`;
const withOverrides = (...overrides: string[]): string =>
    MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", `${overrides.join("")}</Types>`);

/**
 * A reached glossary at `word/zglossary/` relating `styles.xml` and
 * `fontTable.xml` from its own sidecar; `glossaryBody` holds its story
 * markup and `glossaryRels` its relationships part.
 */
const withGlossary = (glossaryBody: string, glossaryStyles: string | null, glossaryRels = ""): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/zglossary/document.xml", GLOSSARY_DOCUMENT_TYPE),
        ...(glossaryStyles === null ? [] : [override("word/zglossary/styles.xml", STYLES_TYPE)]),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
    ),
    "word/styles.xml": MAIN_STYLES,
    "word/zglossary/document.xml": glossaryDocument(glossaryBody),
    "word/zglossary/_rels/document.xml.rels": relationships(
        ...(glossaryStyles === null ? [] : [relationship("rId1", "styles", "styles.xml")]),
        glossaryRels,
    ),
    ...(glossaryStyles === null ? {} : { "word/zglossary/styles.xml": glossaryStyles }),
});

/** The issue's superset specimen: the glossary uses only styles its own styles part defines. */
const SUPERSET = withGlossary(
    `<w:p><w:pPr><w:pStyle w:val="GlossaryStyle"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>` +
        `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    GLOSSARY_STYLES.replace(
        "</w:styles>",
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>`,
    ),
);

/** The superset specimen with the glossary stored in a mixed-case folder, `word/ZGlossary/`. */
const SUPERSET_MIXED_CASE: Record<string, string> = Object.fromEntries(
    Object.entries(SUPERSET).map(([name, content]) => [
        name.replace("zglossary/", "ZGlossary/"),
        content.replaceAll("zglossary/", "ZGlossary/"),
    ]),
);

/** The glossary uses `Heading1`, which only the MAIN styles part defines. */
const MAIN_ONLY_STYLE = withGlossary(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`, GLOSSARY_STYLES);

/** The glossary relates no styles part at all and uses a style nothing defines. */
const NO_GLOSSARY_STYLES = withGlossary(`<w:p><w:pPr><w:pStyle w:val="GlossaryStyle"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`, null);

/** The glossary run repeats its own styles part's docDefaults fonts, which the main styles part does not set. */
const GLOSSARY_RUN_FONTS = withGlossary(
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="GFont" w:hAnsi="GFont"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
    GLOSSARY_STYLES.replace("<w:rPr/>", `<w:rPr><w:rFonts w:ascii="GFont" w:hAnsi="GFont"/></w:rPr>`),
);

/** The glossary cell repeats the borders its own styles part's `Grid` declares. */
const GLOSSARY_CELL_BORDERS = withGlossary(
    `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tr><w:tc><w:tcPr><w:tcBorders>${BORDERS}</w:tcBorders></w:tcPr>` +
        `<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    GLOSSARY_STYLES,
);

const commentsUsing = (styleId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
    `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:comment></w:comments>`;

/**
 * The glossary's own comments part — reached only through its sidecar —
 * uses `Heading1`, which only the MAIN styles part defines.
 */
const GLOSSARY_COMMENTS_STYLE: Record<string, string> = {
    ...withGlossary(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`, GLOSSARY_STYLES, relationship("rId2", "comments", "comments.xml")),
    "[Content_Types].xml": withOverrides(
        override("word/zglossary/document.xml", GLOSSARY_DOCUMENT_TYPE),
        override("word/zglossary/styles.xml", STYLES_TYPE),
        override("word/zglossary/comments.xml", COMMENTS_TYPE),
    ),
    "word/zglossary/comments.xml": commentsUsing("Heading1"),
};

/** `word/comments.xml`, reached by BOTH roots, uses `GlossaryStyle`, which only the glossary's styles part defines. */
const SHARED_COMMENTS: Record<string, string> = {
    ...withGlossary(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`, GLOSSARY_STYLES, relationship("rId2", "comments", "../comments.xml")),
    "[Content_Types].xml": withOverrides(
        override("word/zglossary/document.xml", GLOSSARY_DOCUMENT_TYPE),
        override("word/zglossary/styles.xml", STYLES_TYPE),
        override("word/comments.xml", COMMENTS_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "comments", "comments.xml"),
        relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
    ),
    "word/comments.xml": commentsUsing("GlossaryStyle"),
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

const codesAt = (issues: readonly ValidationIssue[], code: string): string[] =>
    issues.filter((i) => i.code === code).map((i) => i.path ?? "-");

const messagesOf = (issues: readonly ValidationIssue[], code: string): string[] =>
    issues.filter((i) => i.code === code).map((i) => i.message);

const writeParts = async (dir: string, parts: Record<string, string>): Promise<void> => {
    for (const [name, content] of Object.entries(parts)) {
        await fs.mkdir(path.dirname(path.join(dir, name)), { recursive: true });
        await fs.writeFile(path.join(dir, name), content, "utf-8");
    }
};

const PROFILES = ["strict", "lenient", "word-valid"] as const;

describe("a reached glossary is judged against its own styles part (#698)", () => {
    it.each(PROFILES)("the superset specimen draws no style-reference-undefined (%s)", async (profile) => {
        const issues = await validated(SUPERSET, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
    });

    it.each(PROFILES)("a glossary stored under a mixed-case folder is judged against its own styles part (%s)", async (profile) => {
        const issues = await validated(SUPERSET_MIXED_CASE, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
    });

    it.each(PROFILES)("a style only the main styles part defines is reported against the glossary's (%s)", async (profile) => {
        const issues = await validated(MAIN_ONLY_STYLE, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual(["word/zglossary/document.xml"]);
        expect(messagesOf(issues, "style-reference-undefined")).toEqual([
            `<w:pStyle w:val='Heading1'/> references a style not defined in word/zglossary/styles.xml. ` +
                `Word substitutes the default style and triggers a repair dialog.`,
        ]);
    });

    it.each(PROFILES)("a glossary that relates no styles part is not judged (%s)", async (profile) => {
        const issues = await validated(NO_GLOSSARY_STYLES, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual([]);
    });

    it.each(PROFILES)("a glossary run is judged against the glossary's docDefaults (%s)", async (profile) => {
        const issues = await validated(GLOSSARY_RUN_FONTS, profile);
        expect(codesAt(issues, "run-props-redundant")).toEqual(["word/zglossary/document.xml"]);
    });

    it.each(PROFILES)("a glossary cell is judged against the glossary's table style (%s)", async (profile) => {
        const issues = await validated(GLOSSARY_CELL_BORDERS, profile);
        expect(codesAt(issues, "cell-borders-redundant")).toEqual(["word/zglossary/document.xml"]);
    });

    it.each(PROFILES)(
        "a part reached only through the glossary's relationships is judged against the glossary's styles (%s)",
        async (profile) => {
            const issues = await validated(GLOSSARY_COMMENTS_STYLE, profile);
            expect(codesAt(issues, "style-reference-undefined")).toEqual(["word/zglossary/comments.xml"]);
            expect(messagesOf(issues, "style-reference-undefined")).toEqual([
                `<w:pStyle w:val='Heading1'/> references a style not defined in word/zglossary/styles.xml. ` +
                    `Word substitutes the default style and triggers a repair dialog.`,
            ]);
        },
    );

    it.each(PROFILES)("a part both roots reach stays with the main document (%s)", async (profile) => {
        const issues = await validated(SHARED_COMMENTS, profile);
        expect(codesAt(issues, "style-reference-undefined")).toEqual(["word/comments.xml"]);
        expect(messagesOf(issues, "style-reference-undefined")).toEqual([
            `<w:pStyle w:val='GlossaryStyle'/> references a style not defined in word/styles.xml. ` +
                `Word substitutes the default style and triggers a repair dialog.`,
        ]);
    });

    it("repair writes a glossary's missing well-known style into the glossary's styles part only", async () => {
        await withTempDir(async (dir) => {
            const parts = withGlossary(
                `<w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
                GLOSSARY_STYLES,
            );
            await writeParts(dir, parts);
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairMissingStyleDefinitions();
            expect(repairs).toBe(1);
            expect(await fs.readFile(path.join(dir, "word/zglossary/styles.xml"), "utf-8")).toContain(`w:styleId="CommentReference"`);
            expect(await fs.readFile(path.join(dir, "word/styles.xml"), "utf-8")).toBe(MAIN_STYLES);
        });
    });
});
