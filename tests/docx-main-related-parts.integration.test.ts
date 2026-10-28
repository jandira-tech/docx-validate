/**
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 *
 * The font-table check and the comment-thread repair read the parts the
 * main document relates, not whichever twin sorts last by name, so a
 * reached glossary's own parts are never judged or rewritten for the main
 * document (#690).
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
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_COMMENTS_IDS_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const MS_COMMENTS_EXTENSIBLE_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
const W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const FONT_TABLE_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_IDS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
const GLOSSARY_DOCUMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";

const EMPTY_FONT_TABLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="${W}"/>`;
const FULL_FONT_TABLE =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:fonts xmlns:w="${W}"><w:font w:name="Inter"><w:family w:val="swiss"/></w:font></w:fonts>`;

const GLOSSARY_DOCUMENT =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="x"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{00000000-0000-0000-0000-000000000001}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>x</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

const commentsXml = (paraId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
    `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
    `<w:p w14:paraId="${paraId}" w14:textId="${paraId}"><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>`;

const commentsIdsXml = (paraId: string, durableId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w16cid:commentsIds xmlns:w16cid="${W16CID}">` +
    `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/></w16cid:commentsIds>`;

const relationships = (...relationship: string[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship.join("")}</Relationships>`;
const relationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const typedRelationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
const override = (part: string, type: string): string => `<Override PartName="/${part}" ContentType="${type}"/>`;
const withOverrides = (...overrides: string[]): string =>
    MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", `${overrides.join("")}</Types>`);

/** The main package plus a reached glossary stored under `word/<folder>/`, each relating its own font table. */
const withGlossaryFontTable = (folder: string, main: string, glossary: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/fontTable.xml", FONT_TABLE_TYPE),
        override(`word/${folder}/document.xml`, GLOSSARY_DOCUMENT_TYPE),
        override(`word/${folder}/fontTable.xml`, FONT_TABLE_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "fontTable", "fontTable.xml"),
        relationship("rId9", "glossaryDocument", `${folder}/document.xml`),
    ),
    "word/fontTable.xml": main,
    [`word/${folder}/document.xml`]: GLOSSARY_DOCUMENT,
    [`word/${folder}/_rels/document.xml.rels`]: relationships(relationship("rId1", "fontTable", "fontTable.xml")),
    [`word/${folder}/fontTable.xml`]: glossary,
});

/** No fontTable relationship at all: the last `fontTable.xml` by name, `word/fontTable.xml`, is still judged. */
const UNRELATED_FONT_TABLE: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/fontTable.xml", FONT_TABLE_TYPE),
        override("word/afont/fontTable.xml", FONT_TABLE_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(relationship("rId1", "styles", "styles.xml")),
    "word/fontTable.xml": EMPTY_FONT_TABLE,
    "word/afont/fontTable.xml": FULL_FONT_TABLE,
};

/** A reached glossary under `word/aglossary/` with its own `comments.xml`/`commentsIds.xml` pair, and a main `comments.xml` the main document relates. */
const glossaryPair = (mainRels: string[], glossaryCommentsParaId: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/comments.xml", COMMENTS_TYPE),
        override("word/aglossary/document.xml", GLOSSARY_DOCUMENT_TYPE),
        override("word/aglossary/comments.xml", COMMENTS_TYPE),
        override("word/aglossary/commentsIds.xml", COMMENTS_IDS_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(...mainRels),
    "word/comments.xml": commentsXml("0000AAAA"),
    "word/aglossary/document.xml": GLOSSARY_DOCUMENT,
    "word/aglossary/_rels/document.xml.rels": relationships(
        relationship("rId1", "comments", "comments.xml"),
        typedRelationship("rId2", MS_COMMENTS_IDS_TYPE, "commentsIds.xml"),
    ),
    "word/aglossary/comments.xml": commentsXml(glossaryCommentsParaId),
    "word/aglossary/commentsIds.xml": commentsIdsXml(glossaryCommentsParaId, glossaryCommentsParaId),
});

/** The main document relates its comments but no commentsIds: only the glossary holds a commentsIds part. */
const CROSS_PAIRED = glossaryPair(
    [
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "comments", "comments.xml"),
        relationship("rId9", "glossaryDocument", "aglossary/document.xml"),
    ],
    "0000BBBB",
);

/** The main document relates its comments and its commentsIds; the glossary holds its own consistent pair. */
const MAIN_PAIRED: Record<string, string> = {
    ...glossaryPair(
        [
            relationship("rId1", "styles", "styles.xml"),
            relationship("rId2", "comments", "comments.xml"),
            typedRelationship("rId3", MS_COMMENTS_IDS_TYPE, "commentsIds.xml"),
            relationship("rId9", "glossaryDocument", "aglossary/document.xml"),
        ],
        "0000CCCC",
    ),
    "[Content_Types].xml": withOverrides(
        override("word/comments.xml", COMMENTS_TYPE),
        override("word/commentsIds.xml", COMMENTS_IDS_TYPE),
        override("word/aglossary/document.xml", GLOSSARY_DOCUMENT_TYPE),
        override("word/aglossary/comments.xml", COMMENTS_TYPE),
        override("word/aglossary/commentsIds.xml", COMMENTS_IDS_TYPE),
    ),
    "word/commentsIds.xml": commentsIdsXml("0000BBBB", "0000BBBB"),
};

const commentsExtensibleXml = (durableId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">` +
    `<w16cex:commentExtensible w16cex:durableId="${durableId}" w16cex:dateUtc="2026-01-01T00:00:00Z"/></w16cex:commentsExtensible>`;

/**
 * The main document relates its comment-thread parts under other names in
 * `word/threads/`, beside unrelated `commentsIds.xml` and
 * `commentsExtensible.xml` decoys in `word/`: the relationships, not the
 * names, choose the parts the repair rewrites.
 */
const RENAMED_THREAD_PARTS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/comments.xml", COMMENTS_TYPE),
        override("word/threads/ids.xml", COMMENTS_IDS_TYPE),
        override("word/commentsIds.xml", COMMENTS_IDS_TYPE),
    ),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "comments", "comments.xml"),
        typedRelationship("rId3", MS_COMMENTS_IDS_TYPE, "threads/ids.xml"),
        typedRelationship("rId4", MS_COMMENTS_EXTENSIBLE_TYPE, "threads/cex.xml"),
    ),
    "word/comments.xml": commentsXml("0000AAAA"),
    "word/threads/ids.xml": commentsIdsXml("0000BBBB", "0000BBBB"),
    "word/threads/cex.xml": commentsExtensibleXml("0000BBBB"),
    "word/commentsIds.xml": commentsIdsXml("0000DDDD", "0000DDDD"),
    "word/commentsExtensible.xml": commentsExtensibleXml("0000DDDD"),
};

/** No comments relationship at all: the by-name pair in `word/` is still repaired. */
const UNRELATED_COMMENTS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_TYPE), override("word/commentsIds.xml", COMMENTS_IDS_TYPE)),
    "word/_rels/document.xml.rels": relationships(relationship("rId1", "styles", "styles.xml")),
    "word/comments.xml": commentsXml("0000AAAA"),
    "word/commentsIds.xml": commentsIdsXml("0000BBBB", "0000BBBB"),
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

const writeParts = async (dir: string, parts: Record<string, string>): Promise<void> => {
    for (const [name, content] of Object.entries(parts)) {
        await fs.mkdir(path.dirname(path.join(dir, name)), { recursive: true });
        await fs.writeFile(path.join(dir, name), content, "utf-8");
    }
};

const PROFILES = ["strict", "lenient", "word-valid"] as const;

describe("the font-table check reads the main document's font table (#690)", () => {
    it.each(PROFILES)("an empty main font table is reported beside a full glossary twin at word/glossary/ (%s)", async (profile) => {
        const issues = await validated(withGlossaryFontTable("glossary", EMPTY_FONT_TABLE, FULL_FONT_TABLE), profile);
        expect(codesAt(issues, "font-table-empty")).toEqual(["word/fontTable.xml"]);
    });

    it.each(PROFILES)("an empty main font table is reported beside a full glossary twin at word/zglossary/ (%s)", async (profile) => {
        const issues = await validated(withGlossaryFontTable("zglossary", EMPTY_FONT_TABLE, FULL_FONT_TABLE), profile);
        expect(codesAt(issues, "font-table-empty")).toEqual(["word/fontTable.xml"]);
    });

    it.each(PROFILES)("a full main font table is not reported beside an empty glossary twin (%s)", async (profile) => {
        const issues = await validated(withGlossaryFontTable("zglossary", FULL_FONT_TABLE, EMPTY_FONT_TABLE), profile);
        expect(codesAt(issues, "font-table-empty")).toEqual([]);
    });

    it.each(PROFILES)("without a fontTable relationship, the last fontTable.xml by name is still judged (%s)", async (profile) => {
        const issues = await validated(UNRELATED_FONT_TABLE, profile);
        expect(codesAt(issues, "font-table-empty")).toEqual(["word/fontTable.xml"]);
    });
});

describe("the comment-thread repair pairs parts within the main document (#690)", () => {
    it("never pairs a glossary's commentsIds with the main comments", async () => {
        await withTempDir(async (dir) => {
            await writeParts(dir, CROSS_PAIRED);
            const glossaryIds = path.join(dir, "word", "aglossary", "commentsIds.xml");
            const before = await fs.readFile(glossaryIds, "utf-8");
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairCommentThreading();
            expect(repairs).toBe(0);
            expect(await fs.readFile(glossaryIds, "utf-8")).toBe(before);
        });
    });

    it("repairs the main commentsIds the main document relates, leaving the glossary's pair alone", async () => {
        await withTempDir(async (dir) => {
            await writeParts(dir, MAIN_PAIRED);
            const glossaryIds = path.join(dir, "word", "aglossary", "commentsIds.xml");
            const glossaryBefore = await fs.readFile(glossaryIds, "utf-8");
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairCommentThreading();
            expect(repairs).toBe(2);
            const mainIds = await fs.readFile(path.join(dir, "word", "commentsIds.xml"), "utf-8");
            expect(mainIds).toContain('w16cid:paraId="0000AAAA"');
            expect(mainIds).toContain('w16cid:durableId="0000AAAA"');
            expect(await fs.readFile(glossaryIds, "utf-8")).toBe(glossaryBefore);
        });
    });

    it("repairs the thread parts the main document relates by exact type, not the by-name decoys", async () => {
        await withTempDir(async (dir) => {
            await writeParts(dir, RENAMED_THREAD_PARTS);
            const decoys = ["commentsIds.xml", "commentsExtensible.xml"].map((name) => path.join(dir, "word", name));
            const decoysBefore = await Promise.all(decoys.map((decoy) => fs.readFile(decoy, "utf-8")));
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairCommentThreading();
            expect(repairs).toBe(3);
            const ids = await fs.readFile(path.join(dir, "word", "threads", "ids.xml"), "utf-8");
            expect(ids).toContain('w16cid:paraId="0000AAAA"');
            expect(ids).toContain('w16cid:durableId="0000AAAA"');
            expect(await fs.readFile(path.join(dir, "word", "threads", "cex.xml"), "utf-8")).toContain('w16cex:durableId="0000AAAA"');
            expect(await Promise.all(decoys.map((decoy) => fs.readFile(decoy, "utf-8")))).toEqual(decoysBefore);
        });
    });

    it("without a comments relationship, the by-name pair is still repaired", async () => {
        await withTempDir(async (dir) => {
            await writeParts(dir, UNRELATED_COMMENTS);
            const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairCommentThreading();
            expect(repairs).toBe(2);
            const mainIds = await fs.readFile(path.join(dir, "word", "commentsIds.xml"), "utf-8");
            expect(mainIds).toContain('w16cid:paraId="0000AAAA"');
        });
    });
});
