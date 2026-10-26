/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFileSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { commentThreadingCheck } from "../src/scripts/office/checks/comments";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const broken = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures/broken");
const markerFixture = path.join(broken, "comments.unmatched-comment-marker.docx");
const commentIssues = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code?.startsWith("comment-"));

const PROFILES = ["word-valid", "lenient", "strict"] as const;

// #696: the comment checks read the parts the main document relates, so a
// reached glossary's own document and comment parts are never judged for
// the main document.
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const STRICT_W_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_COMMENTS_IDS_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const MAIN_DOCUMENT_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_IDS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
const GLOSSARY_DOCUMENT_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";

const GLOSSARY_DOCUMENT =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="${W14_NS}">` +
    `<w:docParts><w:docPart><w:docPartPr><w:name w:val="x"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{00000000-0000-0000-0000-000000000001}"/></w:docPartPr>` +
    `<w:docPartBody><w:p w14:paraId="00000009" w14:textId="00000009"><w:r><w:t>x</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

const commentsPart = (id: string, paraId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="${W14_NS}">` +
    `<w:comment w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
    `<w:p w14:paraId="${paraId}" w14:textId="${paraId}"><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>`;

const commentsIdsPart = (paraId: string, durableId: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w16cid:commentsIds xmlns:w16cid="${W16CID_NS}">` +
    `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/></w16cid:commentsIds>`;

/** The same comments part with its comment in the Strict WordprocessingML namespace, ids included. */
const strictCommentsPart = (id: string, paraId: string): string =>
    commentsPart(id, paraId).replace(`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`, `xmlns:w="${STRICT_W_NS}"`);

/** The minimal document with a paired marker triple for comment 0, plus `extra` markers. */
const markedDocument = (extra = ""): string =>
    MINIMAL_PARTS["word/document.xml"]!.replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        `<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${extra}`,
    );

/** The minimal document with the same paired marker triple for comment 0, spelled in the Strict WordprocessingML namespace. */
const strictMarkedDocument = (): string =>
    markedDocument()
        .replace('mc:Ignorable="w14"', `mc:Ignorable="w14" xmlns:ws="${STRICT_W_NS}"`)
        .replace('<w:commentRangeStart w:id="0"/>', '<ws:commentRangeStart ws:id="0"/>')
        .replace('<w:commentRangeEnd w:id="0"/>', '<ws:commentRangeEnd ws:id="0"/>')
        .replace('<w:commentReference w:id="0"/>', '<ws:commentReference ws:id="0"/>');

const override = (part: string, type: string): string => `<Override PartName="/${part}" ContentType="${type}"/>`;
const withOverrides = (...overrides: string[]): string =>
    MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", `${overrides.join("")}</Types>`);
const rel = (id: string, segment: string, target: string): string =>
    `<Relationship Id="${id}" Type="${REL_NS}/${segment}" Target="${target}"/>`;
const typedRel = (id: string, type: string, target: string): string => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
const relsPart = (...entries: string[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join("")}</Relationships>`;

/** The main document beside its comments part, optionally with a reached glossary document and the glossary's own comments. */
const commentPackage = (
    options: { readonly orphan?: boolean; readonly glossary?: boolean; readonly glossaryComments?: boolean } = {},
): Record<string, string> => {
    const parts: Record<string, string> = {
        ...MINIMAL_PARTS,
        "word/document.xml": markedDocument(options.orphan === true ? '<w:commentRangeStart w:id="1"/>' : ""),
        "word/comments.xml": commentsPart("0", "0000AAAA"),
    };
    const types = [override("word/comments.xml", COMMENTS_CT)];
    const documentRels = [rel("rId1", "styles", "styles.xml"), rel("rId2", "comments", "comments.xml")];
    if (options.glossary === true || options.glossaryComments === true) {
        documentRels.push(rel("rId9", "glossaryDocument", "glossary/document.xml"));
        parts["word/glossary/document.xml"] = GLOSSARY_DOCUMENT;
        types.push(override("word/glossary/document.xml", GLOSSARY_DOCUMENT_CT));
    }
    if (options.glossaryComments === true) {
        // The glossary's own comments hold a different comment id than the main comments.
        parts["word/glossary/comments.xml"] = commentsPart("5", "0000BBBB");
        parts["word/glossary/_rels/document.xml.rels"] = relsPart(rel("rId1", "comments", "comments.xml"));
        types.push(override("word/glossary/comments.xml", COMMENTS_CT));
    }
    parts["word/_rels/document.xml.rels"] = relsPart(...documentRels);
    parts["[Content_Types].xml"] = withOverrides(...types);
    return parts;
};

/** The main document relates its commentsIds by exact type at `word/threads/ids.xml`, beside an agreeing `word/commentsIds.xml` decoy. */
const RENAMED_COMMENTS_IDS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/comments.xml", COMMENTS_CT),
        override("word/threads/ids.xml", COMMENTS_IDS_CT),
        override("word/commentsIds.xml", COMMENTS_IDS_CT),
    ),
    "word/_rels/document.xml.rels": relsPart(
        rel("rId1", "styles", "styles.xml"),
        rel("rId2", "comments", "comments.xml"),
        typedRel("rId3", MS_COMMENTS_IDS_TYPE, "threads/ids.xml"),
    ),
    "word/document.xml": markedDocument(),
    "word/comments.xml": commentsPart("0", "0000AAAA"),
    "word/threads/ids.xml": commentsIdsPart("0000BBBB", "0000BBBB"),
    "word/commentsIds.xml": commentsIdsPart("0000AAAA", "0000AAAA"),
};

/** No comments relationship at all: the by-name parts in `word/` are judged, as before. */
const UNRELATED_COMMENT_PARTS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_CT), override("word/commentsIds.xml", COMMENTS_IDS_CT)),
    "word/document.xml": markedDocument('<w:commentRangeStart w:id="1"/>'),
    "word/comments.xml": commentsPart("0", "0000AAAA"),
    "word/commentsIds.xml": commentsIdsPart("0000BBBB", "0000BBBB"),
};

/** The main document stored at `main/doc.xml`, with a decoy `word/document.xml` holding an orphan marker. */
const MOVED_MAIN_DOCUMENT: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("main/doc.xml", MAIN_DOCUMENT_CT), override("main/comments.xml", COMMENTS_CT)),
    "_rels/.rels": relsPart(`<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="main/doc.xml"/>`),
    "main/doc.xml": markedDocument(),
    "main/comments.xml": commentsPart("0", "0000AAAA"),
    "main/_rels/doc.xml.rels": relsPart(rel("rId1", "styles", "../word/styles.xml"), rel("rId2", "comments", "comments.xml")),
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:commentRangeStart w:id="1"/><w:r><w:t>Hello</w:t></w:r>',
    ),
};

const BIN_DEFAULT = '<Default Extension="bin" ContentType="application/octet-stream"/>';
const ZGLOSSARY_RELATIONSHIP = rel("rId9", "glossaryDocument", "zglossary/document.xml");

/** The glossary at `word/zglossary/` holds the package's only comments, with an agreeing commentsIds; the main document has no markers. */
const glossaryOnlyComments = (mainRelatesIds: boolean): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/zglossary/document.xml", GLOSSARY_DOCUMENT_CT),
        override("word/zglossary/comments.xml", COMMENTS_CT),
        override("word/zglossary/commentsIds.xml", COMMENTS_IDS_CT),
        ...(mainRelatesIds ? [override("word/commentsIds.xml", COMMENTS_IDS_CT)] : []),
    ),
    "word/_rels/document.xml.rels": relsPart(
        rel("rId1", "styles", "styles.xml"),
        ZGLOSSARY_RELATIONSHIP,
        ...(mainRelatesIds ? [typedRel("rId3", MS_COMMENTS_IDS_TYPE, "commentsIds.xml")] : []),
    ),
    ...(mainRelatesIds ? { "word/commentsIds.xml": commentsIdsPart("0000DDDD", "0000DDDD") } : {}),
    "word/zglossary/document.xml": GLOSSARY_DOCUMENT,
    "word/zglossary/_rels/document.xml.rels": relsPart(rel("rId1", "comments", "comments.xml")),
    "word/zglossary/comments.xml": commentsPart("5", "0000BBBB"),
    "word/zglossary/commentsIds.xml": commentsIdsPart("0000BBBB", "0000BBBB"),
});

/** The main document relates its comments, and its commentsIds by exact type to `word/ids.bin`, a part that is not XML, beside a disagreeing by-name `word/commentsIds.xml`. */
const NON_XML_COMMENTS_IDS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(
        override("word/comments.xml", COMMENTS_CT),
        override("word/commentsIds.xml", COMMENTS_IDS_CT),
    ).replace("<Override", `${BIN_DEFAULT}<Override`),
    "word/_rels/document.xml.rels": relsPart(
        rel("rId1", "styles", "styles.xml"),
        rel("rId2", "comments", "comments.xml"),
        typedRel("rId3", MS_COMMENTS_IDS_TYPE, "ids.bin"),
    ),
    "word/document.xml": markedDocument(),
    "word/comments.xml": commentsPart("0", "0000AAAA"),
    "word/ids.bin": commentsIdsPart("0000BBBB", "0000BBBB"),
    "word/commentsIds.xml": commentsIdsPart("0000CCCC", "0000CCCC"),
};

/** The main document relates its comments to `word/notes.bin`, a part that is not XML, beside a by-name `word/comments.xml` its markers do not match. */
const NON_XML_COMMENTS: Record<string, string> = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_CT)).replace("<Override", `${BIN_DEFAULT}<Override`),
    "word/_rels/document.xml.rels": relsPart(rel("rId1", "styles", "styles.xml"), rel("rId2", "comments", "notes.bin")),
    "word/notes.bin": commentsPart("0", "0000AAAA"),
    "word/comments.xml": commentsPart("7", "0000AAAA"),
};

describe("dangling comment references (#784, session 58 s58a)", { timeout: 30000 }, () => {
    it.each(["5", "0"])("judges comment reference 0 against a part holding comment %s", async (held) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "comments.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                        "</w:r></w:p>",
                        '</w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>',
                    ),
                    "word/comments.xml": commentsPart(held, "0000AAAA"),
                    "word/_rels/document.xml.rels": relsPart(rel("rId1", "styles", "styles.xml"), rel("rIdN", "comments", "comments.xml")),
                    "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_CT)),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(held === "0");
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual(
                held === "0"
                    ? []
                    : [
                          expect.objectContaining({
                              code: "comment-marker-missing",
                              path: "word/document.xml",
                              message: expect.stringContaining('id="0"'),
                          }),
                      ],
            );
            const selected = await validate(file, { profile: "word-valid", checks: ["comments"] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues).toEqual(commentIssues(result.issues));
        });
    });

    it.each(["commentRangeStart", "commentRangeEnd"])("keeps a dangling %s without a commentReference advisory", async (marker) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "range-only.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...commentPackage(),
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("</w:p>", `<w:${marker} w:id="9"/></w:p>`),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues).toContainEqual(
                expect.objectContaining({
                    code: "comment-marker-missing",
                    severity: "warning",
                    message: expect.stringContaining('id="9"'),
                }),
            );
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });
});

/** The comments part of both #799 specimens, holding comment 5 outside any branch. */
const COMMENTS_HOLDING_FIVE = commentsPart("5", "0000AAAA");

describe("dangling comment references in an ISO Strict package (#817, session 63 s63a)", { timeout: 30000 }, () => {
    // Session 63 (s63a, two runs, all readings agree) opened the Word-saved ISO
    // Strict fixture `word-clean-strict01.docx` (pinned clean-open in session 46;
    // comment 0 fully marked, the comments part defining only comment 0) clean
    // unchanged, and offered recovery once one more run held
    // `w:commentReference w:id="99"` right after comment 0's reference: the
    // Strict twin of session 58's s58a refusal, so the reference blocks as its
    // Transitional counterpart does.
    const STRICT01 = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures/working/word-clean-strict01.docx");

    /** The fixture re-zipped, with `dangling` one more run holding `w:commentReference w:id="99"` after comment 0's reference. */
    const strict01 = async (dangling = false): Promise<Buffer> => {
        const zip = await JSZip.loadAsync(readFileSync(STRICT01));
        const parts: Record<string, string | Uint8Array> = {};
        for (const [name, entry] of Object.entries(zip.files)) {
            if (entry.dir) continue;
            parts[name] = await entry.async("uint8array");
        }
        if (dangling) {
            parts["word/document.xml"] = Buffer.from(parts["word/document.xml"] as Uint8Array)
                .toString("utf8")
                .replace(
                    '<w:commentReference w:id="0"/></w:r>',
                    '<w:commentReference w:id="0"/></w:r><w:r><w:commentReference w:id="99"/></w:r>',
                );
        }
        return buildDocx(parts);
    };

    it("the unchanged Strict calibration stays clean, as Word opened it in both runs", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "strict01.docx");
            await writeFile(file, await strict01());
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(commentIssues(result.issues)).toEqual([]);
        });
    });

    it("a Strict w:commentReference naming no defined comment blocks word-valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "strict01-dangling.docx");
            await writeFile(file, await strict01(true));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: "comment-marker-missing",
                    path: "word/document.xml",
                    message: 'marker id="99" references non-existent comment',
                }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: ["comments"] });
            expect(selected.valid).toBe(false);
            expect(selected.issues).toEqual(commentIssues(result.issues));
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues).toContainEqual(
                expect.objectContaining({ code: "comment-marker-missing", severity: "error", path: "word/document.xml" }),
            );
        });
    });

    /** The minimal package with `document` beside `comments`, the comments part related and typed. */
    const mixedSpecimen = (document: string, comments: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": document,
        "word/comments.xml": comments,
        "word/_rels/document.xml.rels": relsPart(rel("rId1", "styles", "styles.xml"), rel("rId2", "comments", "comments.xml")),
        "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_CT)),
    });

    // Unprobed: no Word session has shown a package mixing the two
    // WordprocessingML namespaces — session 63's s63a is uniformly Strict.
    // The ids are pooled across both namespaces, so a reference satisfied
    // only by a comment in the other passes; the Transitional direction
    // drew `comment-marker-missing` before #817 pooled them.
    it("pools the ids: a Transitional reference satisfied only by a Strict comment passes (unprobed)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "strict-comment.docx");
            await writeFile(file, await buildDocx(mixedSpecimen(markedDocument(), strictCommentsPart("0", "0000AAAA"))));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(markerMissingFindings(result.issues)).toEqual([]);
        });
    });

    it("pools the ids: a Strict reference satisfied only by a Transitional comment passes (unprobed)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "strict-reference.docx");
            await writeFile(file, await buildDocx(mixedSpecimen(strictMarkedDocument(), commentsPart("0", "0000AAAA"))));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(markerMissingFindings(result.issues)).toEqual([]);
        });
    });
});

/**
 * The minimal document's root with the unknown `Requires` prefix bound, so an
 * `mc:Choice Requires="u"` is never selected and its Fallback stands.
 */
const DOCUMENT_ROOT_WITH_UNKNOWN_PREFIX = MINIMAL_PARTS["word/document.xml"]!.replace(
    'mc:Ignorable="w14"',
    'mc:Ignorable="w14" xmlns:u="urn:example:unknown"',
);

/** The body's only comment reference sits inside the unselected Choice. */
const REFERENCE_ONLY_AN_UNSELECTED_CHOICE_HOLDS = DOCUMENT_ROOT_WITH_UNKNOWN_PREFIX.replace(
    "</w:r></w:p>",
    `</w:r><mc:AlternateContent><mc:Choice Requires="u"><w:r><w:commentReference w:id="0"/></w:r></mc:Choice>` +
        `<mc:Fallback><w:r><w:t>F</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
);

/** The body's comment reference 0 sits outside any AlternateContent. */
const REFERENCE_OUTSIDE_ANY_ALTERNATE_CONTENT = DOCUMENT_ROOT_WITH_UNKNOWN_PREFIX.replace(
    "</w:r></w:p>",
    '</w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>',
);

/** Session 62 (s62a): comment 0 is defined only inside an unselected Choice the comments part holds; the Fallback holds comment 5. */
const DEFINITION_ONLY_AN_UNSELECTED_CHOICE_HOLDS =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:example:unknown">` +
    `<mc:AlternateContent><mc:Choice Requires="u">` +
    `<w:comment w:id="0" w:author="A"><w:p><w:r><w:t>C0</w:t></w:r></w:p></w:comment></mc:Choice>` +
    `<mc:Fallback><w:comment w:id="5" w:author="A"><w:p><w:r><w:t>C5</w:t></w:r></w:p></w:comment></mc:Fallback>` +
    `</mc:AlternateContent></w:comments>`;

/** The minimal package with `document` and `comments` in place, its comments part related and typed. */
const alternateContentPackage = (document: string, comments: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": document,
    "word/comments.xml": comments,
    "word/_rels/document.xml.rels": relsPart(rel("rId1", "styles", "styles.xml"), rel("rId2", "comments", "comments.xml")),
    "[Content_Types].xml": withOverrides(override("word/comments.xml", COMMENTS_CT)),
});

const markerMissingFindings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === "comment-marker-missing");

describe("comment markers read the effective view (#799)", { timeout: 30000 }, () => {
    it("a reference only an unselected Choice holds names no missing comment", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "choice-reference.docx");
            await writeFile(
                file,
                await buildDocx(alternateContentPackage(REFERENCE_ONLY_AN_UNSELECTED_CHOICE_HOLDS, COMMENTS_HOLDING_FIVE)),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(markerMissingFindings(result.issues)).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: ["comment-markers"] });
            expect(selected.valid).toBe(true);
            expect(markerMissingFindings(selected.issues)).toEqual([]);
        });
    });

    it("a definition only an unselected Choice holds satisfies no reference outside one", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "choice-definition.docx");
            await writeFile(
                file,
                await buildDocx(
                    alternateContentPackage(REFERENCE_OUTSIDE_ANY_ALTERNATE_CONTENT, DEFINITION_ONLY_AN_UNSELECTED_CHOICE_HOLDS),
                ),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: "comment-marker-missing",
                    path: "word/document.xml",
                    message: 'marker id="0" references non-existent comment',
                }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: ["comment-markers"] });
            expect(selected.valid).toBe(false);
            expect(markerMissingFindings(selected.issues)).toEqual(markerMissingFindings(result.issues));

            // Session 62 (s62a) offered recovery for this shape. Strict reports the missing
            // comment, both count mismatches and the missing paraId, and no xsd finding: the
            // comments part keeps its own root.
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(strict.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: "comment-marker-missing",
                    path: "word/document.xml",
                    message: 'marker id="0" references non-existent comment',
                }),
                expect.objectContaining({
                    code: "comment-thread-count-mismatch",
                    path: "word/document.xml",
                    message: "commentRangeStart count (0) does not match comment count in comments.xml (1)",
                }),
                expect.objectContaining({
                    code: "comment-thread-count-mismatch",
                    path: "word/document.xml",
                    message: "commentRangeEnd count (0) does not match comment count in comments.xml (1)",
                }),
                expect.objectContaining({
                    code: "paraid-missing-element",
                    path: "word/comments.xml",
                }),
            ]);
        });
    });

    it("a reference outside any AlternateContent still names its missing comment (#784, session 58 s58a)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "plain-reference.docx");
            await writeFile(file, await buildDocx(alternateContentPackage(REFERENCE_OUTSIDE_ANY_ALTERNATE_CONTENT, COMMENTS_HOLDING_FIVE)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: "comment-marker-missing",
                    path: "word/document.xml",
                    message: 'marker id="0" references non-existent comment',
                }),
            ]);
        });
    });
});

describe("comment threading reads the effective view (#808)", { timeout: 30000 }, () => {
    const alternate = (requires: string, choice: string): string =>
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
        `xmlns:u="urn:example:unknown" xmlns:w14="${W14_NS}">` +
        `<mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
    const countFindings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === "comment-thread-count-mismatch");
    const mismatches = (markers: readonly string[], count: number, severity: "error" | "warning"): ValidationIssue[] =>
        markers.map((marker) => ({
            code: "comment-thread-count-mismatch",
            severity,
            path: "word/document.xml",
            message: `${marker} count (1) does not match comment count in comments.xml (${count})`,
        }));

    for (const requires of ["u", "w14"]) {
        it.each(["word-valid", "strict", "lenient"] as const)(
            `an extra reference in a Choice requiring ${requires} counts only when selected (%s)`,
            async (profile) => {
                await withTempDir(async (dir) => {
                    // Unprobed: no Word session has shown this extra-reference shape.
                    const parts = alternateContentPackage(
                        markedDocument(alternate(requires, '<w:r><w:commentReference w:id="0"/></w:r>')),
                        commentsPart("0", "0000AAAA"),
                    );
                    const file = path.join(dir, "extra-choice-reference.docx");
                    await writeFile(file, await buildDocx(parts));
                    const expected: ValidationIssue[] =
                        requires === "u"
                            ? []
                            : [
                                  {
                                      code: "comment-thread-count-mismatch",
                                      severity: profile === "word-valid" ? "warning" : "error",
                                      path: "word/document.xml",
                                      message: "commentReference count (2) does not match comment count in comments.xml (1)",
                                  },
                              ];
                    const full = await validate(file, { profile });
                    expect(countFindings(full.issues)).toEqual(expected);
                    expect(full.valid).toBe(profile === "word-valid" || requires === "u");
                    if (profile === "word-valid") {
                        const selected = await validate(file, { profile, checks: ["comment-threading"] });
                        expect(countFindings(selected.issues)).toEqual(expected);
                        expect(selected.valid).toBe(true);
                    } else {
                        // Strict refuses API selections; exercise the named check directly too,
                        // and under lenient alike.
                        expect(countFindings(await commentThreadingCheck.run(ValidationContext.fromEntries(parts), profile))).toEqual(
                            expected,
                        );
                    }
                });
            },
        );

        it.each(["word-valid", "strict", "lenient"] as const)(
            `an extra comment in a Choice requiring ${requires} counts only when selected (%s)`,
            async (profile) => {
                await withTempDir(async (dir) => {
                    // Unprobed: no Word session has shown this extra-definition shape.
                    const comments = commentsPart("0", "0000AAAA").replace(
                        "</w:comments>",
                        `${alternate(requires, '<w:comment w:id="7" w:author="A"><w:p w14:paraId="0000BBBB" w14:textId="0000BBBB"/></w:comment>')}</w:comments>`,
                    );
                    const parts = alternateContentPackage(markedDocument(), comments);
                    const file = path.join(dir, "extra-choice-comment.docx");
                    await writeFile(file, await buildDocx(parts));
                    const expected =
                        requires === "u"
                            ? []
                            : mismatches(
                                  ["commentRangeStart", "commentRangeEnd", "commentReference"],
                                  2,
                                  profile === "word-valid" ? "warning" : "error",
                              );
                    const full = await validate(file, { profile });
                    expect(countFindings(full.issues)).toEqual(expected);
                    expect(full.valid).toBe(profile === "word-valid" || requires === "u");
                    if (profile === "word-valid") {
                        const selected = await validate(file, { profile, checks: ["comment-threading"] });
                        expect(countFindings(selected.issues)).toEqual(expected);
                        expect(selected.valid).toBe(true);
                    } else {
                        expect(countFindings(await commentThreadingCheck.run(ValidationContext.fromEntries(parts), profile))).toEqual(
                            expected,
                        );
                    }
                });
            },
        );
    }

    it("keeps three advisory mismatches for session 62 (s62a)'s discarded reference", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "choice-reference.docx");
            await writeFile(
                file,
                await buildDocx(alternateContentPackage(REFERENCE_ONLY_AN_UNSELECTED_CHOICE_HOLDS, COMMENTS_HOLDING_FIVE)),
            );
            // Comment 5 has no markers at all, so the effective view still mismatches
            // all three counts. Word opened this shape clean in session 62 (s62a);
            // the count rule remains a warning under word-valid.
            const expected = ["commentRangeStart", "commentRangeEnd", "commentReference"].map((marker) => ({
                code: "comment-thread-count-mismatch",
                severity: "warning",
                path: "word/document.xml",
                message: `${marker} count (0) does not match comment count in comments.xml (1)`,
            }));
            const full = await validate(file, { profile: "word-valid" });
            expect(full.valid).toBe(true);
            expect(countFindings(full.issues)).toEqual(expected);
            const selected = await validate(file, { profile: "word-valid", checks: ["comment-threading"] });
            expect(selected.valid).toBe(true);
            expect(countFindings(selected.issues)).toEqual(expected);
        });
    });

    it("preserves the malformed comments part's exact threading diagnostic", async () => {
        const parts = alternateContentPackage(markedDocument(), "<broken");
        const ctx = ValidationContext.fromEntries(parts);
        const parsed = ctx.xml("word/comments.xml");
        expect(parsed?.ok).toBe(false);
        if (!parsed || parsed.ok) throw new Error("Expected the malformed comments part to fail parsing");
        expect(await commentThreadingCheck.run(ctx)).toEqual([
            {
                code: "comment-thread-parse",
                severity: "error",
                path: "word/comments.xml",
                message: `Error parsing XML: ${parsed.message}`,
            },
        ]);
    });

    for (const [part, root, ns, child, code] of [
        [
            "commentsIds.xml",
            "w16cid:commentsIds",
            W16CID_NS,
            '<w16cid:commentId w16cid:paraId="0000BBBB" w16cid:durableId="00000007"/>',
            "comment-thread-commentid-paraid-orphan",
        ],
        [
            "commentsExtensible.xml",
            "w16cex:commentsExtensible",
            "http://schemas.microsoft.com/office/word/2018/wordml/cex",
            '<w16cex:commentExtensible w16cex:durableId="00000007"/>',
            "comment-thread-durableid-orphan",
        ],
        [
            "commentsExtended.xml",
            "w15:commentsEx",
            "http://schemas.microsoft.com/office/word/2012/wordml",
            '<w15:commentEx w15:paraId="0000BBBB"/>',
            "comment-thread-paraid-orphan",
        ],
    ] as const) {
        it.each(["u", "w14"])(`${part} judges only a selected Choice requiring %s`, async (requires) => {
            await withTempDir(async (dir) => {
                // Unprobed: the schema's child sequence is read after branch selection.
                const file = path.join(dir, "thread-choice.docx");
                const prefix = root.split(":")[0];
                const matching = part === "commentsExtended.xml" ? '<w15:commentEx w15:paraId="0000AAAA"/>' : "";
                await writeFile(
                    file,
                    await buildDocx({
                        ...alternateContentPackage(markedDocument(), commentsPart("0", "0000AAAA")),
                        [`word/${part}`]: `<${root} xmlns:${prefix}="${ns}">${matching}${alternate(requires, child)}</${root}>`,
                    }),
                );
                const result = await validate(file, { profile: "word-valid", checks: ["comment-threading"] });
                expect(commentIssues(result.issues)).toEqual(
                    requires === "u"
                        ? []
                        : [
                              expect.objectContaining({
                                  code,
                                  path: `word/${part}`,
                                  severity: part === "commentsExtended.xml" ? "warning" : "error",
                              }),
                          ],
                );
                expect(result.valid).toBe(requires === "u" || part === "commentsExtended.xml");
            });
        });
    }
});

describe("named comment validation (#384)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "lenient"] as const)("selects comments with full-run findings under %s", async (profile) => {
        const selected = await validate(markerFixture, { profile, checks: ["comments"] });
        const full = await validate(markerFixture, { profile });
        expect(selected.issues.map((issue) => issue.code)).toEqual(
            expect.arrayContaining(["comment-marker-missing", "comment-orphan-start"]),
        );
        expect(selected.issues).toEqual(commentIssues(full.issues));
        expect(selected.metadata?.executedChecks).toEqual(["xml-syntax", "comment-markers", "comment-threading"]);
        expect(selected.xsdWalk?.started).toBe(false);
        expect(full.metadata?.executedChecks).toEqual(expect.arrayContaining(["comment-markers", "comment-threading"]));
        expect(new Set(full.metadata?.executedChecks).size).toBe(full.metadata?.executedChecks?.length);
    });

    it("keeps the full run's Word-blocking threading errors fatal", async () => {
        const file = path.join(broken, "sample-document.broken-tables.docx");
        const full = await validate(file, {});
        const selected = await validate(file, { checks: ["comments"] });
        expect(selected.issues).toEqual(commentIssues(full.issues));
        expect(selected.issues).toContainEqual(
            expect.objectContaining({ code: "comment-thread-commentid-paraid-orphan", severity: "error" }),
        );
        expect(selected.issues).toContainEqual(expect.objectContaining({ code: "comment-thread-durableid-orphan", severity: "error" }));
        expect(selected.valid).toBe(false);
    });

    it("keeps quick selection unchanged", async () => {
        const result = await validate(markerFixture, { quick: true });
        expect(commentIssues(result.issues)).toEqual([]);
        expect(result.metadata?.executedChecks).not.toContain("comment-markers");
        expect(result.metadata?.omittedChecks).toEqual(expect.arrayContaining(["comment-markers", "comment-threading"]));
    });

    it.each(["lenient", "word-valid", "strict"] as const)("preserves the soft threading rule's severity under %s", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "thread.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/comments.xml":
                        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0"><w:p w14:paraId="11111111"/></w:comment></w:comments>',
                    "word/commentsExtended.xml": '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
                }),
            );
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "comment-thread-paraid-missing")).toEqual([
                expect.objectContaining({ severity: profile === "strict" ? "error" : "warning", path: "word/commentsExtended.xml" }),
            ]);
        });
    });

    // A part name OPC treats as equal to the literal one (ASCII case,
    // percent-encoded unreserved characters) is the same part, as for every
    // other check. The legacy comment rules matched literal basenames only,
    // so before #384 these spellings went unchecked.
    it.each(["word/comments.XML", "word/%63omments.xml"])("judges %s as the comments part under every profile", async (name) => {
        await withTempDir(async (dir) => {
            const document = MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                '<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="7"/></w:r>',
            );
            const comments =
                '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p/></w:comment></w:comments>';
            const build = async (part: string, file: string) => {
                await writeFile(
                    path.join(dir, file),
                    await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": document, [part]: comments }),
                );
                return path.join(dir, file);
            };
            const literal = await build("word/comments.xml", "literal.docx");
            const equivalent = await build(name, "equivalent.docx");
            const findings = (issues: readonly ValidationIssue[]) =>
                commentIssues(issues).map(({ code, severity }) => ({ code, severity }));
            for (const profile of ["lenient", "word-valid", "strict"] as const) {
                const expected = findings((await validate(literal, { profile })).issues);
                expect(expected, profile).toContainEqual(expect.objectContaining({ code: "comment-marker-missing" }));
                expect(findings((await validate(equivalent, { profile })).issues), profile).toEqual(expected);
            }
        });
    });
});

describe("the comment checks read the parts the main document relates (#696)", { timeout: 120_000 }, () => {
    const commentFindings = (issues: readonly ValidationIssue[]): { code: string; path: string }[] =>
        commentIssues(issues).map((issue) => ({ code: issue.code ?? "", path: issue.path ?? "" }));

    const writeSpecimens = async (
        dir: string,
        specimens: readonly [name: string, parts: Record<string, string>][],
    ): Promise<Record<string, string>> => {
        const files: Record<string, string> = {};
        for (const [name, parts] of specimens) {
            files[name] = path.join(dir, `${name}.docx`);
            await writeFile(files[name]!, await buildDocx(parts));
        }
        return files;
    };

    it.each(PROFILES)("a reached glossary at word/glossary/ draws no comment finding (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const files = await writeSpecimens(dir, [
                ["control", commentPackage()],
                ["glossary", commentPackage({ glossary: true })],
            ]);
            const expected = commentFindings((await validate(files["control"]!, { profile })).issues);
            expect(expected).toEqual([]);
            const result = await validate(files["glossary"]!, { profile });
            expect(commentFindings(result.issues)).toEqual(expected);
            expect(result.valid).toBe(true);
        });
    });

    it("a reached glossary at word/glossary/ draws no comment finding in a selected comments run", async () => {
        await withTempDir(async (dir) => {
            const files = await writeSpecimens(dir, [
                ["control", commentPackage()],
                ["glossary", commentPackage({ glossary: true })],
            ]);
            const expected = commentFindings((await validate(files["control"]!, { checks: ["comments"] })).issues);
            expect(expected).toEqual([]);
            expect(commentFindings((await validate(files["glossary"]!, { checks: ["comments"] })).issues)).toEqual(expected);
        });
    });

    it.each(PROFILES)("an orphan marker in the main document is reported beside that glossary (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const files = await writeSpecimens(dir, [
                ["control", commentPackage({ orphan: true })],
                ["glossary", commentPackage({ orphan: true, glossary: true })],
            ]);
            const expected = commentFindings((await validate(files["control"]!, { profile })).issues);
            expect(expected).toEqual(
                expect.arrayContaining([
                    { code: "comment-orphan-start", path: "word/document.xml" },
                    { code: "comment-marker-missing", path: "word/document.xml" },
                ]),
            );
            expect(commentFindings((await validate(files["glossary"]!, { profile })).issues)).toEqual(expected);
        });
    });

    it.each(PROFILES)("the main document's markers are judged against the main comments, not a glossary's (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "glossary-comments.docx");
            await writeFile(file, await buildDocx(commentPackage({ glossary: true, glossaryComments: true })));
            // The main markers reference comment 0, which only the main comments hold.
            const found = commentFindings((await validate(file, { profile })).issues);
            expect(found.some((finding) => finding.code === "comment-marker-missing")).toBe(false);
            expect(found).toEqual([]);
        });
    });

    it.each(PROFILES)("the commentsIds part the main document relates is judged, not a by-name decoy (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "renamed-ids.docx");
            await writeFile(file, await buildDocx(RENAMED_COMMENTS_IDS));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([
                { code: "comment-thread-commentid-paraid-orphan", path: "word/threads/ids.xml" },
            ]);
        });
    });

    it.each(PROFILES)("without a comments relationship, the by-name parts are still judged (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unrelated.docx");
            await writeFile(file, await buildDocx(UNRELATED_COMMENT_PARTS));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([
                { code: "comment-orphan-start", path: "word/document.xml" },
                { code: "comment-marker-missing", path: "word/document.xml" },
                { code: "comment-thread-count-mismatch", path: "word/document.xml" },
                { code: "comment-thread-commentid-paraid-orphan", path: "word/commentsIds.xml" },
            ]);
        });
    });

    it.each(PROFILES)("a main document outside word/ is judged, not a word/document.xml decoy (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "moved-main.docx");
            await writeFile(file, await buildDocx(MOVED_MAIN_DOCUMENT));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([]);
        });
    });

    for (const [name, relationships] of [
        ["malformed", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships><Relationship Id="rId1"'],
        ["without an officeDocument target", relsPart()],
    ] as const) {
        it.each(["word-valid", "strict"] as const)(
            `${name} package relationships leave the comment fallback in word/ (#711, %s)`,
            async (profile) => {
                await withTempDir(async (dir) => {
                    // Review specimen 30, with an orphan main marker so the main pair must still be judged.
                    const file = path.join(dir, "fallback-comments.docx");
                    const parts: Record<string, string> = {
                        ...commentPackage({ orphan: true, glossary: true, glossaryComments: true }),
                        "_rels/.rels": relationships,
                    };
                    for (const part of ["word/glossary/document.xml", "word/glossary/comments.xml"]) {
                        parts[part] = parts[part]!.replace(
                            `xmlns:w14="${W14_NS}"`,
                            `xmlns:w14="${W14_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`,
                        );
                    }
                    await writeFile(file, await buildDocx(parts));
                    const result = await validate(file, { profile });
                    expect(commentFindings(result.issues)).toContainEqual({
                        code: "comment-thread-count-mismatch",
                        path: "word/document.xml",
                    });
                    expect(result.issues.filter((issue) => issue.path?.startsWith("word/glossary/"))).toEqual([]);
                });
            },
        );
    }

    it.each(PROFILES)("a glossary's own comments are never judged for a main document that relates none (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "glossary-only.docx");
            await writeFile(file, await buildDocx(glossaryOnlyComments(false)));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([]);
        });
    });

    it.each(PROFILES)("the main document's related commentsIds is not paired with a glossary's comments (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            // Review of layers 140-142, finding 1: the comments part found by name must not be paired with a
            // thread part the main document relates by exact type.
            const file = path.join(dir, "main-ids.docx");
            await writeFile(file, await buildDocx(glossaryOnlyComments(true)));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([]);
        });
    });

    it.each(PROFILES)("a related commentsIds part that is not XML is left unjudged, with no by-name stand-in (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "non-xml-ids.docx");
            await writeFile(file, await buildDocx(NON_XML_COMMENTS_IDS));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([]);
        });
    });

    it.each(PROFILES)("a related comments part that is not XML is left unjudged, with no by-name stand-in (%s)", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "non-xml-comments.docx");
            await writeFile(file, await buildDocx(NON_XML_COMMENTS));
            expect(commentFindings((await validate(file, { profile })).issues)).toEqual([]);
        });
    });
});
