/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { wordSharedStoryPartCheck } from "../src/scripts/office/checks/word-shared-story-part";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { climbingSharedNotes } from "./support/climbing-target-specimens";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CT = "[Content_Types].xml";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const GLOSSARY = "word/glossary/document.xml";
const GLOSSARY_SIDECAR = "word/glossary/_rels/document.xml.rels";
const GLOSSARY_CT = `${WML}.document.glossary+xml`;
const P = MINIMAL_PARTS as Record<string, string>;
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};
const relationship = (id: string, segment: string, target: string): string =>
    `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"/>`;
const appendRels = (rels: string, extra: string): string => replaceOnce(rels, "</Relationships>", `${extra}</Relationships>`);
const overrides = (ct: string, ...entries: [string, string][]): string =>
    replaceOnce(
        ct,
        "</Types>",
        `${entries.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join("")}</Types>`,
    );
/** `parts` with the main document's `r:` prefix declared, `sect` first in its `w:sectPr` and `run` appended inside its only paragraph. */
const mainWith = (parts: Record<string, string>, run: string, sect = ""): Record<string, string> => ({
    ...parts,
    [MAIN]: replaceOnce(
        replaceOnce(replaceOnce(parts[MAIN]!, `xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`), "<w:sectPr>", `<w:sectPr>${sect}`),
        "</w:r></w:p>",
        `</w:r>${run}</w:p>`,
    ),
});

/** A glossary document whose docPart paragraph also holds `run`, with `sect` after it as the body's `w:sectPr`. */
const glossaryDocument = (run: string, sect = ""): string =>
    `${XML_DECL}<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>Text</w:t></w:r>${run}</w:p>${sect}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
/** The minimal parts with that glossary reached, its sidecar holding `rels`. */
const glossaryParts = (run = "", rels = "", sect = ""): Record<string, string> => ({
    ...P,
    [GLOSSARY]: glossaryDocument(run, sect),
    [GLOSSARY_SIDECAR]: `${XML_DECL}<Relationships xmlns="${PR}">${rels}</Relationships>`,
    [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdG", "glossaryDocument", "glossary/document.xml")),
    [CT]: overrides(P[CT]!, [GLOSSARY, GLOSSARY_CT]),
});

const comment0 = `<w:comment w:id="0" w:author="a"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:comment>`;
const commentRef = (id: string) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;
const noteRef = (kind: "footnote" | "endnote", id: string) => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
const notesXml = (kind: "footnote" | "endnote"): string =>
    `${XML_DECL}<w:${kind}s xmlns:w="${W}"><w:${kind} w:id="1"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:${kind}></w:${kind}s>`;

/**
 * Session 66's shared comments part (s66a `comments-both-reaches-…`,
 * session 67's s67a repeats): `word/comments.xml` with comment 0, related
 * by the main document and/or the glossary, referenced as the flags say.
 */
const commentsShared = (mainRel: boolean, mainRef: boolean, glossRel: boolean, glossRef: boolean): Record<string, string> => {
    const parts = glossaryParts(glossRef ? commentRef("0") : "");
    parts["word/comments.xml"] = `${XML_DECL}<w:comments xmlns:w="${W}">${comment0}</w:comments>`;
    parts[CT] = overrides(parts[CT]!, ["word/comments.xml", `${WML}.comments+xml`]);
    if (glossRel) parts[GLOSSARY_SIDECAR] = appendRels(parts[GLOSSARY_SIDECAR]!, relationship("rIdC", "comments", "../comments.xml"));
    if (mainRel) parts[SIDECAR] = appendRels(parts[SIDECAR]!, relationship("rIdC", "comments", "comments.xml"));
    return mainRef ? mainWith(parts, commentRef("0")) : parts;
};

/**
 * Session 67's shared notes part (s67a): `word/<kind>s.xml` with note 1,
 * related by the main document and by the glossary as `../<kind>s.xml`,
 * each referencing note 1 when its flag says.
 */
const notesShared = (kind: "footnote" | "endnote", mainRef: boolean, glossRel: boolean, glossRef: boolean): Record<string, string> => {
    const parts = glossaryParts(glossRef ? noteRef(kind, "1") : "");
    parts[`word/${kind}s.xml`] = notesXml(kind);
    parts[CT] = overrides(parts[CT]!, [`word/${kind}s.xml`, `${WML}.${kind}s+xml`]);
    parts[SIDECAR] = appendRels(parts[SIDECAR]!, relationship("rIdN", `${kind}s`, `${kind}s.xml`));
    if (glossRel) parts[GLOSSARY_SIDECAR] = appendRels(parts[GLOSSARY_SIDECAR]!, relationship("rIdN", `${kind}s`, `../${kind}s.xml`));
    return mainRef ? mainWith(parts, noteRef(kind, "1")) : parts;
};

/** The unprobed two-glossary share: both reached glossaries relate and reference `word/comments.xml`, the main document neither. */
const twoGlossariesSharedComments = (): Record<string, string> => ({
    ...P,
    "word/comments.xml": `${XML_DECL}<w:comments xmlns:w="${W}">${comment0}</w:comments>`,
    "word/glossary/g1.xml": glossaryDocument(commentRef("0")),
    "word/glossary/g2.xml": glossaryDocument(commentRef("0")),
    "word/glossary/_rels/g1.xml.rels": `${XML_DECL}<Relationships xmlns="${PR}">${relationship("rIdC", "comments", "../comments.xml")}</Relationships>`,
    "word/glossary/_rels/g2.xml.rels": `${XML_DECL}<Relationships xmlns="${PR}">${relationship("rIdC", "comments", "../comments.xml")}</Relationships>`,
    [SIDECAR]: appendRels(
        P[SIDECAR]!,
        relationship("rIdG1", "glossaryDocument", "glossary/g1.xml") + relationship("rIdG2", "glossaryDocument", "glossary/g2.xml"),
    ),
    [CT]: overrides(
        P[CT]!,
        ["word/comments.xml", `${WML}.comments+xml`],
        ["word/glossary/g1.xml", GLOSSARY_CT],
        ["word/glossary/g2.xml", GLOSSARY_CT],
    ),
});

/** The unprobed shared header: both roots relate and reference `word/header1.xml` through the section's default header. */
const sharedHeader = (): Record<string, string> => {
    const reference = `<w:headerReference w:type="default" r:id="rIdH"/>`;
    const parts = glossaryParts("", relationship("rIdH", "header", "../header1.xml"), `<w:sectPr>${reference}</w:sectPr>`);
    return mainWith(
        {
            ...parts,
            "word/header1.xml": `${XML_DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
            [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdH", "header", "header1.xml")),
            [CT]: overrides(parts[CT]!, ["word/header1.xml", `${WML}.header+xml`]),
        },
        "",
        reference,
    );
};

/** The unprobed cross-segment share: the main document reads `word/comments.xml` through its comments relationship, the glossary through a footnotes one. */
const crossSegmentShare = (): Record<string, string> => {
    const parts = glossaryParts(noteRef("footnote", "1"), relationship("rIdX", "footnotes", "../comments.xml"));
    return mainWith(
        {
            ...parts,
            "word/comments.xml": `${XML_DECL}<w:comments xmlns:w="${W}">${comment0}</w:comments>`,
            [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdC", "comments", "comments.xml")),
            [CT]: overrides(parts[CT]!, ["word/comments.xml", `${WML}.comments+xml`]),
        },
        commentRef("0"),
    );
};

const diagnose = async (parts: Record<string, string>) => await wordSharedStoryPartCheck.run(ValidationContext.fromEntries(parts));

describe("word-shared-story-part (#850)", () => {
    it("registers behind xml-syntax in the story group for DOCX", () => {
        const check = createDefaultCheckRegistry().get("word-shared-story-part");
        expect(check).toBe(wordSharedStoryPartCheck);
        expect(check?.groups).toEqual(["story"]);
        expect(check?.formats).toEqual(["docx"]);
        expect(check?.requires).toEqual(["xml-syntax"]);
    });

    it("rejects the comments part both story roots read, citing session 66 (s66a)", async () => {
        expect(await diagnose(commentsShared(true, true, true, true))).toEqual([
            {
                severity: "error",
                code: "word-shared-story-part",
                path: "word/comments.xml",
                message:
                    "word/comments.xml is the comments part of both the main document and the reached glossary word/glossary/document.xml. " +
                    "Microsoft Word offers to recover the file (session 66). Give each story root its own comments part.",
            },
        ]);
    });

    it.each([["footnote"], ["endnote"]] as const)(
        "rejects the shared %ss part both story roots read, citing session 67 (s67a)",
        async (kind) => {
            expect(await diagnose(notesShared(kind, true, true, true))).toEqual([
                {
                    severity: "error",
                    code: "word-shared-story-part",
                    path: `word/${kind}s.xml`,
                    message:
                        `word/${kind}s.xml is the ${kind}s part of both the main document and the reached glossary word/glossary/document.xml. ` +
                        `Microsoft Word offers to recover the file (session 67). Give each story root its own ${kind}s part.`,
                },
            ]);
        },
    );

    it("accepts the same comments part read by one root only (sessions 65 s65d, 66 s66a)", async () => {
        // Read by the main document alone, and by the glossary alone.
        expect(await diagnose(commentsShared(true, true, false, false))).toEqual([]);
        expect(await diagnose(commentsShared(false, false, true, true))).toEqual([]);
    });

    it("accepts a part both roots relate but only one uses, or neither (session 67 s67a)", async () => {
        expect(await diagnose(commentsShared(true, true, true, false))).toEqual([]);
        expect(await diagnose(commentsShared(true, false, true, true))).toEqual([]);
        expect(await diagnose(commentsShared(true, false, true, false))).toEqual([]);
        expect(await diagnose(notesShared("footnote", true, true, false))).toEqual([]);
    });

    // The read model keeps a lazy target whose usage it cannot decide
    // (#437), so a reach alone does not show that root uses the part; the
    // rule needs both roots' use known (#858).
    it("reports no shared part when one root's usage is unknown", async () => {
        const unknownGlossary = commentsShared(true, true, true, false);
        unknownGlossary[GLOSSARY] = replaceOnce(unknownGlossary[GLOSSARY]!, 'encoding="UTF-8"', 'encoding="bogus"');
        expect(await diagnose(unknownGlossary)).toEqual([]);
        const unknownMain = commentsShared(true, false, true, true);
        unknownMain[MAIN] = replaceOnce(unknownMain[MAIN]!, 'encoding="UTF-8"', 'encoding="bogus"');
        expect(await diagnose(unknownMain)).toEqual([]);
        const unknownNotes = notesShared("footnote", true, true, false);
        unknownNotes[GLOSSARY] = replaceOnce(unknownNotes[GLOSSARY]!, 'encoding="UTF-8"', 'encoding="bogus"');
        expect(await diagnose(unknownNotes)).toEqual([]);
    });

    it("keeps the unprobed shares silent: a shared header, two glossaries, a cross-segment pair", async () => {
        expect(await diagnose(sharedHeader())).toEqual([]);
        expect(await diagnose(twoGlossariesSharedComments())).toEqual([]);
        expect(await diagnose(crossSegmentShare())).toEqual([]);
    });

    it("skips absent and unparseable packages", async () => {
        expect(await wordSharedStoryPartCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(await wordSharedStoryPartCheck.run(ValidationContext.fromEntries({ "_rels/.rels": "<broken" }))).toEqual([]);
    });
});

describe("full word-valid runs on the session 66 and 67 specimens", { timeout: 30000 }, () => {
    it("rejects every shared part and accepts every single-root calibration", async () => {
        await withTempDir(async (dir) => {
            const wordValid = async (parts: Record<string, string>, name: string) => {
                const file = path.join(dir, `${name}.docx`);
                await fs.writeFile(file, await buildDocx(parts));
                return validate(file, { profile: "word-valid" });
            };
            // The refused rows: session 66 (s66a) and session 67 (s67a),
            // each valid before this rule, each now carrying exactly this
            // error beside its warnings.
            for (const [parts, name] of [
                [commentsShared(true, true, true, true), "comments-both-reaches"],
                [notesShared("footnote", true, true, true), "footnotes-shared-both-reaches"],
                [notesShared("endnote", true, true, true), "endnotes-shared-both-reaches"],
            ] as const) {
                const result = await wordValid(parts, name);
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code)).toEqual([
                    "word-shared-story-part",
                ]);
            }
            // The clean rows: single-root reads and relate-only shares.
            for (const [parts, name] of [
                [commentsShared(true, true, false, false), "comments-main-only"],
                [commentsShared(false, false, true, true), "comments-glossary-only"],
                [commentsShared(true, true, true, false), "comments-both-relate-main-references"],
                [commentsShared(true, false, true, true), "comments-both-relate-glossary-references"],
                [commentsShared(true, false, true, false), "comments-both-relate-neither-references"],
                [notesShared("footnote", true, true, false), "footnotes-shared-main-references-calibration"],
            ] as const) {
                expect((await wordValid(parts, name)).valid).toBe(true);
            }
            // A glossary whose usage is unknown is rejected for its own
            // XML, never as a shared part (#858).
            const unknownGlossary = commentsShared(true, true, true, false);
            unknownGlossary[GLOSSARY] = replaceOnce(unknownGlossary[GLOSSARY]!, 'encoding="UTF-8"', 'encoding="bogus"');
            const result = await wordValid(unknownGlossary, "comments-unknown-glossary-usage");
            expect(result.valid).toBe(false);
            expect(result.issues.map((issue) => issue.code)).not.toContain("word-shared-story-part");
        });
    });
});

// #996, Word sessions 81 (s81a/s81b, two runs each, all readings agree): a
// glossary reached by a climbing glossaryDocument target, or relating the
// notes part by a climbing target, reads the notes part the main document
// reads, as every read-model walk stops the climb at the root since #986.
// Word offered recovery when both roots used the note, as session 67 did
// in-package, and opened the climbing glossary that uses no note clean.
describe("a notes part two roots use through climbing targets (#996, session 81)", { timeout: 30000 }, () => {
    const judged = (parts: Record<string, string>) =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "climbing-shared-notes.docx");
            await fs.writeFile(file, await buildDocx(parts));
            return validate(file, { profile: "word-valid" });
        });

    it.each([
        ["the glossary reached by ../../word/glossary/document.xml", "../../word/glossary/document.xml", "../footnotes.xml"],
        ["the glossary's notes related by ../../../word/footnotes.xml", "glossary/document.xml", "../../../word/footnotes.xml"],
    ])("blocks word/footnotes.xml both roots use, %s, as Word refused it", async (_row, glossaryTarget, notesTarget) => {
        const result = await judged(climbingSharedNotes(glossaryTarget, notesTarget, true));
        expect(result.valid).toBe(false);
        expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => [issue.code, issue.path])).toEqual([
            ["word-shared-story-part", "word/footnotes.xml"],
        ]);
    });

    it("accepts the climbing glossary that uses no note, which Word opened clean", async () => {
        const result = await judged(climbingSharedNotes("../../word/glossary/document.xml", "../footnotes.xml", false));
        expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        expect(result.valid).toBe(true);
    });
});
