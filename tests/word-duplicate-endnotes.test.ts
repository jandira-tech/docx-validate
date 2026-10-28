/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Session 62 (s62c, s62e; two runs, all readings agree) specimens for
 * #807: Word offered recovery whenever the main document's relationships
 * part held two exact-Type endnotes relationships and the body referenced
 * an endnote — whichever part held the note (the first, the second, or
 * both), and whether the second part was typed endnotes+xml or
 * application/octet-stream. The same two relationships with no endnote
 * reference opened clean (s62e). A reached glossary's own pair is
 * unprobed, so the rule judges only the main document's sidecar. Two
 * unreferenced footnotes relationships opened clean (session 45 s45g);
 * the referenced pair, which session 63 s63d showed Word refusing, has
 * its own rule (#827). The octet-stream row also draws
 * `word-endnote-reference-missing`: such a part defines no note whatever
 * its name (session 64 s64b, #829), so the body reference to the note only
 * it holds dangles. Specimens are built from MINIMAL_PARTS the way
 * the session 62 builder built its rows.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { R } from "./support/word-rule-test-xml";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const U = "urn:example:unknown";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CT = "[Content_Types].xml";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const P = MINIMAL_PARTS as Record<string, string>;
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const TEXT_P = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>";

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

/**
 * `parts` with the main document's `r:` and `u:` prefixes declared and
 * `run` appended inside its only paragraph — the builder's `mainWith`,
 * without the sectPr injection no row here needs.
 */
const mainWith = (parts: Record<string, string>, run: string): Record<string, string> => ({
    ...parts,
    [MAIN]: replaceOnce(
        replaceOnce(parts[MAIN]!, `xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}" xmlns:u="${U}"`),
        "</w:r></w:p>",
        `</w:r>${run}</w:p>`,
    ),
});

type Kind = "footnote" | "endnote";
const notesXml = (kind: Kind, held: string): string =>
    `${XML_DECL}<w:${kind}s xmlns:w="${W}"><w:${kind} w:id="${held}">${TEXT_P}</w:${kind}></w:${kind}s>`;
const noteRef = (kind: Kind, id: string): string => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
/** The minimal parts with `word/<kind>s.xml` holding note `held`, related by the main document and typed by `type`. */
const withNotes = (kind: Kind, held: string, type = `${WML}.${kind}s+xml`): Record<string, string> => ({
    ...P,
    [`word/${kind}s.xml`]: notesXml(kind, held),
    [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdN", `${kind}s`, `${kind}s.xml`)),
    [CT]: overrides(P[CT]!, [`word/${kind}s.xml`, type]),
});

/** s62c `endnotes-second-relationship-octet-stream`: the related endnotes.xml holds 5, a second endnotes relationship names word/notes.bin (Override application/octet-stream) holding 1; the body references 1. */
const secondRelationshipOctetStream = (): Promise<Buffer> => {
    const parts = withNotes("endnote", "5");
    return buildDocx(
        mainWith(
            {
                ...parts,
                "word/notes.bin": notesXml("endnote", "1"),
                [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdN2", "endnotes", "notes.bin")),
                [CT]: overrides(parts[CT]!, ["word/notes.bin", "application/octet-stream"]),
            },
            noteRef("endnote", "1"),
        ),
    );
};
/** s62c `endnotes-second-relationship-xml-calibration`: the same with the second part word/notes2.xml typed endnotes+xml. */
const secondRelationshipXmlCalibration = (): Promise<Buffer> => {
    const parts = withNotes("endnote", "5");
    return buildDocx(
        mainWith(
            {
                ...parts,
                "word/notes2.xml": notesXml("endnote", "1"),
                [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdN2", "endnotes", "notes2.xml")),
                [CT]: overrides(parts[CT]!, ["word/notes2.xml", `${WML}.endnotes+xml`]),
            },
            noteRef("endnote", "1"),
        ),
    );
};
/** Two endnotes relationships in document order: word/endnotes.xml holding `first`, then word/notes2.xml (endnotes+xml) holding `second`. */
const twoEndnotes = (first: string, second: string, referenced: boolean): Promise<Buffer> => {
    const parts = withNotes("endnote", first);
    const both = {
        ...parts,
        "word/notes2.xml": notesXml("endnote", second),
        [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdN2", "endnotes", "notes2.xml")),
        [CT]: overrides(parts[CT]!, ["word/notes2.xml", `${WML}.endnotes+xml`]),
    };
    return buildDocx(referenced ? mainWith(both, noteRef("endnote", "1")) : both);
};

/** Session 45's `footnotes-dup` shape: two footnotes relationships to two parts, no reference (s45g). */
const footnotesDup = (): Promise<Buffer> => {
    const footnotes = `${XML_DECL}<w:footnotes xmlns:w="${W}"/>`;
    return buildDocx({
        ...P,
        "word/footnotes.xml": footnotes,
        "word/footnotes2.xml": footnotes,
        [SIDECAR]: appendRels(
            P[SIDECAR]!,
            `${relationship("rId2", "footnotes", "footnotes.xml")}${relationship("rId3", "footnotes", "footnotes2.xml")}`,
        ),
        [CT]: overrides(P[CT]!, ["word/footnotes.xml", `${WML}.footnotes+xml`], ["word/footnotes2.xml", `${WML}.footnotes+xml`]),
    });
};

const GLOSSARY =
    `${XML_DECL}<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r>${noteRef("endnote", "1")}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
/** A reached glossary whose own sidecar holds two endnotes relationships, and whose body references an endnote: the unprobed scope. */
const glossaryPair = (): Promise<Buffer> =>
    buildDocx({
        ...P,
        "word/glossary/document.xml": GLOSSARY,
        "word/glossary/endnotes.xml": notesXml("endnote", "1"),
        "word/glossary/endnotes2.xml": notesXml("endnote", "2"),
        "word/glossary/_rels/document.xml.rels": `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship(
            "rIdN",
            "endnotes",
            "endnotes.xml",
        )}${relationship("rIdN2", "endnotes", "endnotes2.xml")}</Relationships>`,
        [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdG", "glossaryDocument", "glossary/document.xml")),
        [CT]: overrides(
            P[CT]!,
            ["word/glossary/document.xml", `${WML}.document.glossary+xml`],
            ["word/glossary/endnotes.xml", `${WML}.endnotes+xml`],
            ["word/glossary/endnotes2.xml", `${WML}.endnotes+xml`],
        ),
    });

const write = async (dir: string, name: string, make: () => Promise<Buffer>): Promise<string> => {
    const file = path.join(dir, `${name}.docx`);
    writeFileSync(file, await make());
    return file;
};

const duplicateEndnotesFinding = (): ValidationIssue => ({
    severity: "error",
    code: "word-duplicate-endnotes-relationship",
    path: "word/_rels/document.xml.rels",
    message:
        "word/_rels/document.xml.rels contains 2 endnotes relationships (rIdN, rIdN2) and the main document references an endnote. " +
        "Microsoft Word offers to recover the file. Keep one endnotes relationship.",
});

const errorIssues = (result: Awaited<ReturnType<typeof validate>>): ValidationIssue[] =>
    result.issues.filter((issue) => issue.severity === "error");

describe("word-valid duplicate endnotes relationship (#807; session 62)", { timeout: 60000 }, () => {
    it("registers the check in WORD_RULE_CHECKS", () => {
        expect(WORD_RULE_CHECKS.some((rule) => rule.id === "word-duplicate-endnotes-relationship")).toBe(true);
    });

    it.each([
        ["s62c endnotes-second-relationship-xml-calibration", secondRelationshipXmlCalibration],
        ["s62e two-endnotes-relationships-first-holds-1", () => twoEndnotes("1", "5", true)],
        ["s62e two-endnotes-relationships-both-hold-1", () => twoEndnotes("1", "1", true)],
    ])("rejects %s with only the duplicate finding", async (name, make) => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, name, make), { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorIssues(result)).toEqual([duplicateEndnotesFinding()]);
        });
    });

    it("rejects s62c's octet-stream second part with the duplicate finding and the dangling reference", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(
                await write(dir, "s62c endnotes-second-relationship-octet-stream", secondRelationshipOctetStream),
                {
                    profile: "word-valid",
                },
            );
            expect(result.valid).toBe(false);
            // The note only the octet-stream part holds defines nothing
            // (session 64 s64b, #829), so the body reference dangles too.
            expect(errorIssues(result)).toEqual([
                {
                    severity: "error",
                    code: "word-endnote-reference-missing",
                    path: "word/document.xml",
                    message:
                        "An endnote reference uses id 1, and this package has no w:endnote with that id. " +
                        "Microsoft Word offers to recover the file. " +
                        "Remove the w:endnoteReference, or add the matching endnote.",
                },
                duplicateEndnotesFinding(),
            ]);
        });
    });

    it("reports the pair in a selected word-duplicate-endnotes-relationship run", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "two-endnotes-first-holds-1", () => twoEndnotes("1", "5", true)), {
                profile: "word-valid",
                checks: ["word-duplicate-endnotes-relationship"],
            }).catch((error: unknown) => error);
            expect(result, "the check id is registered for selection").not.toBeInstanceOf(Error);
            const selected = result as Awaited<ReturnType<typeof validate>>;
            expect(selected.valid).toBe(false);
            expect(selected.issues).toEqual([duplicateEndnotesFinding()]);
        });
    });

    it("passes s62e's no-reference row and reports nothing", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "two-endnotes-relationships-no-reference", () => twoEndnotes("5", "1", false)), {
                profile: "word-valid",
            });
            expect(result.valid).toBe(true);
            expect(result.issues.some((issue) => issue.code === "word-duplicate-endnotes-relationship")).toBe(false);
        });
    });

    it("passes one endnotes relationship with a referenced, defined endnote", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(
                await write(dir, "single-referenced-endnote", () =>
                    buildDocx(mainWith(withNotes("endnote", "1"), noteRef("endnote", "1"))),
                ),
                { profile: "word-valid" },
            );
            expect(result.valid).toBe(true);
            expect(result.issues.some((issue) => issue.code === "word-duplicate-endnotes-relationship")).toBe(false);
        });
    });

    it("reports nothing for session 45's unreferenced footnotes pair (s45g)", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "footnotes-dup", footnotesDup), { profile: "word-valid" });
            expect(result.issues.some((issue) => issue.code === "word-duplicate-endnotes-relationship")).toBe(false);
        });
    });

    it("reports nothing for a glossary's own pair with a body endnote reference in the glossary (unprobed)", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "glossary-endnotes-pair", glossaryPair), { profile: "word-valid" });
            expect(result.issues.some((issue) => issue.code === "word-duplicate-endnotes-relationship")).toBe(false);
        });
    });
});
