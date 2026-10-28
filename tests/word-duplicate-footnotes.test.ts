/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Session 63 (s63d; two runs, all readings agree) specimens for #827:
 * Word offered recovery whenever the main document's relationships part
 * held two exact-Type footnotes relationships and the body referenced a
 * footnote — whether the second part held another id or the same
 * footnote 1 — while two unreferenced footnotes relationships opened
 * clean (session 45 s45g). A reached glossary's own pair is unprobed, so
 * the rule judges only the main document's sidecar. The rule counts
 * exact Types whatever their target or TargetMode, in either
 * relationship namespace, so the same-part, internal-plus-External and
 * Strict-Type pairs below are judged too; no session showed them, so
 * they are unprobed. Specimens are built from MINIMAL_PARTS the way
 * the session 63 builder built its rows.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { R, RS } from "./support/word-rule-test-xml";

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
const relationship = (id: string, segment: string, target: string, extra = ""): string =>
    `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"${extra}/>`;
/** A footnotes relationship whose exact Type sits in the Strict relationship-type namespace. */
const strictRelationship = (id: string, target: string): string => `<Relationship Id="${id}" Type="${RS}/footnotes" Target="${target}"/>`;
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

const footnotesXml = (held: string): string =>
    `${XML_DECL}<w:footnotes xmlns:w="${W}"><w:footnote w:id="${held}">${TEXT_P}</w:footnote></w:footnotes>`;
const footnoteRef = (id: string): string => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
/** The minimal parts with `word/footnotes.xml` holding footnote `held`, related by the main document and typed footnotes+xml. */
const withFootnotes = (held: string): Record<string, string> => ({
    ...P,
    "word/footnotes.xml": footnotesXml(held),
    [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdN", "footnotes", "footnotes.xml")),
    [CT]: overrides(P[CT]!, ["word/footnotes.xml", `${WML}.footnotes+xml`]),
});

/** Two footnotes relationships in document order: word/footnotes.xml holding `first`, then word/notes2.xml (footnotes+xml) holding `second`. */
const twoFootnotes = (first: string, second: string, referenced: boolean): Promise<Buffer> => {
    const parts = withFootnotes(first);
    const both = {
        ...parts,
        "word/notes2.xml": footnotesXml(second),
        [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdN2", "footnotes", "notes2.xml")),
        [CT]: overrides(parts[CT]!, ["word/notes2.xml", `${WML}.footnotes+xml`]),
    };
    return buildDocx(referenced ? mainWith(both, footnoteRef("1")) : both);
};

/** Two footnotes relationships naming the same `word/footnotes.xml`, the body referencing footnote 1: the unprobed same-part shape. */
const samePartPair = (): Promise<Buffer> => {
    const parts = withFootnotes("1");
    return buildDocx(
        mainWith(
            { ...parts, [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdN2", "footnotes", "footnotes.xml")) },
            footnoteRef("1"),
        ),
    );
};

/** An internal footnotes relationship beside an External one, the body referencing footnote 1: the unprobed TargetMode shape. */
const internalExternalPair = (): Promise<Buffer> => {
    const parts = withFootnotes("1");
    return buildDocx(
        mainWith(
            {
                ...parts,
                [SIDECAR]: appendRels(
                    parts[SIDECAR]!,
                    relationship("rIdN2", "footnotes", "https://example.invalid/footnotes.xml", ' TargetMode="External"'),
                ),
            },
            footnoteRef("1"),
        ),
    );
};

/** Two footnotes relationships typed in the Strict relationship-type namespace, to two parts; the body references footnote 1. */
const strictTypePair = (): Promise<Buffer> =>
    buildDocx(
        mainWith(
            {
                ...P,
                "word/footnotes.xml": footnotesXml("1"),
                "word/notes2.xml": footnotesXml("5"),
                [SIDECAR]: appendRels(
                    P[SIDECAR]!,
                    `${strictRelationship("rIdN", "footnotes.xml")}${strictRelationship("rIdN2", "notes2.xml")}`,
                ),
                [CT]: overrides(P[CT]!, ["word/footnotes.xml", `${WML}.footnotes+xml`], ["word/notes2.xml", `${WML}.footnotes+xml`]),
            },
            footnoteRef("1"),
        ),
    );

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
    `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r>${footnoteRef("1")}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
/** A reached glossary whose own sidecar holds two footnotes relationships, and whose body references a footnote: the unprobed scope. */
const glossaryPair = (): Promise<Buffer> =>
    buildDocx({
        ...P,
        "word/glossary/document.xml": GLOSSARY,
        "word/glossary/footnotes.xml": footnotesXml("1"),
        "word/glossary/footnotes2.xml": footnotesXml("2"),
        "word/glossary/_rels/document.xml.rels": `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship(
            "rIdN",
            "footnotes",
            "footnotes.xml",
        )}${relationship("rIdN2", "footnotes", "footnotes2.xml")}</Relationships>`,
        [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdG", "glossaryDocument", "glossary/document.xml")),
        [CT]: overrides(
            P[CT]!,
            ["word/glossary/document.xml", `${WML}.document.glossary+xml`],
            ["word/glossary/footnotes.xml", `${WML}.footnotes+xml`],
            ["word/glossary/footnotes2.xml", `${WML}.footnotes+xml`],
        ),
    });

const write = async (dir: string, name: string, make: () => Promise<Buffer>): Promise<string> => {
    const file = path.join(dir, `${name}.docx`);
    writeFileSync(file, await make());
    return file;
};

const duplicateFootnotesFinding = (): ValidationIssue => ({
    severity: "error",
    code: "word-duplicate-footnotes-relationship",
    path: "word/_rels/document.xml.rels",
    message:
        "word/_rels/document.xml.rels contains 2 footnotes relationships (rIdN, rIdN2) and the main document references a footnote. " +
        "Microsoft Word offers to recover the file. Keep one footnotes relationship.",
});

const errorIssues = (result: Awaited<ReturnType<typeof validate>>): ValidationIssue[] =>
    result.issues.filter((issue) => issue.severity === "error");

describe("word-valid duplicate footnotes relationship (#827; session 63)", { timeout: 60000 }, () => {
    it("registers the check in WORD_RULE_CHECKS", () => {
        expect(WORD_RULE_CHECKS.some((rule) => rule.id === "word-duplicate-footnotes-relationship")).toBe(true);
    });

    it.each([
        ["s63d two-footnotes-relationships-first-holds-1", () => twoFootnotes("1", "5", true)],
        ["s63d two-footnotes-relationships-both-hold-1", () => twoFootnotes("1", "1", true)],
    ])("rejects %s with only the duplicate finding", async (name, make) => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, name, make), { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorIssues(result)).toEqual([duplicateFootnotesFinding()]);
        });
    });

    // No session showed these shapes; the rule counts exact Types whatever
    // the target or TargetMode and in either relationship namespace.
    it.each([
        ["two relationships to the same part", samePartPair],
        ["an internal and an External relationship", internalExternalPair],
        ["two Strict-Type relationships", strictTypePair],
    ])("rejects %s with only the duplicate finding (unprobed)", async (name, make) => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, `footnotes-pair-${name}`, make), { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorIssues(result)).toEqual([duplicateFootnotesFinding()]);
        });
    });

    it("reports the pair in a selected word-duplicate-footnotes-relationship run", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "two-footnotes-first-holds-1", () => twoFootnotes("1", "5", true)), {
                profile: "word-valid",
                checks: ["word-duplicate-footnotes-relationship"],
            }).catch((error: unknown) => error);
            expect(result, "the check id is registered for selection").not.toBeInstanceOf(Error);
            const selected = result as Awaited<ReturnType<typeof validate>>;
            expect(selected.valid).toBe(false);
            expect(selected.issues).toEqual([duplicateFootnotesFinding()]);
        });
    });

    it("passes session 45's unreferenced footnotes pair (s45g)", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "footnotes-dup", footnotesDup), { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.some((issue) => issue.code === "word-duplicate-footnotes-relationship")).toBe(false);
        });
    });

    it("passes one footnotes relationship with a referenced, defined footnote", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(
                await write(dir, "single-referenced-footnote", () => buildDocx(mainWith(withFootnotes("1"), footnoteRef("1")))),
                { profile: "word-valid" },
            );
            expect(result.valid).toBe(true);
            expect(result.issues.some((issue) => issue.code === "word-duplicate-footnotes-relationship")).toBe(false);
        });
    });

    it("reports nothing for a glossary's own pair with a body footnote reference in the glossary (unprobed)", async () => {
        await withTempDir(async (dir) => {
            const result = await validate(await write(dir, "glossary-footnotes-pair", glossaryPair), { profile: "word-valid" });
            expect(result.issues.some((issue) => issue.code === "word-duplicate-footnotes-relationship")).toBe(false);
        });
    });
});
