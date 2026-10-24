/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { wordHyperlinkRelationshipCheck } from "../src/scripts/office/checks/word-hyperlink";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { document, R, rels, relation, RS, W } from "./support/word-rule-test-xml";

describe("word-hyperlink-relationship", () => {
    it("does not resolve a hyperlink through a nested relationship (#434)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document('<w:p><w:hyperlink r:id="nested"/></w:p>'),
            "word/_rels/document.xml.rels": rels(`<x:Wrap xmlns:x="urn:x">${relation("nested", "hyperlink", "External")}</x:Wrap>`),
        });
        expect(await wordHyperlinkRelationshipCheck.run(ctx)).toEqual([
            expect.objectContaining({ code: "word-hyperlink-relationship", path: "word/document.xml", severity: "error" }),
        ]);
    });

    it.each(["RELS", "Rels", "%72els"])("resolves an OPC-equivalent .%s sidecar (#402)", async (extension) => {
        const entries = {
            "word/document.xml": document('<w:p><w:hyperlink r:id="rId2"/></w:p>'),
            [`word/_rels/document.xml.${extension}`]: rels(relation("rId2", "hyperlink", "External")),
        };
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries(entries))).toEqual([]);
    });

    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<w:p><w:hyperlink r:id="rId1"/></w:p>') });
        expect(await wordHyperlinkRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-hyperlink-relationship",
                path: "word/document.xml",
                message:
                    `<w:hyperlink> r:id="rId1" does not reference an external hyperlink relationship (no Relationship with Id 'rId1'). ` +
                    "Microsoft Word offers to recover the file. " +
                    `Use a Relationship of Type ${R}/hyperlink (or the ISO Strict ${RS}/hyperlink) with TargetMode External, or remove the r:id.`,
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:hyperlink r:id="rId1"/><w:hyperlink/></w:p>`),
            "word/_rels/document.xml.rels": rels(relation("rId1", "hyperlink", "External")),
        });
        expect(await wordHyperlinkRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordHyperlinkRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it.each([
    ["image", "External", "Type"],
    ["hyperlink", "", "TargetMode"],
])("explains an incompatible relationship %s/%s", async (type, mode, reason) => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink r:id="rId1"/>`),
        "word/_rels/document.xml.rels": rels(relation("rId1", type, mode)),
    });
    expect((await wordHyperlinkRelationshipCheck.run(ctx))[0]?.message).toContain(reason);
});
it("reads strict relationship ids, missing attributes, and unparseable sidecars", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink rs:id="rId1"/><w:hyperlink/>`),
        "word/_rels/document.xml.rels": rels(`<Relationship/><Relationship Id="rId1"/>`),
    });
    expect((await wordHyperlinkRelationshipCheck.run(ctx))[0]?.message).toContain("Type '(none)'");
    const invalid = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink r:id="rId1"/>`),
        "word/_rels/document.xml.rels": "<Relationships",
    });
    expect(await wordHyperlinkRelationshipCheck.run(invalid)).toEqual([]);
});

/**
 * Session 28's #614 shape (chunk s28a in
 * docs/plans/evidence/word-probes-2026-10-24.md): `MINIMAL_PARTS` with a body
 * `w:hyperlink r:id="rIdH"` and an External relationship to
 * `https://example.org/`, only the relationship Type varied.
 */
const hyperlinkRelTypeParts = (type: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
        `<w:document xmlns:w="${W}"`,
        `<w:document xmlns:w="${W}" xmlns:r="${R}"`,
    ).replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        `<w:r><w:t>Hello</w:t></w:r><w:hyperlink r:id="rIdH"><w:r><w:t>link</w:t></w:r></w:hyperlink>`,
    ),
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="rIdH" Type="${type}" Target="https://example.org/" TargetMode="External"/></Relationships>`,
    ),
});

/**
 * Session 28's #614 refused Types (repair-warning in Word 16.115): the two
 * the old ends-in test wrongly passed, then the one it already refused.
 */
const REFUSED_REL_TYPES = [
    { name: "hyperlink-rel-type-example-com", type: "http://example.com/hyperlink" },
    { name: "hyperlink-rel-type-custom-hyperlink", type: `${R}/custom/hyperlink` },
    { name: "hyperlink-rel-type-upper", type: `${R}/HYPERLINK` },
] as const;

/** Session 28's #614 controls (clean-open in Word 16.115): the only accepted Types, byte for byte. */
const CLEAN_REL_TYPES = [
    { name: "hyperlink-rel-type-strict-uri", type: `${RS}/hyperlink` },
    { name: "hyperlink-rel-type-transitional", type: `${R}/hyperlink` },
] as const;

describe("word-hyperlink-relationship accepts exactly two Types (#614, Word 16.115 session 28)", { timeout: 240_000 }, () => {
    describe.each(REFUSED_REL_TYPES)("refused Type specimen %s", (specimen) => {
        it("word-valid fails with exactly one Word-blocking error naming the r:id and Type", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${specimen.name}.docx`);
                writeFileSync(file, await buildDocx(hyperlinkRelTypeParts(specimen.type)));
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === "word-hyperlink-relationship");
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/document.xml");
                expect(found[0]?.message).toContain("rIdH");
                expect(found[0]?.message).toContain(specimen.type);
            });
        });

        it("strict reports the code too (strict ⊇ word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${specimen.name}.docx`);
                writeFileSync(file, await buildDocx(hyperlinkRelTypeParts(specimen.type)));
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid).toBe(false);
                expect(
                    strict.issues.some((issue) => issue.code === "word-hyperlink-relationship" && issue.path === "word/document.xml"),
                ).toBe(true);
            });
        });

        it("lenient reports nothing for it", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${specimen.name}.docx`);
                writeFileSync(file, await buildDocx(hyperlinkRelTypeParts(specimen.type)));
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
            });
        });
    });

    describe.each(CLEAN_REL_TYPES)("clean Type specimen %s", (specimen) => {
        it("word-valid passes with no new finding", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${specimen.name}.docx`);
                writeFileSync(file, await buildDocx(hyperlinkRelTypeParts(specimen.type)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship")).toEqual([]);
            });
        });
    });
});
