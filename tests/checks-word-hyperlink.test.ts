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

describe("hyperlink TargetMode scope (#773)", () => {
    it.each(["external", "EXTERNAL", "External ", ""])("leaves a present %j to the main sidecar's target-mode check", async (mode) => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            'TargetMode="External"',
            `TargetMode="${mode}"`,
        );
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it("shares the relationship-resolved scope with an OPC-equivalent sidecar", async () => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["word/_rels/document.xml.RELS"] = parts["word/_rels/document.xml.rels"]!.replace(
            'TargetMode="External"',
            'TargetMode="external"',
        );
        delete parts["word/_rels/document.xml.rels"];
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it("leaves a reached glossary's misspelled mode to the target-mode check (#775)", async () => {
        const parts = { ...MINIMAL_PARTS };
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
        );
        parts["word/glossary/document.xml"] =
            `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartBody><w:p>` +
            '<w:hyperlink r:id="rId2"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
            "</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>";
        parts["word/glossary/_rels/document.xml.rels"] = rels(relation("rId2", "hyperlink", "external"));
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it("keeps the hyperlink finding when the package does not name its owner as a main document", async () => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            'TargetMode="External"',
            'TargetMode="external"',
        );
        delete parts["_rels/.rels"];
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries(parts))).toEqual([
            expect.objectContaining({ code: "word-hyperlink-relationship", severity: "error", path: "word/document.xml" }),
        ]);
    });
});

/**
 * #1089: a full word-valid or strict run reports a `w:hyperlink` the rule
 * refuses once, through the rule. The legacy walk's `rels-id-mismatch` or
 * `rels-id-missing` copy of the same element is handed off (#616), for an id
 * in either relationships namespace (#1065, #1079). Lenient runs no Word
 * rule and keeps the legacy finding.
 */
describe("a refused w:hyperlink is one report in a full run (#1089)", () => {
    const LEGACY = new Set(["rels-id-mismatch", "rels-id-missing"]);
    /** Session 28's shape with the hyperlink's ids replaced by `ids`; `rId1` is MINIMAL_PARTS' styles relationship. */
    const withIds = (ids: string): Record<string, string> => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["word/document.xml"] = parts["word/document.xml"]!.replace(
            `<w:document xmlns:w="${W}"`,
            `<w:document xmlns:rs="${RS}" xmlns:w="${W}"`,
        ).replace('<w:hyperlink r:id="rIdH">', `<w:hyperlink ${ids}>`);
        return parts;
    };
    it.each([
        ["word-valid", 'r:id="rId1"'],
        ["strict", 'r:id="rId1"'],
        ["word-valid", 'r:id="rIdX"'],
        ["strict", 'r:id="rIdX"'],
        ["word-valid", 'rs:id="rIdY"'],
        ["strict", 'rs:id="rIdY"'],
        ["word-valid", 'r:id="rIdX" rs:id="rIdY"'],
        ["strict", 'r:id="rIdX" rs:id="rIdY"'],
    ] as const)("reports it once under %s for %s", async (profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "refused-hyperlink.docx");
            writeFileSync(file, await buildDocx(withIds(ids)));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship")).toHaveLength(1);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? ""))).toEqual([]);
        });
    });

    it.each([
        ['r:id="rId1"', "rels-id-mismatch"],
        ['r:id="rIdX"', "rels-id-missing"],
    ] as const)("lenient keeps the legacy finding for %s", async (ids, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "refused-hyperlink.docx");
            writeFileSync(file, await buildDocx(withIds(ids)));
            const result = await validate(file, { profile: "lenient" });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship")).toEqual([]);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? "")).map((issue) => issue.code)).toEqual([code]);
        });
    });
});

/**
 * #1099: the #1089 hand-off covers every story part the rule judges, not
 * only the main document. Session 28's shape gains a read header
 * (`word/header1.xml`, related and referenced by the body `w:sectPr`) whose
 * hyperlink names nothing its own sidecar defines.
 */
describe("a refused w:hyperlink in a read header is one report in a full run (#1099)", () => {
    const LEGACY = new Set(["rels-id-mismatch", "rels-id-missing"]);
    const withHeaderHyperlink = (ids: string): Record<string, string> => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
        );
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdHdr" Type="${R}/header" Target="header1.xml"/></Relationships>`,
        );
        parts["word/document.xml"] = parts["word/document.xml"]!.replace(
            "<w:sectPr>",
            `<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>`,
        );
        parts["word/header1.xml"] =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:rs="${RS}">` +
            `<w:p><w:hyperlink ${ids}><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:hdr>`;
        parts["word/_rels/header1.xml.rels"] = rels(relation("rIdL", "hyperlink", "External"));
        return parts;
    };

    it.each([
        ["word-valid", 'r:id="rIdX"'],
        ["strict", 'r:id="rIdX"'],
        ["word-valid", 'rs:id="rIdY"'],
        ["strict", 'rs:id="rIdY"'],
    ] as const)("reports it once under %s for %s", async (profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "header-hyperlink.docx");
            writeFileSync(file, await buildDocx(withHeaderHyperlink(ids)));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship").map((issue) => issue.path)).toEqual([
                "word/header1.xml",
            ]);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? ""))).toEqual([]);
        });
    });

    it("lenient keeps the legacy finding", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "header-hyperlink.docx");
            writeFileSync(file, await buildDocx(withHeaderHyperlink('r:id="rIdX"')));
            const result = await validate(file, { profile: "lenient" });
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? "")).map((issue) => `${issue.code} ${issue.path}`)).toEqual([
                "rels-id-missing word/header1.xml",
            ]);
        });
    });
});

/**
 * #1110: the same hand-off for every other story kind the rule judges. Since
 * #1106 the rule judges only the parts Word reads, so each part is made one:
 * a footer the body `w:sectPr` references, a comments part a
 * `w:commentReference` uses, and footnotes and endnotes parts a note
 * reference uses. Each hyperlink names nothing its part's own sidecar defines.
 */
describe("a refused w:hyperlink in a read footer, comments or notes part is one report in a full run (#1110)", () => {
    const LEGACY = new Set(["rels-id-mismatch", "rels-id-missing"]);
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:rs="${RS}"`;
    /** Each story kind: its part, relationship segment, the body or section markup that uses it, and its root around `p`. */
    const STORIES = {
        footer: {
            part: "word/footer1.xml",
            segment: "footer",
            body: "",
            section: '<w:footerReference w:type="default" r:id="rIdS"/>',
            wrap: (p: string) => `<w:ftr ${NS}>${p}</w:ftr>`,
        },
        comments: {
            part: "word/comments.xml",
            segment: "comments",
            body: '<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>',
            section: "",
            wrap: (p: string) => `<w:comments ${NS}><w:comment w:id="0" w:author="A">${p}</w:comment></w:comments>`,
        },
        footnotes: {
            part: "word/footnotes.xml",
            segment: "footnotes",
            body: '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
            section: "",
            wrap: (p: string) => `<w:footnotes ${NS}><w:footnote w:id="1">${p}</w:footnote></w:footnotes>`,
        },
        endnotes: {
            part: "word/endnotes.xml",
            segment: "endnotes",
            body: '<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>',
            section: "",
            wrap: (p: string) => `<w:endnotes ${NS}><w:endnote w:id="1">${p}</w:endnote></w:endnotes>`,
        },
    } as const;
    type Story = keyof typeof STORIES;
    /** Session 28's accepted main document plus `story`, read by Word, whose hyperlink carries `ids`. */
    const withStoryHyperlink = (story: Story, ids: string): Record<string, string> => {
        const { part, segment, body, section, wrap } = STORIES[story];
        const name = path.posix.basename(part);
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/${part}" ContentType="${WML}.${segment}+xml"/></Types>`,
        );
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdS" Type="${R}/${segment}" Target="${name}"/></Relationships>`,
        );
        parts["word/document.xml"] = parts["word/document.xml"]!.replace("<w:sectPr>", `${body}<w:sectPr>${section}`);
        parts[part] =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            wrap(`<w:p><w:hyperlink ${ids}><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`);
        parts[`word/_rels/${name}.rels`] = rels(relation("rIdL", "hyperlink", "External"));
        return parts;
    };
    const STORY_ROWS = (Object.keys(STORIES) as Story[]).flatMap((story) =>
        (["word-valid", "strict"] as const).flatMap((profile) =>
            ['r:id="rIdX"', 'rs:id="rIdY"'].map((ids) => [story, profile, ids] as const),
        ),
    );

    it.each(STORY_ROWS)("a %s part Word reads: reports it once under %s for %s", async (story, profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "story-hyperlink.docx");
            writeFileSync(file, await buildDocx(withStoryHyperlink(story, ids)));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship").map((issue) => issue.path)).toEqual([
                STORIES[story].part,
            ]);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? ""))).toEqual([]);
        });
    });

    it.each(Object.keys(STORIES) as Story[])("a %s part Word reads: lenient keeps the legacy finding", async (story) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "story-hyperlink.docx");
            writeFileSync(file, await buildDocx(withStoryHyperlink(story, 'r:id="rIdX"')));
            const result = await validate(file, { profile: "lenient" });
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? "")).map((issue) => `${issue.code} ${issue.path}`)).toEqual([
                `rels-id-missing ${STORIES[story].part}`,
            ]);
        });
    });
});

/**
 * #1129: the same hand-off in a reached glossary document, the other story
 * root the rule's read set holds. Session 28's accepted main document gains a
 * `glossaryDocument` relationship, and the glossary's `w:docPartBody` holds a
 * hyperlink naming nothing its own sidecar defines.
 */
describe("a refused w:hyperlink in a reached glossary is one report in a full run (#1129)", () => {
    const LEGACY = new Set(["rels-id-mismatch", "rels-id-missing"]);
    const GLOSSARY = "word/glossary/document.xml";
    const withGlossaryHyperlink = (ids: string): Record<string, string> => {
        const parts = hyperlinkRelTypeParts(`${R}/hyperlink`);
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/${GLOSSARY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>`,
        );
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
        );
        parts[GLOSSARY] =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}" xmlns:rs="${RS}"><w:docParts><w:docPart>` +
            `<w:docPartPr><w:name w:val="Part"/><w:guid w:val="{00000000-0000-0000-0000-000000000001}"/></w:docPartPr>` +
            `<w:docPartBody><w:p><w:hyperlink ${ids}><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:docPartBody>` +
            `</w:docPart></w:docParts></w:glossaryDocument>`;
        parts["word/glossary/_rels/document.xml.rels"] = rels(relation("rIdL", "hyperlink", "External"));
        return parts;
    };

    it.each([
        ["word-valid", 'r:id="rIdX"'],
        ["strict", 'r:id="rIdX"'],
        ["word-valid", 'rs:id="rIdY"'],
        ["strict", 'rs:id="rIdY"'],
        ["word-valid", 'r:id="rIdX" rs:id="rIdY"'],
        ["strict", 'r:id="rIdX" rs:id="rIdY"'],
    ] as const)("reports it once under %s for %s", async (profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "glossary-hyperlink.docx");
            writeFileSync(file, await buildDocx(withGlossaryHyperlink(ids)));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship").map((issue) => issue.path)).toEqual([
                GLOSSARY,
            ]);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? ""))).toEqual([]);
        });
    });

    it("lenient keeps the legacy finding", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "glossary-hyperlink.docx");
            writeFileSync(file, await buildDocx(withGlossaryHyperlink('r:id="rIdX"')));
            const result = await validate(file, { profile: "lenient" });
            expect(result.issues.filter((issue) => issue.code === "word-hyperlink-relationship")).toEqual([]);
            expect(result.issues.filter((issue) => LEGACY.has(issue.code ?? "")).map((issue) => `${issue.code} ${issue.path}`)).toEqual([
                `rels-id-missing ${GLOSSARY}`,
            ]);
        });
    });
});
