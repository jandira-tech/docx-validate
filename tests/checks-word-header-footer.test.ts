/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-header-footer-relationship` (#606): a `w:headerReference` or
 * `w:footerReference` in the main document's body sectPr, or in a reached
 * glossary's docPartBody sectPr, whose `r:id` names a relationship whose Type
 * does not end in `/header` / `/footer` makes Word 16.115 offer recovery
 * (session 26 — chunks s26a, s26b, s26c — in
 * docs/plans/evidence/word-probes-2026-10-24.md). Every specimen below copies
 * the exact shape of the session 26 builders.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { withTempDir } from "../src/lib/run-cli";
import { WORD_HEADER_FOOTER_RELATIONSHIP, wordHeaderFooterRelationshipCheck } from "../src/scripts/office/checks/word-header-footer";
import { validate } from "../src/scripts/office/validate";
import * as docxValidators from "../src/scripts/office/validators/docx";
import {
    glossaryStoryReferenceParts,
    STORY_FOOTER_XML,
    STORY_HEADER_XML,
    storyContentType,
    storyReferenceParts,
    storyReferenceToOwnStylesParts,
} from "./support/story-content-type-specimens";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { document, R, rels, relation } from "./support/word-rule-test-xml";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const T = storyContentType;
const CODE = WORD_HEADER_FOOTER_RELATIONSHIP;

/** One session 26 specimen: its parts, where the reference lives, and what the message must name. */
interface Specimen {
    readonly name: string;
    readonly parts: () => Record<string, string>;
    readonly path: string;
    readonly element: "headerReference" | "footerReference";
    readonly rid: string;
    readonly relType: string;
}

/** Every refused specimen of session 26 (repair-warning in Word 16.115). */
const REFUSED: readonly Specimen[] = [
    {
        name: "headerref-footer-rel-typed-footer-exact",
        parts: () => storyReferenceParts("headerReference", "footer", "header9.xml", STORY_HEADER_XML, T("footer")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "headerref-footer-rel-typed-footer-charset",
        parts: () => storyReferenceParts("headerReference", "footer", "header9.xml", STORY_HEADER_XML, `${T("footer")}; charset=utf-8`),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "headerref-footer-rel-typed-header-exact",
        parts: () => storyReferenceParts("headerReference", "footer", "header9.xml", STORY_HEADER_XML, T("header")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "headerref-footer-rel-ftr-typed-footer",
        parts: () => storyReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "headerref-footer-rel-hdr-typed-styles",
        parts: () => storyReferenceParts("headerReference", "footer", "header9.xml", STORY_HEADER_XML, T("styles")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "footerref-header-rel-hdr-typed-header",
        parts: () => storyReferenceParts("footerReference", "header", "header9.xml", STORY_HEADER_XML, T("header")),
        path: "word/document.xml",
        element: "footerReference",
        rid: "rIdX",
        relType: `${REL}/header`,
    },
    {
        name: "footerref-header-rel-ftr-typed-footer",
        parts: () => storyReferenceParts("footerReference", "header", "footer9.xml", STORY_FOOTER_XML, T("footer")),
        path: "word/document.xml",
        element: "footerReference",
        rid: "rIdX",
        relType: `${REL}/header`,
    },
    {
        name: "headerref-styles-rel",
        parts: () => storyReferenceToOwnStylesParts(),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rId1",
        relType: `${REL}/styles`,
    },
    {
        name: "glossary-headerref-footer-rel-ftr-typed-footer",
        parts: () => glossaryStoryReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
        path: "word/glossary/document.xml",
        element: "headerReference",
        rid: "rIdF",
        relType: `${REL}/footer`,
    },
    {
        name: "glossary-footerref-header-rel-hdr-typed-header",
        parts: () => glossaryStoryReferenceParts("footerReference", "header", "header9.xml", STORY_HEADER_XML, T("header")),
        path: "word/glossary/document.xml",
        element: "footerReference",
        rid: "rIdF",
        relType: `${REL}/header`,
    },
];

/** Every clean specimen of session 26 (clean-open in Word 16.115). */
const CLEAN: readonly Specimen[] = [
    {
        name: "headerref-header-rel-hdr-typed-header",
        parts: () => storyReferenceParts("headerReference", "header", "header9.xml", STORY_HEADER_XML, T("header")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/header`,
    },
    {
        name: "footerref-footer-rel-hdr-typed-footer",
        parts: () => storyReferenceParts("footerReference", "footer", "footer9.xml", STORY_HEADER_XML, T("footer")),
        path: "word/document.xml",
        element: "footerReference",
        rid: "rIdX",
        relType: `${REL}/footer`,
    },
    {
        name: "headerref-header-rel-ftr-typed-header",
        parts: () => storyReferenceParts("headerReference", "header", "header9.xml", STORY_FOOTER_XML, T("header")),
        path: "word/document.xml",
        element: "headerReference",
        rid: "rIdX",
        relType: `${REL}/header`,
    },
    {
        name: "glossary-footerref-footer-rel-ftr-typed-footer",
        parts: () => glossaryStoryReferenceParts("footerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
        path: "word/glossary/document.xml",
        element: "footerReference",
        rid: "rIdF",
        relType: `${REL}/footer`,
    },
];

const specimenFile = async (dir: string, specimen: Specimen): Promise<string> => {
    const file = path.join(dir, `${specimen.name}.docx`);
    writeFileSync(file, await buildDocx(specimen.parts()));
    return file;
};

/** Write `parts` as an unpacked package directory and return its path. */
const writeUnpacked = (dir: string, parts: Record<string, string>): string => {
    mkdirSync(dir, { recursive: true });
    for (const [part, xml] of Object.entries(parts)) {
        const file = path.join(dir, part);
        mkdirSync(path.dirname(file), { recursive: true });
        writeFileSync(file, xml);
    }
    return dir;
};

// Reached through the module namespace so a not-yet-added export fails the
// #612 tests on an assertion, never on the import.
const { DOCXSchemaValidator, withoutRepeatedHeaderFooterFindings } = docxValidators;

describe("word-header-footer-relationship (#606, Word 16.115 session 26)", { timeout: 240_000 }, () => {
    describe.each(REFUSED)("refused specimen %s", (specimen) => {
        it("word-valid fails with the Word-blocking error naming element, r:id and relationship type", async () => {
            await withTempDir(async (dir) => {
                const result = await validate(await specimenFile(dir, specimen), { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe(specimen.path);
                expect(found[0]?.message).toContain(`<w:${specimen.element}>`);
                expect(found[0]?.message).toContain(specimen.rid);
                expect(found[0]?.message).toContain(specimen.relType);
            });
        });
    });

    describe.each(CLEAN)("clean specimen %s", (specimen) => {
        it("word-valid passes with no new finding", async () => {
            await withTempDir(async (dir) => {
                const result = await validate(await specimenFile(dir, specimen), { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });

    it("the minimal control stays valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx(MINIMAL_PARTS));
            expect((await validate(file, { profile: "word-valid" })).valid).toBe(true);
        });
    });

    describe("one finding per element across profiles", () => {
        it("strict reports the new code and no rels-id-mismatch for the same element", async () => {
            await withTempDir(async (dir) => {
                const file = await specimenFile(dir, REFUSED[0]!);
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
                expect(strict.issues.some((issue) => issue.code === "rels-id-mismatch" && issue.message.includes("headerReference"))).toBe(
                    false,
                );
                expect(strict.valid).toBe(false);
            });
        });

        it("lenient keeps rels-id-missing's territory and its own rels-id-mismatch", async () => {
            await withTempDir(async (dir) => {
                const file = await specimenFile(dir, REFUSED[0]!);
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.some((issue) => issue.code === CODE)).toBe(false);
                expect(lenient.issues.some((issue) => issue.code === "rels-id-mismatch" && issue.message.includes("headerReference"))).toBe(
                    true,
                );
            });
        });

        it("strict still reports rels-id-mismatch where the Word check does not reach", async () => {
            await withTempDir(async (dir) => {
                // The refused glossary shape, but the main document relates no
                // glossary: Word does not read it, so only the legacy check
                // judges its headerReference there.
                const unreached = {
                    ...glossaryStoryReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
                    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!,
                };
                const file = path.join(dir, "unreached-glossary.docx");
                writeFileSync(file, await buildDocx(unreached));
                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues.some((issue) => issue.code === CODE)).toBe(false);
                const strict = await validate(file, { profile: "strict" });
                expect(
                    strict.issues.some((issue) => issue.code === "rels-id-mismatch" && issue.path === "word/glossary/document.xml"),
                ).toBe(true);
                expect(strict.issues.some((issue) => issue.code === CODE)).toBe(false);
            });
        });
    });

    it("leaves an unresolvable r:id to rels-id-missing", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "missing-rid.docx");
            writeFileSync(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                        `<w:document xmlns:w="${W}"`,
                        `<w:document xmlns:w="${W}" xmlns:r="${REL}"`,
                    ).replace("<w:sectPr>", `<w:sectPr><w:headerReference w:type="default" r:id="rIdNone"/>`),
                }),
            );
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.some((issue) => issue.code === CODE)).toBe(false);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.some((issue) => issue.code === "rels-id-missing" && issue.message.includes("headerReference"))).toBe(true);
            expect(strict.issues.some((issue) => issue.code === CODE)).toBe(false);
        });
    });

    it("quick and named runs agree with the full word-valid run (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = await specimenFile(dir, REFUSED[0]!);
            const full = await validate(file, { profile: "word-valid" });
            const expected = full.issues.filter((issue) => issue.code === CODE);
            expect(expected).toHaveLength(1);
            for (const selection of [{ quick: true }, { checks: ["relationships"] }, { checks: [CODE] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(partial.valid, tag).toBe(false);
                expect(
                    partial.issues.filter((issue) => issue.code === CODE),
                    tag,
                ).toEqual(expected);
            }
        });
    });
});

/** A package whose main document is `word/document.xml`, with the given body and sidecar relationships. */
const mainPackage = (body: string, docRels: string): Record<string, string> => ({
    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
    "word/document.xml": document(body),
    "word/_rels/document.xml.rels": rels(docRels),
});

describe("word-header-footer-relationship check runner", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries(
            mainPackage(
                '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
                relation("rIdH", "footer", "", "footer9.xml"),
            ),
        );
        expect(await wordHeaderFooterRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: CODE,
                path: "word/document.xml",
                message:
                    `<w:headerReference> r:id="rIdH" names Relationship 'rIdH' of Type '${R}/footer'; ` +
                    "a w:headerReference needs a Relationship whose Type ends in /header. " +
                    "Microsoft Word offers to recover the file. Point the r:id at such a Relationship, or retype that Relationship.",
            },
        ]);
    });

    it("accepts matching references, and refuses the mirror", async () => {
        const clean = ValidationContext.fromEntries(
            mainPackage(
                '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>',
                relation("rIdH", "header", "", "header9.xml") + relation("rIdF", "footer", "", "footer9.xml"),
            ),
        );
        expect(await wordHeaderFooterRelationshipCheck.run(clean)).toEqual([]);
        const mirror = ValidationContext.fromEntries(
            mainPackage(
                '<w:sectPr><w:footerReference w:type="default" r:id="rIdH"/></w:sectPr>',
                relation("rIdH", "header", "", "header9.xml"),
            ),
        );
        expect(await wordHeaderFooterRelationshipCheck.run(mirror)).toEqual([
            expect.objectContaining({ code: CODE, path: "word/document.xml" }),
        ]);
    });

    it("reads Strict relationship ids and names an empty Type", async () => {
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document('<w:sectPr><w:footerReference w:type="default" rs:id="rIdH"/></w:sectPr>'),
            "word/_rels/document.xml.rels": rels(`<Relationship Id="rIdH"/>`),
        });
        const issues = await wordHeaderFooterRelationshipCheck.run(ctx);
        expect(issues).toHaveLength(1);
        expect(issues[0]?.message).toContain("Type '(none)'");
        const clean = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document('<w:sectPr><w:footerReference w:type="default" rs:id="rIdH"/></w:sectPr>'),
            "word/_rels/document.xml.rels": rels(relation("rIdH", "footer", "", "footer9.xml")),
        });
        expect(await wordHeaderFooterRelationshipCheck.run(clean)).toEqual([]);
    });

    it("judges a reached glossary's references at the glossary's path", async () => {
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document(""),
            "word/_rels/document.xml.rels": rels(`<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/>`),
            "word/glossary/document.xml": document('<w:sectPr><w:headerReference w:type="default" r:id="rIdF"/></w:sectPr>'),
            "word/glossary/_rels/document.xml.rels": rels(relation("rIdF", "footer", "", "footer9.xml")),
        });
        expect(await wordHeaderFooterRelationshipCheck.run(ctx)).toEqual([
            expect.objectContaining({ code: CODE, path: "word/glossary/document.xml" }),
        ]);
    });

    it("reports a story root once when two relationship types reach it", async () => {
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>'),
            "word/_rels/document.xml.rels": rels(
                relation("rIdH", "footer", "", "footer9.xml") +
                    `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="document.xml"/>`,
            ),
        });
        expect(await wordHeaderFooterRelationshipCheck.run(ctx)).toEqual([expect.objectContaining({ code: CODE })]);
    });

    it("skips absent, irrelevant and unparseable inputs", async () => {
        expect(await wordHeaderFooterRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(await wordHeaderFooterRelationshipCheck.run(ValidationContext.fromEntries({ "_rels/.rels": "<Relationships" }))).toEqual([]);
        expect(
            await wordHeaderFooterRelationshipCheck.run(
                ValidationContext.fromEntries({
                    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
                    "word/document.xml": "<broken",
                    "word/_rels/document.xml.rels": rels(relation("rIdH", "footer", "", "footer9.xml")),
                }),
            ),
        ).toEqual([]);
        expect(
            await wordHeaderFooterRelationshipCheck.run(
                ValidationContext.fromEntries({
                    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
                    "word/document.xml": document('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>'),
                    "word/_rels/document.xml.rels": "<Relationships",
                }),
            ),
        ).toEqual([]);
        expect(
            await wordHeaderFooterRelationshipCheck.run(
                ValidationContext.fromEntries({
                    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
                    "word/document.xml": document('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>'),
                }),
            ),
        ).toEqual([]);
        expect(
            await wordHeaderFooterRelationshipCheck.run(
                ValidationContext.fromEntries(
                    mainPackage('<w:sectPr><w:headerReference w:type="default"/></w:sectPr>', relation("rIdH", "footer")),
                ),
            ),
        ).toEqual([]);
    });
});

describe("#612: rels-id-mismatch reports every mismatch; full runs deduplicate", { timeout: 240_000 }, () => {
    describe("a standalone validateAllRelationshipIds call", () => {
        it.each([
            {
                name: "headerref-styles-rel",
                parts: () => storyReferenceToOwnStylesParts(),
                path: "word/document.xml",
                message: "<headerReference> references 'rId1' which points to 'styles' but should point to a 'header' relationship",
            },
            {
                name: "glossary-headerref-footer-rel-ftr-typed-footer",
                parts: () => glossaryStoryReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
                path: "word/glossary/document.xml",
                message: "<headerReference> references 'rIdF' which points to 'footer' but should point to a 'header' relationship",
            },
        ] as const)("reports rels-id-mismatch for %s under every profile", async (specimen) => {
            await withTempDir(async (dir) => {
                const unpacked = writeUnpacked(path.join(dir, specimen.name), specimen.parts());
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const result = await new DOCXSchemaValidator({ unpackedDir: unpacked, profile }).validateAllRelationshipIds();
                    const found = result.issues.filter((issue) => issue.code === "rels-id-mismatch");
                    expect(found, profile).toHaveLength(1);
                    expect(found[0]?.path, profile).toBe(specimen.path);
                    expect(found[0]?.message, profile).toBe(specimen.message);
                    expect(result.valid, profile).toBe(false);
                }
            });
        });
    });

    describe("a full run on a foreign-namespace reference", () => {
        it.each(["word-valid", "strict"] as const)(
            "keeps exactly one rels-id-mismatch for an x:headerReference under %s",
            async (profile) => {
                await withTempDir(async (dir) => {
                    // The Word rule judges exact Word expanded names, so it
                    // never reports urn:x's headerReference; the mismatch
                    // must survive the full run's hand-off.
                    const parts = storyReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer"));
                    parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                        `<w:document xmlns:w="${W}"`,
                        `<w:document xmlns:w="${W}" xmlns:x="urn:x"`,
                    ).replaceAll("<w:headerReference", "<x:headerReference");
                    const file = path.join(dir, "foreign-namespace-headerref.docx");
                    writeFileSync(file, await buildDocx(parts));
                    const result = await validate(file, { profile });
                    expect(
                        result.issues.filter((issue) => issue.code === "rels-id-mismatch" && issue.message.includes("headerReference")),
                    ).toHaveLength(1);
                    expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                });
            },
        );
    });

    describe("a full run on an element the Word rule reports", () => {
        it("keeps one finding per element under word-valid and strict", async () => {
            await withTempDir(async (dir) => {
                const file = await specimenFile(dir, REFUSED[0]!);
                for (const profile of ["word-valid", "strict"] as const) {
                    const result = await validate(file, { profile });
                    expect(
                        result.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml"),
                        profile,
                    ).toBe(true);
                    expect(
                        result.issues.some((issue) => issue.code === "rels-id-mismatch" && issue.message.includes("headerReference")),
                        profile,
                    ).toBe(false);
                }
            });
        });
    });

    describe("withoutRepeatedHeaderFooterFindings", () => {
        it("is exported by the validator module", () => {
            expect(withoutRepeatedHeaderFooterFindings).toBeTypeOf("function");
        });

        /** `validateAllRelationshipIds`' mismatch message (validators/base.ts). */
        const relsMismatch = (part: string, localName: "headerReference" | "footerReference", rid: string): ValidationIssue => ({
            severity: "error",
            code: "rels-id-mismatch",
            path: part,
            message: `<${localName}> references '${rid}' which points to 'styles' but should point to a '${localName === "headerReference" ? "header" : "footer"}' relationship`,
        });

        /** The Word rule's message for the same element (checks/word-header-footer.ts). */
        const wordFinding = (part: string, localName: "headerReference" | "footerReference", rid: string): ValidationIssue => ({
            severity: "error",
            code: CODE,
            path: part,
            message:
                `<w:${localName}> r:id="${rid}" names Relationship '${rid}' of Type '${R}/styles'; ` +
                `a w:${localName} needs a Relationship whose Type ends in /${localName === "headerReference" ? "header" : "footer"}. ` +
                "Microsoft Word offers to recover the file. Point the r:id at such a Relationship, or retype that Relationship.",
        });

        it("drops matching pairs one-for-one, so one Word finding drops only one mismatch", () => {
            const pair = [
                relsMismatch("word/document.xml", "headerReference", "rId1"),
                relsMismatch("word/document.xml", "headerReference", "rId1"),
            ];
            const one = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: pair },
                { valid: false, issues: [wordFinding("word/document.xml", "headerReference", "rId1")] },
            );
            expect(one.issues).toEqual([pair[1]]);
            const both = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: pair },
                {
                    valid: false,
                    issues: [
                        wordFinding("word/document.xml", "headerReference", "rId1"),
                        wordFinding("word/document.xml", "headerReference", "rId1"),
                    ],
                },
            );
            expect(both.issues).toEqual([]);
        });

        it("keeps a mismatch in another part, for another r:id, or for another element", () => {
            const matching = relsMismatch("word/document.xml", "headerReference", "rId1");
            const otherPart = relsMismatch("word/glossary/document.xml", "headerReference", "rId1");
            const otherRid = relsMismatch("word/document.xml", "headerReference", "rId2");
            const otherElement = relsMismatch("word/document.xml", "footerReference", "rId1");
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [matching, otherPart, otherRid, otherElement] },
                { valid: false, issues: [wordFinding("word/document.xml", "headerReference", "rId1")] },
            );
            expect(result.issues).toEqual([otherPart, otherRid, otherElement]);
        });

        it("normalizes backslash part paths before matching", () => {
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [relsMismatch("word/document.xml", "headerReference", "rId1")] },
                { valid: false, issues: [wordFinding("word\\document.xml", "headerReference", "rId1")] },
            );
            expect(result.issues).toEqual([]);
        });
    });
});
