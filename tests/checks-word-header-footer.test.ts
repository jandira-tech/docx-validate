/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `word-header-footer-relationship` (#606, #613, #619): a `w:headerReference` or
 * `w:footerReference` in the main document's body sectPr, or in a reached
 * glossary's docPartBody sectPr, whose `r:id` names a relationship whose Type
 * is not exactly the Transitional or the ISO Strict `header` (respectively
 * `footer`) relationship URI makes Word 16.115 offer recovery (sessions 26
 * and 27 — chunks s26a/s26b/s26c and s27a/s27b/s27c — in
 * docs/plans/evidence/word-probes-2026-10-24.md), and so does a reference
 * whose `r:id` names no Relationship at all while the part's relationships
 * part exists and parses (session 32, #619). Every session 26 specimen
 * below copies the exact shape of the session 26 builders; every session 27
 * specimen takes the clean session 26 shape and replaces only the
 * relationship Type.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { withTempDir } from "../src/lib/run-cli";
import { WORD_HEADER_FOOTER_RELATIONSHIP, wordHeaderFooterRelationshipCheck } from "../src/scripts/office/checks/word-header-footer";
import { registerStoryReferenceIdentity } from "../src/scripts/office/story-reference-identity";
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

const STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/**
 * Session 27's #613 shape: the clean session 26 shape with only the
 * relationship Type replaced. `typeSource` is written verbatim into the
 * sidecar — character references like `&#xA;` survive attribute-value
 * normalization — and `parsedType` is the attribute value the check reads and
 * the finding message names.
 */
interface RelTypeSpecimen {
    readonly name: string;
    readonly reference: "headerReference" | "footerReference";
    readonly typeSource: string;
    readonly parsedType: string;
}

const relTypeSpecimenParts = (specimen: RelTypeSpecimen): Record<string, string> => {
    const segment = specimen.reference === "headerReference" ? "header" : "footer";
    const parts = storyReferenceParts(
        specimen.reference,
        segment,
        `${segment}9.xml`,
        segment === "header" ? STORY_HEADER_XML : STORY_FOOTER_XML,
        storyContentType(segment),
    );
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
        `Type="${REL}/${segment}"`,
        `Type="${specimen.typeSource}"`,
    );
    return parts;
};

/**
 * Session 27's #613 refused Types (repair-warning in Word 16.115): the four
 * the old ends-in test wrongly passed, then the five it already refused.
 */
const REFUSED_TYPES: readonly RelTypeSpecimen[] = [
    {
        name: "headerref-header-rel-type-example-com",
        reference: "headerReference",
        typeSource: "http://example.com/header",
        parsedType: "http://example.com/header",
    },
    {
        name: "headerref-header-rel-type-custom-header",
        reference: "headerReference",
        typeSource: `${REL}/custom/header`,
        parsedType: `${REL}/custom/header`,
    },
    {
        name: "footerref-footer-rel-type-example-com",
        reference: "footerReference",
        typeSource: "http://example.com/footer",
        parsedType: "http://example.com/footer",
    },
    {
        name: "footerref-footer-rel-type-custom-footer",
        reference: "footerReference",
        typeSource: `${REL}/custom/footer`,
        parsedType: `${REL}/custom/footer`,
    },
    {
        name: "headerref-header-rel-type-upper",
        reference: "headerReference",
        typeSource: `${REL}/HEADER`,
        parsedType: `${REL}/HEADER`,
    },
    {
        name: "footerref-footer-rel-type-upper",
        reference: "footerReference",
        typeSource: `${REL}/FOOTER`,
        parsedType: `${REL}/FOOTER`,
    },
    {
        name: "headerref-header-rel-type-padded",
        reference: "headerReference",
        typeSource: ` ${REL}/header `,
        parsedType: ` ${REL}/header `,
    },
    {
        name: "headerref-header-rel-type-tab-newline",
        reference: "headerReference",
        typeSource: `&#x9;${REL}/header&#xA;`,
        parsedType: `\t${REL}/header\n`,
    },
    {
        name: "headerref-header-rel-type-bare",
        reference: "headerReference",
        typeSource: "header",
        parsedType: "header",
    },
];

/** Session 27's #613 controls (clean-open in Word 16.115): the only accepted Types, byte for byte. */
const CLEAN_TYPES: readonly RelTypeSpecimen[] = [
    {
        name: "headerref-header-rel-type-strict-uri",
        reference: "headerReference",
        typeSource: `${STRICT_REL}/header`,
        parsedType: `${STRICT_REL}/header`,
    },
    {
        name: "footerref-footer-rel-type-strict-uri",
        reference: "footerReference",
        typeSource: `${STRICT_REL}/footer`,
        parsedType: `${STRICT_REL}/footer`,
    },
    {
        name: "headerref-header-rel-type-transitional-uri",
        reference: "headerReference",
        typeSource: `${REL}/header`,
        parsedType: `${REL}/header`,
    },
    {
        name: "footerref-footer-rel-type-transitional-uri",
        reference: "footerReference",
        typeSource: `${REL}/footer`,
        parsedType: `${REL}/footer`,
    },
];

describe(
    "word-header-footer-relationship accepts exactly two Types per reference (#613, Word 16.115 session 27)",
    { timeout: 240_000 },
    () => {
        describe.each(REFUSED_TYPES)("refused Type specimen %s", (specimen) => {
            it("word-valid fails with exactly one Word-blocking error naming element, r:id and Type", async () => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `${specimen.name}.docx`);
                    writeFileSync(file, await buildDocx(relTypeSpecimenParts(specimen)));
                    const result = await validate(file, { profile: "word-valid" });
                    const found = result.issues.filter((issue) => issue.code === CODE);
                    expect(result.valid).toBe(false);
                    expect(found).toHaveLength(1);
                    expect(found[0]?.severity).toBe("error");
                    expect(found[0]?.path).toBe("word/document.xml");
                    expect(found[0]?.message).toContain(`<w:${specimen.reference}>`);
                    expect(found[0]?.message).toContain("rIdX");
                    expect(found[0]?.message).toContain(specimen.parsedType);
                });
            });

            it("strict reports the new code too (strict ⊇ word-valid)", async () => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `${specimen.name}.docx`);
                    writeFileSync(file, await buildDocx(relTypeSpecimenParts(specimen)));
                    const strict = await validate(file, { profile: "strict" });
                    expect(strict.valid).toBe(false);
                    expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
                });
            });
        });

        describe.each(CLEAN_TYPES)("clean Type specimen %s", (specimen) => {
            it("word-valid passes with no new finding", async () => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `${specimen.name}.docx`);
                    writeFileSync(file, await buildDocx(relTypeSpecimenParts(specimen)));
                    const result = await validate(file, { profile: "word-valid" });
                    expect(result.valid).toBe(true);
                    expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                });
            });
        });
    },
);

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
                    `a w:headerReference needs a Relationship of Type ${R}/header ` +
                    `(or the ISO Strict ${STRICT_REL}/header). ` +
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

        /**
         * `validateAllRelationshipIds`' mismatch issue for an element, with
         * the structured identity base.ts registers (#616). The prose is
         * opaque to the helper on purpose: quotes or crafted prefixes in an
         * `r:id` must not change the pairing.
         */
        const relsMismatch = (
            part: string,
            localName: "headerReference" | "footerReference",
            rid: string,
            namespaceURI = W,
        ): ValidationIssue => {
            const issue: ValidationIssue = {
                severity: "error",
                code: "rels-id-mismatch",
                path: part,
                message: "mismatch prose the helper never parses (#616)",
            };
            registerStoryReferenceIdentity(issue, { part, namespaceURI, localName, rid });
            return issue;
        };

        /** The Word rule's finding for the same element, with the identity the check registers (#616). */
        const wordFinding = (
            part: string,
            localName: "headerReference" | "footerReference",
            rid: string,
            namespaceURI = W,
        ): ValidationIssue => {
            const issue: ValidationIssue = {
                severity: "error",
                code: CODE,
                path: part,
                message: "word-rule prose the helper never parses (#616)",
            };
            registerStoryReferenceIdentity(issue, { part, namespaceURI, localName, rid });
            return issue;
        };

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

        it("keeps a mismatch in another part, for another r:id, for another element, or in another namespace", () => {
            const matching = relsMismatch("word/document.xml", "headerReference", "rId1");
            const otherPart = relsMismatch("word/glossary/document.xml", "headerReference", "rId1");
            const otherRid = relsMismatch("word/document.xml", "headerReference", "rId2");
            const otherElement = relsMismatch("word/document.xml", "footerReference", "rId1");
            const otherNamespace = relsMismatch("word/document.xml", "headerReference", "rId1", "urn:x");
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [matching, otherPart, otherRid, otherElement, otherNamespace] },
                { valid: false, issues: [wordFinding("word/document.xml", "headerReference", "rId1")] },
            );
            expect(result.issues).toEqual([otherPart, otherRid, otherElement, otherNamespace]);
        });

        /**
         * `validateAllRelationshipIds`' missing-id issue for an element, with
         * the structured identity base.ts registers (#619).
         */
        const relsMissing = (
            part: string,
            localName: "headerReference" | "footerReference",
            rid: string,
            namespaceURI = W,
        ): ValidationIssue => {
            const issue: ValidationIssue = {
                severity: "error",
                code: "rels-id-missing",
                path: part,
                message: "missing prose the helper never parses (#619)",
            };
            registerStoryReferenceIdentity(issue, { part, namespaceURI, localName, rid });
            return issue;
        };

        it("drops a rels-id-missing paired by identity with a Word finding (#619)", () => {
            const missing = relsMissing("word/document.xml", "footerReference", "rId9");
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [missing] },
                { valid: false, issues: [wordFinding("word/document.xml", "footerReference", "rId9")] },
            );
            expect(result.issues).toEqual([]);
        });

        it("keeps a rels-id-missing with no registered identity (#619)", () => {
            const stranger: ValidationIssue = {
                severity: "error",
                code: "rels-id-missing",
                path: "word/document.xml",
                message: "a producer that registered nothing (#619)",
            };
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [stranger] },
                { valid: false, issues: [wordFinding("word/document.xml", "footerReference", "rId9")] },
            );
            expect(result.issues).toEqual([stranger]);
        });

        it("keeps a finding with no registered identity", () => {
            const stranger: ValidationIssue = {
                severity: "error",
                code: "rels-id-mismatch",
                path: "word/document.xml",
                message: "a producer that registered nothing (#616)",
            };
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [stranger] },
                { valid: false, issues: [wordFinding("word/document.xml", "headerReference", "rId1")] },
            );
            expect(result.issues).toEqual([stranger]);
        });

        it("never drops through a Word finding with no registered identity", () => {
            const silent: ValidationIssue = {
                severity: "error",
                code: CODE,
                path: "word/document.xml",
                message: "a Word-rule issue that registered nothing (#616)",
            };
            const mismatch = relsMismatch("word/document.xml", "headerReference", "rId1");
            const result = withoutRepeatedHeaderFooterFindings({ valid: false, issues: [mismatch] }, { valid: false, issues: [silent] });
            expect(result.issues).toEqual([mismatch]);
        });

        it("normalizes backslash part paths before matching", () => {
            const result = withoutRepeatedHeaderFooterFindings(
                { valid: false, issues: [relsMismatch("word\\document.xml", "headerReference", "rId1")] },
                { valid: false, issues: [wordFinding("word/document.xml", "headerReference", "rId1")] },
            );
            expect(result.issues).toEqual([]);
        });
    });
});

describe("#616: the hand-off carries structured identity, never message prose", { timeout: 240_000 }, () => {
    /** XML-escape a value the way a producer writes it into an attribute. */
    const esc = (s: string): string => s.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");

    /**
     * F1 (review BA): the refused specimen with `rIdX` renamed, in the
     * reference and its Relationship, to an `r:id` carrying the delimiter
     * each former regex could not consume.
     */
    it.each([
        { name: "apostrophe", rid: "rId'X" },
        { name: "double-quote", rid: `rId"X` },
    ] as const)("a full run carries only the Word finding for the %s r:id", async (specimen) => {
        await withTempDir(async (dir) => {
            const parts = storyReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer"));
            for (const part of Object.keys(parts)) parts[part] = parts[part]!.replaceAll("rIdX", esc(specimen.rid));
            const file = path.join(dir, `quoted-${specimen.name}.docx`);
            writeFileSync(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.filter((issue) => issue.code === CODE && issue.path === "word/document.xml"),
                    profile,
                ).toHaveLength(1);
                expect(
                    result.issues.filter((issue) => issue.code === "rels-id-mismatch" && issue.message.includes("headerReference")),
                    profile,
                ).toEqual([]);
            }
        });
    });

    /**
     * F2 (review BA): the spoof pair — a foreign `x:headerReference` whose
     * `rIdX` is the footer relationship, plus a `w:headerReference` whose
     * crafted `r:id` once made the Word message look like it named `rIdX`.
     */
    it("a full run keeps the foreign element's mismatch and the Word element's finding", async () => {
        await withTempDir(async (dir) => {
            const parts = storyReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer"));
            const spoof = `rIdX" names Relationship 'spoof`;
            parts["word/document.xml"] = parts["word/document.xml"]!.replace("<w:document ", '<w:document xmlns:x="urn:x" ').replace(
                '<w:headerReference w:type="default" r:id="rIdX"/>',
                `<x:headerReference w:type="default" r:id="rIdX"/><w:headerReference w:type="default" r:id="${esc(spoof)}"/>`,
            );
            parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="${esc(spoof)}" Type="http://example.com/header" Target="footer9.xml"/></Relationships>`,
            );
            const file = path.join(dir, "spoof-pair.docx");
            writeFileSync(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                const mismatches = result.issues.filter((issue) => issue.code === "rels-id-mismatch");
                expect(mismatches, profile).toHaveLength(1);
                expect(mismatches[0]?.message, profile).toBe(
                    "<headerReference> references 'rIdX' which points to 'footer' but should point to a 'header' relationship",
                );
                expect(
                    result.issues.filter((issue) => issue.code === CODE && issue.path === "word/document.xml"),
                    profile,
                ).toHaveLength(1);
            }
        });
    });
});

/**
 * Session 32's #619 shape: the MINIMAL_PARTS main document with `xmlns:r`
 * declared, whose body sectPr's `reference` names `rid` — a Relationship no
 * Id in the kept sidecar carries. Word 16.115 offered recovery for the
 * footer shape and for the same package without the sidecar; see
 * docs/plans/evidence/word-probes-2026-10-24.md, "Session 32".
 */
const danglingReferenceParts = (reference: "headerReference" | "footerReference", rid = "rId9"): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
        `<w:document xmlns:w="${W}"`,
        `<w:document xmlns:w="${W}" xmlns:r="${REL}"`,
    ).replace("<w:sectPr>", `<w:sectPr><w:${reference} w:type="default" r:id="${rid}"/>`),
});

describe("#619: a story reference whose r:id names no Relationship (Word 16.115 session 32)", { timeout: 240_000 }, () => {
    it("word-valid and strict report exactly one Word rule finding and no rels-id-missing (the footer shape)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "footerref-no-relationship.docx");
            writeFileSync(file, await buildDocx(danglingReferenceParts("footerReference")));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid, profile).toBe(false);
                expect(found, profile).toHaveLength(1);
                expect(found[0]?.severity, profile).toBe("error");
                expect(found[0]?.path, profile).toBe("word/document.xml");
                expect(found[0]?.message, profile).toContain("<w:footerReference>");
                expect(found[0]?.message, profile).toContain('r:id="rId9"');
                expect(found[0]?.message, profile).toContain("no Relationship with Id 'rId9'");
                expect(
                    result.issues.filter((issue) => issue.code === "rels-id-missing"),
                    profile,
                ).toEqual([]);
            }
        });
    });

    it("the w:headerReference mirror behaves alike", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "headerref-no-relationship.docx");
            writeFileSync(file, await buildDocx(danglingReferenceParts("headerReference")));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid, profile).toBe(false);
                expect(found, profile).toHaveLength(1);
                expect(found[0]?.path, profile).toBe("word/document.xml");
                expect(found[0]?.message, profile).toContain("<w:headerReference>");
                expect(found[0]?.message, profile).toContain("no Relationship with Id 'rId9'");
                expect(
                    result.issues.filter((issue) => issue.code === "rels-id-missing"),
                    profile,
                ).toEqual([]);
            }
        });
    });

    it("judges a reached glossary's reference against the glossary's own .rels", async () => {
        await withTempDir(async (dir) => {
            const parts = glossaryStoryReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer"));
            // The sidecar keeps its footer relationship, under an Id the
            // reference does not name.
            parts["word/glossary/_rels/document.xml.rels"] = parts["word/glossary/_rels/document.xml.rels"]!.replace(
                'Id="rIdF"',
                'Id="rIdOther"',
            );
            const file = path.join(dir, "glossary-no-relationship.docx");
            writeFileSync(file, await buildDocx(parts));
            const wordValid = await validate(file, { profile: "word-valid" });
            const found = wordValid.issues.filter((issue) => issue.code === CODE);
            expect(found).toHaveLength(1);
            expect(found[0]?.path).toBe("word/glossary/document.xml");
            expect(found[0]?.message).toContain("<w:headerReference>");
            expect(found[0]?.message).toContain("no Relationship with Id 'rIdF'");
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === CODE && issue.path === "word/glossary/document.xml")).toHaveLength(1);
            expect(
                strict.issues.filter((issue) => issue.code === "rels-id-missing" && issue.path === "word/glossary/document.xml"),
            ).toEqual([]);
        });
    });

    it("a part with no relationships part stays rels-missing-sidecar's alone", async () => {
        await withTempDir(async (dir) => {
            const parts = danglingReferenceParts("footerReference");
            delete parts["word/_rels/document.xml.rels"];
            const file = path.join(dir, "no-sidecar.docx");
            writeFileSync(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.filter((issue) => issue.code === CODE),
                    profile,
                ).toEqual([]);
                expect(
                    result.issues.some((issue) => issue.code === "rels-missing-sidecar"),
                    profile,
                ).toBe(true);
            }
        });
    });

    it("a foreign-namespace reference keeps its rels-id-missing under strict", async () => {
        await withTempDir(async (dir) => {
            // The Word rule judges exact Word expanded names, so it never
            // reports urn:x's footerReference; the missing finding must
            // survive the full run's hand-off.
            const parts = danglingReferenceParts("footerReference");
            parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                `<w:document xmlns:w="${W}"`,
                `<w:document xmlns:w="${W}" xmlns:x="urn:x"`,
            ).replace("<w:footerReference", "<x:footerReference");
            const file = path.join(dir, "foreign-namespace-dangling.docx");
            writeFileSync(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.filter((issue) => issue.code === CODE),
                    profile,
                ).toEqual([]);
            }
            const strict = await validate(file, { profile: "strict" });
            expect(
                strict.issues.filter((issue) => issue.code === "rels-id-missing" && issue.message.includes("footerReference")),
            ).toHaveLength(1);
        });
    });

    it("a standalone validateAllRelationshipIds call still reports the rels-id-missing", async () => {
        await withTempDir(async (dir) => {
            const unpacked = writeUnpacked(path.join(dir, "dangling"), danglingReferenceParts("footerReference"));
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const result = await new DOCXSchemaValidator({ unpackedDir: unpacked, profile }).validateAllRelationshipIds();
                const found = result.issues.filter((issue) => issue.code === "rels-id-missing");
                expect(found, profile).toHaveLength(1);
                expect(found[0]?.path, profile).toBe("word/document.xml");
                expect(found[0]?.message, profile).toBe(
                    "<footerReference> r:id references non-existent relationship 'rId9' (valid IDs: rId1)",
                );
                expect(result.valid, profile).toBe(false);
            }
        });
    });

    it("an unreached glossary's dangling reference keeps its rels-id-missing under strict", async () => {
        await withTempDir(async (dir) => {
            // The glossary shape whose reference names nothing in the
            // glossary's .rels, but the main document relates no glossary:
            // Word does not read it, so only the legacy check judges it.
            const parts: Record<string, string> = {
                ...glossaryStoryReferenceParts("headerReference", "footer", "footer9.xml", STORY_FOOTER_XML, T("footer")),
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!,
            };
            parts["word/glossary/_rels/document.xml.rels"] = parts["word/glossary/_rels/document.xml.rels"]!.replace(
                'Id="rIdF"',
                'Id="rIdOther"',
            );
            const file = path.join(dir, "unreached-glossary-no-relationship.docx");
            writeFileSync(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.filter((issue) => issue.code === CODE),
                    profile,
                ).toEqual([]);
            }
            const strict = await validate(file, { profile: "strict" });
            expect(
                strict.issues.filter((issue) => issue.code === "rels-id-missing" && issue.path === "word/glossary/document.xml"),
            ).toHaveLength(1);
        });
    });

    it("quick and named runs agree with the full word-valid run (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "footerref-no-relationship.docx");
            writeFileSync(file, await buildDocx(danglingReferenceParts("footerReference")));
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

    it("lenient keeps the rels-id-missing and runs no Word rule", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "footerref-no-relationship.docx");
            writeFileSync(file, await buildDocx(danglingReferenceParts("footerReference")));
            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.issues.some((issue) => issue.code === CODE)).toBe(false);
            const found = lenient.issues.filter((issue) => issue.code === "rels-id-missing");
            expect(found).toHaveLength(1);
            expect(found[0]?.message).toBe("<footerReference> r:id references non-existent relationship 'rId9' (valid IDs: rId1)");
        });
    });
});
