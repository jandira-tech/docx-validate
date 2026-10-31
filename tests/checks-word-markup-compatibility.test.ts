/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { compatibilityMainParts, compatibilityStoryParts } from "./support/word-compatibility-parts";
import { wordMarkupCompatibilityCheck as check } from "../src/scripts/office/checks/word-markup-compatibility";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const ROOT = `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" xmlns:e="urn:future"`;
const document = (body: string, attributes = "") => `<w:document ${ROOT}${attributes}><w:body>${body}</w:body></w:document>`;
const FALLBACK_ONLY = "<mc:AlternateContent><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>";
const MUST_UNDERSTAND_FOO = ` xmlns:foo="urn:foo" mc:MustUnderstand="foo"`;

describe("markup compatibility as Word processes it", () => {
    it("rejects an AlternateContent without a Choice and a namespace Word must understand but does not", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(FALLBACK_ONLY),
            "word/footnotes.xml": `<w:footnotes ${ROOT} mc:MustUnderstand="e"/>`,
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
            expect.objectContaining({
                code: "mce-must-understand",
                severity: "error",
                path: "word/footnotes.xml",
                message: expect.stringContaining("e"),
            }),
        ]);
    });

    it("accepts what Word understands: a w14 requirement and a Choice with a Fallback", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<mc:AlternateContent><mc:Choice Requires="e"><w:tbl/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>`,
                ` mc:MustUnderstand="w14"`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("ignores content in a branch Word does not select", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback>${FALLBACK_ONLY}</mc:Fallback></mc:AlternateContent>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("reports errors only, judging the XML parts under word/ and no other folder", async () => {
        const ctx = ValidationContext.fromEntries({
            // An Office extension attribute outside mc:Ignorable is a warning
            // of the schema view, not a Word rejection.
            "word/settings.xml": `<w:settings ${ROOT} w14:foo="1"/>`,
            "docProps/custom.xml": document(FALLBACK_ONLY),
            "docProps/foo.%78ml": document(FALLBACK_ONLY),
            "customXml/item1.xml": document(FALLBACK_ONLY),
            "word/styles.xml": document(FALLBACK_ONLY).replace("</w:document>", ""),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("judges the markup compatibility of relationships parts (#636)", async () => {
        const ctx = ValidationContext.fromEntries({
            // Session 34 s34c: mc:MustUnderstand of an unknown namespace on
            // the _rels/.rels root, which Word 16.115 refused.
            "_rels/.rels": `<Relationships xmlns="${PR}" xmlns:mc="${MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"/>`,
            // Session 34 s34d: an undeclared Requires prefix in the main
            // document's sidecar, which Word 16.115 refused.
            "word/_rels/document.xml.rels":
                `<Relationships xmlns="${PR}" xmlns:mc="${MC}">` +
                `<mc:AlternateContent><mc:Choice Requires="nope"><Relationship Id="rId1" Type="urn:t" Target="x.xml"/></mc:Choice>` +
                `<mc:Fallback/></mc:AlternateContent></Relationships>`,
            // Unobserved part, still reported: the profile decides severity.
            "word/_rels/header1.xml.rels": `<Relationships xmlns="${PR}" xmlns:mc="${MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"/>`,
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "_rels/.rels" }),
            expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/_rels/document.xml.rels" }),
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/_rels/header1.xml.rels" }),
        ]);
    });

    it("judges the OPC-equivalent spellings of a word/*.xml part name, as OPC compares them (#639)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/FOO.XML": document("", MUST_UNDERSTAND_FOO),
            "word/foo.%78ml": document("", MUST_UNDERSTAND_FOO),
            "word/foo.xml": document("", MUST_UNDERSTAND_FOO),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/FOO.XML" }),
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/foo.%78ml" }),
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/foo.xml" }),
        ]);
    });

    it("reports relationships-part errors only, keeping warnings to the schema view", async () => {
        const ctx = ValidationContext.fromEntries({
            // An Office extension attribute outside mc:Ignorable is a warning
            // of the schema view, not a Word rejection.
            "word/_rels/header1.xml.rels": `<Relationships xmlns="${PR}" xmlns:chart="http://schemas.microsoft.com/office/2011/relationships/chart" chart:foo="1"/>`,
        });
        expect(await check.run(ctx)).toEqual([]);
    });
});

// Word 16.115 offered recovery for character data directly inside an
// mc:AlternateContent — before the Choice, between the Choice and the
// Fallback (session 65 s65c), after the Fallback, after a Choice with no
// Fallback, and between two Choices (session 67 s67b; every chunk two
// runs, all readings agreeing) — in a used comments part's stored root and
// in a run-level body AlternateContent alike, while the same roots open
// clean with the text inside the Choice or the Fallback (session 65 s65b),
// with whitespace only between the branches, and without the text (session
// 67 s67b). mc.xsd declares none of AlternateContent, Choice and Fallback
// mixed, so the schema view forbids all three positions alike and cannot
// draw Word's line; whitespace is note 8's [ \t\n\r], which pretty-printed
// parts carry. AlternateContent in parts other than the main document and
// comments, and in relationships parts, is unprobed (#851).
describe("character data directly inside AlternateContent (#851, sessions 65 and 67)", () => {
    const COMMENT = `<w:comment w:id="0" w:author="a"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:comment>`;
    const CHOICE = `<mc:Choice Requires="w14"><w:comments xmlns:w="${W}">${COMMENT}</w:comments></mc:Choice>`;
    const UNKNOWN_CHOICE = `<mc:Choice Requires="u"><w:comments xmlns:w="${W}">${COMMENT}</w:comments></mc:Choice>`;
    const FALLBACK = "<mc:Fallback/>";
    /** A `word/comments.xml` whose stored root is an AlternateContent holding `inner`, with the body referencing comment 0. */
    const commentsRoot = (inner: string): Record<string, string> => ({
        "word/document.xml": document(`<w:r><w:commentReference w:id="0"/></w:r>`),
        "word/comments.xml": `<mc:AlternateContent ${ROOT} xmlns:u="urn:example:unknown">${inner}</mc:AlternateContent>`,
    });
    /** The main document's paragraph holding a run-level AlternateContent with `between` between its Choice run and its Fallback run. */
    const bodyRun = (between: string): Record<string, string> => ({
        "word/document.xml": document(
            `<mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>Chosen</w:t></w:r></mc:Choice>${between}` +
                `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback></mc:AlternateContent>`,
        ),
    });
    const malformed = async (parts: Record<string, string>, part: string) =>
        (await check.run(ValidationContext.fromEntries(parts))).filter((issue) => issue.code === "mce-malformed" && issue.path === part);

    // Every position session 65 (s65c) and session 67 (s67b) showed Word
    // refusing, in the comments part's stored root and in a body run alike.
    it.each([
        ["text before the Choice", `Stray${CHOICE}${FALLBACK}`],
        ["text between the Choice and the Fallback", `${CHOICE}Stray${FALLBACK}`],
        ["text after the Fallback", `${CHOICE}${FALLBACK}Stray`],
        ["text after a Choice with no Fallback", `${CHOICE}Stray`],
        ["text between two Choices", `${CHOICE}Stray${UNKNOWN_CHOICE}${FALLBACK}`],
    ] as const)("rejects a comments root AlternateContent holding %s", async (_shape, inner) => {
        expect(await malformed(commentsRoot(inner), "word/comments.xml")).toEqual([
            expect.objectContaining({ severity: "error", path: "word/comments.xml" }),
        ]);
    });

    it("rejects a body run-level AlternateContent holding text between its Choice and its Fallback (session 67 s67b)", async () => {
        expect(await malformed(bodyRun("Stray"), "word/document.xml")).toEqual([
            expect.objectContaining({ severity: "error", path: "word/document.xml" }),
        ]);
    });

    // The calibrations the same sessions opened clean: the text inside a
    // branch (session 65 s65b), whitespace only between the branches, and
    // the same roots without any text (session 67 s67b).
    it.each([
        ["text inside the Choice before the root", CHOICE.replace("<w:comments", "Stray<w:comments") + FALLBACK],
        ["text inside the Fallback", `${CHOICE}<mc:Fallback>Stray</mc:Fallback>`],
        ["whitespace only between the branches", `${CHOICE}\n  ${FALLBACK}`],
        ["a Choice with no Fallback and no text", CHOICE],
        ["two Choices and no text", `${CHOICE}${UNKNOWN_CHOICE}${FALLBACK}`],
    ] as const)("accepts a comments root AlternateContent holding %s", async (_shape, inner) => {
        expect(await check.run(ValidationContext.fromEntries(commentsRoot(inner)))).toEqual([]);
    });

    it("accepts the body calibration without the text (session 67 s67b)", async () => {
        expect(await check.run(ValidationContext.fromEntries(bodyRun("")))).toEqual([]);
    });

    // #1016: sessions 65 and 67 stored Transitional packages.
    it("cites sessions 65 and 67 in a Transitional package and drops them beside a Strict part (#1016)", async () => {
        expect((await malformed(bodyRun("Stray"), "word/document.xml")).map((issue) => issue.message)).toEqual([
            "AlternateContent holds character data directly inside it; Word 16.115 offered recovery for every position probed (sessions 65 and 67).",
        ]);
        const strictApp = {
            ...bodyRun("Stray"),
            "docProps/app.xml": '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"/>',
        };
        const strictComments = {
            ...commentsRoot(`${CHOICE}Stray${FALLBACK}`),
            "word/styles.xml": '<w:styles xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"/>',
        };
        expect((await malformed(strictApp, "word/document.xml")).map((issue) => issue.message)).toEqual([
            "AlternateContent holds character data directly inside it.",
        ]);
        expect((await malformed(strictComments, "word/comments.xml")).map((issue) => issue.message)).toEqual([
            "AlternateContent holds character data directly inside it.",
        ]);
    });
});

describe("Word-read XML parts regardless of stored name (#1217)", () => {
    for (const profile of ["word-valid", "strict"] as const) {
        for (const main of ["word/document.bin", "other/main.xml", "WORD/doc%75ment.BIN"]) {
            for (const [code, xml] of [
                ["mce-must-understand", document("", MUST_UNDERSTAND_FOO)],
                ["mce-malformed", document(FALLBACK_ONLY)],
            ]) {
                it(`judges ${code} in ${main} under ${profile}`, async () => {
                    const ctx = ValidationContext.fromEntries(compatibilityMainParts(main!, xml));
                    expect(await check.run(ctx, profile)).toEqual([expect.objectContaining({ code, path: main, severity: "error" })]);
                });
            }
        }
        for (const [segment, tag] of [
            ["glossaryDocument", "glossaryDocument"],
            ["header", "hdr"],
            ["footer", "ftr"],
            ["comments", "comments"],
            ["footnotes", "footnotes"],
            ["endnotes", "endnotes"],
        ]) {
            it(`judges a read ${segment} stored outside word/ as binary under ${profile}`, async () => {
                const part = `other/${segment}.bin`;
                const ctx = ValidationContext.fromEntries(
                    compatibilityStoryParts(part, segment!, `<w:${tag} ${ROOT}${MUST_UNDERSTAND_FOO}/>`),
                );
                expect(await check.run(ctx, profile)).toEqual([
                    expect.objectContaining({ code: "mce-must-understand", path: part, severity: "error" }),
                ]);
            });
        }
    }
    it("does not judge an unused binary header or an untyped binary main document", async () => {
        const unused = compatibilityStoryParts("word/header1.bin", "header", `<w:hdr ${ROOT}${MUST_UNDERSTAND_FOO}/>`, false);
        expect(await check.run(ValidationContext.fromEntries(unused), "word-valid")).toEqual([]);
        const untyped = compatibilityMainParts("word/document.bin", document(FALLBACK_ONLY));
        untyped["[Content_Types].xml"] = untyped["[Content_Types].xml"]!.replace(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "application/octet-stream",
        );
        expect(await check.run(ValidationContext.fromEntries(untyped), "word-valid")).toEqual([]);
    });
    for (const strict of [false, true]) {
        it(`keeps character-data session citations conditional for a binary main, Strict sibling: ${strict}`, async () => {
            const parts = compatibilityMainParts(
                "word/document.bin",
                document('<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice>Stray<mc:Fallback/></mc:AlternateContent>'),
            );
            if (strict) parts["docProps/app.xml"] = '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"/>';
            expect(await check.run(ValidationContext.fromEntries(parts), "word-valid")).toEqual([
                expect.objectContaining({
                    code: "mce-malformed",
                    path: "word/document.bin",
                    message: strict
                        ? "AlternateContent holds character data directly inside it."
                        : "AlternateContent holds character data directly inside it; Word 16.115 offered recovery for every position probed (sessions 65 and 67).",
                }),
            ]);
        });
    }
});
