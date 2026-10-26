/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, CT, MC, R, RS } from "./support/word-rule-test-xml";
import { wordContentTypeInvalidCheck } from "../src/scripts/office/checks/word-content-types";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { contentTypeSpecimens } from "./support/word-content-types-specimens";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { glossaryStylesParts } from "./support/glossary-content-type-specimens";
import { RELATED_PART_SPECIMENS, relatedContentTypeParts } from "./support/related-content-type-specimens";

it("registers duplicate declarations for all formats, reads the Markup-Compatibility-processed stream and skips absent/unparseable streams (#431, #459)", async () => {
    const check = createDefaultCheckRegistry().get("ct-duplicate-declaration");
    expect(check).toMatchObject({
        id: "ct-duplicate-declaration",
        groups: ["content-types"],
        formats: ["docx", "pptx", "xlsx"],
        requires: ["xml-syntax"],
    });
    const stream = "[CONTENT_TYPES].XML";
    const xml = `<Types xmlns="${CT}"><Default Extension="xml"/><Default Extension="XML"/></Types>`;
    const ctx = ValidationContext.fromEntries({ [stream]: xml });
    expect(await check!.run(ctx)).toEqual([expect.objectContaining({ code: "ct-duplicate-declaration", path: stream, severity: "error" })]);
    expect(ctx.text(stream)).toBe(xml);
    for (const context of [
        ValidationContext.fromEntries({}),
        ValidationContext.fromEntries({ [stream]: "<broken" }),
        ctx.withUnparseable(new Map([[stream, "malformed"]])),
    ]) {
        expect(await check!.run(context)).toEqual([]);
    }
});

describe("word-content-type-invalid", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            "[Content_Types].xml": `<Types xmlns="${CT}"><Override PartName="/word/document.xml" ContentType="invalid-document"/></Types>`,
        });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-content-type-invalid",
                path: "[Content_Types].xml",
                message: "Microsoft Word refuses content type 'invalid-document' for /word/document.xml",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "[Content_Types].xml": `<Types xmlns="${CT}"><Override PartName="/word/document.xml" ContentType="application/xml"/><Default Extension="xml"/></Types>`,
        });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordContentTypeInvalidCheck.run(
                ValidationContext.fromEntries({ "[Content_Types].xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("checks default extensions and leaves nonmatching/absent types alone", async () => {
    const ctx = ValidationContext.fromEntries({
        ...MINIMAL_PARTS,
        "word/data.bin": "payload",
        "[Content_Types].xml": `<Types xmlns="${CT}"><Default Extension="bin" ContentType="invalid-bin"/><Default/><Override ContentType="invalid-part"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
        {
            severity: "error",
            code: "word-content-type-invalid",
            path: "[Content_Types].xml",
            message: "Microsoft Word refuses content type 'invalid-bin' for bin",
        },
    ]);
});

it.each(["[CONTENT_TYPES].XML", "[content_types].xml"])("finds invalid types in stored stream %s (#396)", async (name) => {
    const ctx = ValidationContext.fromEntries({
        "word/x.xml": MINIMAL_PARTS["word/document.xml"]!,
        [name]: `<Types xmlns="${CT}"><Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
        {
            severity: "error",
            code: "word-content-type-invalid",
            path: name,
            message: "Microsoft Word refuses content type 'invalid-type' for /word/x.xml",
        },
    ]);
});

describe("stored content-type scope (#401)", () => {
    it.each(contentTypeSpecimens)("$name", async ({ parts, rejected }) => {
        const ctx = await ValidationContext.fromArchive(await buildDocx(parts));
        const findings = await wordContentTypeInvalidCheck.run(ctx);
        expect(findings).toHaveLength(rejected ? 1 : 0);
        if (rejected) {
            expect(findings[0]).toMatchObject({ code: "word-content-type-invalid", severity: "error", path: "[Content_Types].xml" });
        }
    });
});

it("does not read a percent-escaped name as the stream (#396)", async () => {
    const ctx = ValidationContext.fromEntries({
        "%5BContent_Types%5D.xml": `<Types xmlns="${CT}"><Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([]);
});

// Word 16.115.26100424 offered recovery for the empty main Override:
// session 22 in docs/plans/evidence/word-probes-2026-10-24.md (#581).
describe("empty main-document Override (#581)", () => {
    const mainType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const emptyStream = MINIMAL_PARTS["[Content_Types].xml"]!.replace(mainType, "");
    const finding = (part = "word/document.xml", stream = "[Content_Types].xml") => ({
        code: "word-content-type-invalid",
        severity: "error",
        path: stream,
        message: `Microsoft Word refuses empty content type '' for ${part} (main document Override)`,
    });

    it("reports the empty value and the main part once, without changing the stream", async () => {
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "[Content_Types].xml": emptyStream });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([finding()]);
        expect(ctx.text("[Content_Types].xml")).toBe(emptyStream);
    });

    it.each([R, RS])("follows %s/officeDocument to a renamed main part using OPC equivalence", async (namespace) => {
        const parts = { ...MINIMAL_PARTS };
        delete parts["word/document.xml"];
        delete parts["[Content_Types].xml"];
        const ctx = ValidationContext.fromEntries({
            ...parts,
            "renamed/main.xml": MINIMAL_PARTS["word/document.xml"]!,
            "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${namespace}/officeDocument" Target="RENAMED/%6dain.xml"/></Relationships>`,
            "[CONTENT_TYPES].XML": emptyStream.replace("/word/document.xml", "/RENAMED/main.xml"),
        });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([finding("renamed/main.xml", "[CONTENT_TYPES].XML")]);
    });

    it.each([
        ["proper control", MINIMAL_PARTS["[Content_Types].xml"]!],
        [
            "empty Override on an unstored part",
            MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", '<Override PartName="/word/absent.xml" ContentType=""/></Types>'),
        ],
        ["empty Default", MINIMAL_PARTS["[Content_Types].xml"]!.replace('ContentType="application/xml"', 'ContentType=""')],
        [
            "main typed by empty Default",
            emptyStream
                .replace('<Override PartName="/word/document.xml" ContentType=""/>', "")
                .replace('ContentType="application/xml"', 'ContentType=""'),
        ],
        [
            "missing ContentType attribute",
            emptyStream.replace('PartName="/word/document.xml" ContentType=""', 'PartName="/word/document.xml"'),
        ],
        [
            "whitespace is unobserved",
            emptyStream.replace('PartName="/word/document.xml" ContentType=""', 'PartName="/word/document.xml" ContentType=" "'),
        ],
        [
            "first Override wins",
            emptyStream.replace(
                '<Override PartName="/word/document.xml"',
                `<Override PartName="/word/document.xml" ContentType="${mainType}"/><Override PartName="/word/document.xml"`,
            ),
        ],
    ])("leaves %s outside the empty-value finding", async (_label, stream) => {
        expect(
            await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries({ ...MINIMAL_PARTS, "[Content_Types].xml": stream })),
        ).toEqual([]);
    });

    it.each([
        ["missing", undefined],
        ["malformed", "<broken"],
        [
            "external",
            MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', 'Target="word/document.xml" TargetMode="External"'),
        ],
        ["unknown relationship", MINIMAL_PARTS["_rels/.rels"]!.replace("/officeDocument", "/unknown")],
    ])("does not infer a main part from %s package relationships", async (_label, rels) => {
        const parts: Record<string, string> = { ...MINIMAL_PARTS, "[Content_Types].xml": emptyStream };
        delete parts["_rels/.rels"];
        if (rels !== undefined) parts["_rels/.rels"] = rels;
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it.each([false, true])("judges the effective stream with empty fallback selected=%s", async (selected) => {
        const proper = `<Override PartName="/word/document.xml" ContentType="${mainType}"/>`;
        const empty = '<Override PartName="/word/document.xml" ContentType=""/>';
        const stream = `<Types xmlns="${CT}" xmlns:mc="${MC}" xmlns:f="urn:unobserved"><mc:AlternateContent><mc:Choice Requires="f">${selected ? proper : empty}</mc:Choice><mc:Fallback>${selected ? empty : proper}</mc:Fallback></mc:AlternateContent></Types>`;
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "[Content_Types].xml": stream });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual(selected ? [finding()] : []);
    });
});

// Sessions 49 and 50 (2026-10-26): Word 16.115 offered recovery for an empty
// ContentType on the Override of the styles part the main document relates
// (s49b) and of a related settings or fontTable part (s50c).
describe("empty related Override (#720, sessions 49 and 50)", () => {
    const stylesType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const stylesEmptyParts = (): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(stylesType, ""),
    });

    it.each([
        {
            label: "the styles part the main document relates",
            parts: stylesEmptyParts(),
            part: "word/styles.xml",
            segment: "styles",
            session: 49,
        },
        {
            label: "a related settings part",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], ""),
            part: "word/settings.xml",
            segment: "settings",
            session: 50,
        },
        {
            label: "a related fontTable part",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[1], ""),
            part: "word/fontTable.xml",
            segment: "fontTable",
            session: 50,
        },
    ])("reports $label exactly once with its session", async ({ parts, part, segment, session }) => {
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries(parts))).toEqual([
            {
                severity: "error",
                code: "word-content-type-invalid",
                path: "[Content_Types].xml",
                message: `Microsoft Word refuses empty content type '' for ${part} (main document ${segment} relationship Override, session ${session})`,
            },
        ]);
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries({ ...MINIMAL_PARTS }))).toEqual([]);
    });

    it("keeps the main document's empty Override and the unshown parts as today", async () => {
        const mainType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        const main = { ...MINIMAL_PARTS, "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(mainType, "") };
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries(main))).toEqual([
            {
                severity: "error",
                code: "word-content-type-invalid",
                path: "[Content_Types].xml",
                message: "Microsoft Word refuses empty content type '' for word/document.xml (main document Override)",
            },
        ]);
        // A reached glossary's styles part is not the main document's own
        // relationship, and a related numbering part was not shown to Word.
        const glossary = glossaryStylesParts(true);
        glossary["[Content_Types].xml"] = glossary["[Content_Types].xml"]!.replace(
            `<Override PartName="/word/glossary/styles.xml" ContentType="${stylesType}"/>`,
            '<Override PartName="/word/glossary/styles.xml" ContentType=""/>',
        );
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries(glossary))).toEqual([]);
        expect(
            await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries(relatedContentTypeParts(RELATED_PART_SPECIMENS[2], ""))),
        ).toEqual([]);
    });
});
