/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { runChecks } from "../src/lib/check-scheduler";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { fallback, PACKAGE_RELS, wrapRelationship } from "./support/relationships-markup-compatibility-specimens";

const CODE = "word-document-root";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN = '<x:foo xmlns:x="urn:x"/>';
const IGNORED = `<x:document xmlns:x="urn:future:document" xmlns:w="${W}" xmlns:mc="${MC}" mc:Ignorable="x"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></x:document>`;
const relationships = (attributes: string) => `<Relationships xmlns="${PR}"><Relationship Id="main" ${attributes}/></Relationships>`;

const rootIssues = async (parts: Readonly<Record<string, string>>) => {
    const ctx = ValidationContext.fromEntries(parts);
    const check = createDefaultCheckRegistry()
        .list()
        .find((definition) => definition.id === CODE);
    return (await check?.run(ctx)) ?? [];
};

describe(CODE, () => {
    it("judges the main root through the selected package Fallback (#632)", async () => {
        const parts = wrapRelationship({ ...MINIMAL_PARTS, "word/document.xml": FOREIGN }, PACKAGE_RELS, "rId1", fallback);
        expect(await rootIssues(parts)).toEqual([expect.objectContaining({ code: CODE, path: "word/document.xml", severity: "error" })]);
    });

    it("does not judge a main root through a nested package relationship (#434)", async () => {
        const rels = MINIMAL_PARTS["_rels/.rels"]!.replace(/(<Relationship\s[^>]*\/>)/, '<x:Wrap xmlns:x="urn:x">$1</x:Wrap>');
        expect(await rootIssues({ ...MINIMAL_PARTS, "_rels/.rels": rels, "word/document.xml": FOREIGN })).toEqual([]);
    });

    it.each([FOREIGN, `<x:document xmlns:x="urn:x"/>`, `<w:foo xmlns:w="${W}"/>`, "<document/>"])(
        "rejects a non-document effective root: %s",
        async (xml) => {
            expect(await rootIssues({ ...MINIMAL_PARTS, "word/document.xml": xml })).toEqual([
                expect.objectContaining({
                    code: CODE,
                    severity: "error",
                    path: "word/document.xml",
                    message: expect.stringMatching(/Microsoft Word refuses to open/),
                }),
            ]);
        },
    );

    it("rejects an ignored root because Word silently drops all content", async () => {
        expect(await rootIssues({ ...MINIMAL_PARTS, "word/document.xml": IGNORED })).toEqual([
            expect.objectContaining({
                code: CODE,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/Microsoft Word.*drops all content/),
            }),
        ]);
    });

    it.each([W, STRICT_W])("accepts document in %s, regardless of prefix", async (namespace) => {
        expect(await rootIssues({ ...MINIMAL_PARTS, "word/document.xml": `<document xmlns="${namespace}"><body/></document>` })).toEqual(
            [],
        );
    });

    it("accepts the Transitional minimal control", async () => {
        expect(await rootIssues(MINIMAL_PARTS)).toEqual([]);
    });

    it.each([R, "http://purl.oclc.org/ooxml/officeDocument/relationships"])(
        "finds a renamed main part through %s/officeDocument, normalizing its target",
        async (namespace) => {
            expect(
                await rootIssues({
                    ...MINIMAL_PARTS,
                    "_rels/.rels": relationships(`Type="${namespace}/officeDocument" Target="/renamed/../renamed/main.xml"`),
                    "renamed/main.xml": FOREIGN,
                }),
            ).toEqual([expect.objectContaining({ code: CODE, path: "renamed/main.xml", severity: "error" })]);
        },
    );

    it.each(["WORD/document.xml", "/Word/Document.XML", "word/%64ocument.xml"])(
        "finds the main part behind the OPC-equivalent target %s (#344)",
        async (target) => {
            expect(
                await rootIssues({
                    ...MINIMAL_PARTS,
                    "_rels/.rels": relationships(`Type="${R}/officeDocument" Target="${target}"`),
                    "word/document.xml": FOREIGN,
                }),
            ).toEqual([expect.objectContaining({ code: CODE, path: "word/document.xml", severity: "error" })]);
        },
    );

    it.each(["word/document.xml#frag", "word/document.xml?x=1"])(
        "finds the main part behind %s, whose query or fragment is not part of the path (#361)",
        async (target) => {
            expect(
                await rootIssues({
                    ...MINIMAL_PARTS,
                    "_rels/.rels": relationships(`Type="${R}/officeDocument" Target="${target}"`),
                    "word/document.xml": FOREIGN,
                }),
            ).toEqual([expect.objectContaining({ code: CODE, path: "word/document.xml", severity: "error" })]);
        },
    );

    it("reads package relationships stored under a case-variant name (#364)", async () => {
        const { "_rels/.rels": packageRelationships, ...rest } = MINIMAL_PARTS;
        expect(await rootIssues({ ...rest, "_RELS/.RELS": packageRelationships, "word/document.xml": FOREIGN })).toEqual([
            expect.objectContaining({ code: CODE, path: "word/document.xml", severity: "error" }),
        ]);
    });

    it("does not equate an escaped reserved character with its literal (#344)", async () => {
        // `%2F` is an escaped "/", so this target names a different part.
        expect(
            await rootIssues({
                ...MINIMAL_PARTS,
                "_rels/.rels": relationships(`Type="${R}/officeDocument" Target="word%2Fdocument.xml"`),
                "word/document.xml": FOREIGN,
            }),
        ).toEqual([]);
    });

    it("leaves a non-main ignored root as processor info", async () => {
        const parts = { ...MINIMAL_PARTS, "word/styles.xml": IGNORED };
        expect(await rootIssues(parts)).toEqual([]);
        expect(ValidationContext.fromEntries(parts).compatibilityIssues("word/styles.xml")).toContainEqual(
            expect.objectContaining({ code: "mce-root-ignored", severity: "info" }),
        );
    });

    it.each<Readonly<Record<string, string>>>([
        {},
        { "_rels/.rels": "<Relationships" },
        { "_rels/.rels": relationships(`Type="${R}/styles" Target="word/document.xml"`) },
        { "_rels/.rels": relationships(`Type="${R}/officeDocument" Target="word/document.xml" TargetMode="External"`) },
        { "_rels/.rels": relationships(`Type="${R}/officeDocument"`) },
        { "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]! },
    ])("leaves absent, malformed, external, or unrelated package targets to their checks: %j", async (parts) => {
        expect(await rootIssues(parts)).toEqual([]);
    });

    it("leaves malformed main XML to xml-syntax, including recoverable DOM input", async () => {
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "word/document.xml": '<x:foo xmlns:x="urn:x">a & b</x:foo>' });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["xml"] });
        expect(run.issues).toContainEqual(expect.objectContaining({ code: "xml-syntax", path: "word/document.xml" }));
        expect(run.issues.filter((issue) => issue.code === CODE)).toEqual([]);
        expect(run.outcomes).toContainEqual(expect.objectContaining({ id: CODE, status: "passed", issues: [] }));
    });
});
