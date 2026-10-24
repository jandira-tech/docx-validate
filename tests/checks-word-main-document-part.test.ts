/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const CODE = "word-main-document-part";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STRICT_OFFICE_DOCUMENT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const relationships = (...attributes: string[]) =>
    `<Relationships xmlns="${PR}">${attributes.map((attrs, i) => `<Relationship Id="rId${i + 1}" ${attrs}/>`).join("")}</Relationships>`;
const without = (...names: string[]) => Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => !names.includes(name)));

const mainPartIssues = async (parts: Readonly<Record<string, string>>) => {
    const check = createDefaultCheckRegistry()
        .list()
        .find((definition) => definition.id === CODE);
    expect(check).toBeDefined();
    return (await check!.run(ValidationContext.fromEntries(parts))) ?? [];
};

describe(CODE, () => {
    it("does not count a nested main relationship (#434)", async () => {
        const rels = MINIMAL_PARTS["_rels/.rels"]!.replace(/(<Relationship\s[^>]*\/>)/, '<x:Wrap xmlns:x="urn:x">$1</x:Wrap>');
        expect(await mainPartIssues({ ...MINIMAL_PARTS, "_rels/.rels": rels })).toEqual([
            expect.objectContaining({ code: CODE, severity: "error", path: "_rels/.rels" }),
        ]);
    });

    it("is a registered Word rule in the relationships group, so --quick runs it", () => {
        const check = WORD_RULE_CHECKS.find((definition) => definition.id === CODE);
        expect(check).toMatchObject({ formats: ["docx"], groups: ["relationships"] });
    });

    it("accepts the minimal package", async () => {
        expect(await mainPartIssues(MINIMAL_PARTS)).toEqual([]);
    });

    it.each([OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT])("accepts a main part named by %s", async (type) => {
        expect(
            await mainPartIssues({ ...MINIMAL_PARTS, "_rels/.rels": relationships(`Type="${type}" Target="/word/document.xml"`) }),
        ).toEqual([]);
    });

    it("accepts an OPC-equivalent spelling of the main part's name", async () => {
        expect(
            await mainPartIssues({
                ...MINIMAL_PARTS,
                "_rels/.rels": relationships(`Type="${OFFICE_DOCUMENT}" Target="WORD/%64ocument.xml"`),
            }),
        ).toEqual([]);
    });

    it("accepts package relationships stored under an equivalent name", async () => {
        const { "_rels/.rels": rels, ...rest } = MINIMAL_PARTS;
        expect(await mainPartIssues({ ...rest, "_RELS/.RELS": rels! })).toEqual([]);
    });

    it.each([
        ["an empty package", {}],
        ["a package holding only [Content_Types].xml", { "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]! }],
        ["a package without _rels/.rels", without("_rels/.rels")],
    ])("rejects %s: it has no package relationships", async (_label, parts) => {
        expect(await mainPartIssues(parts)).toEqual([
            expect.objectContaining({ code: CODE, severity: "error", message: expect.stringMatching(/_rels\/\.rels/) }),
        ]);
    });

    it.each([
        ["no relationships", relationships()],
        ["only a core-properties relationship", relationships(`Type="${CORE}" Target="docProps/core.xml"`)],
        [
            "only an External officeDocument relationship",
            relationships(`Type="${OFFICE_DOCUMENT}" Target="https://example.com/main.xml" TargetMode="External"`),
        ],
    ])("rejects package relationships with %s", async (_label, rels) => {
        expect(await mainPartIssues({ ...MINIMAL_PARTS, "_rels/.rels": rels })).toEqual([
            expect.objectContaining({
                code: CODE,
                severity: "error",
                path: "_rels/.rels",
                message: expect.stringMatching(/officeDocument/),
            }),
        ]);
    });

    it.each(["word/missing.xml", "../word/document.xml", "", "word/document.xml/"])(
        "rejects an officeDocument target that names no part: %j",
        async (target) => {
            expect(
                await mainPartIssues({ ...MINIMAL_PARTS, "_rels/.rels": relationships(`Type="${OFFICE_DOCUMENT}" Target="${target}"`) }),
            ).toEqual([expect.objectContaining({ code: CODE, severity: "error", path: "_rels/.rels" })]);
        },
    );

    it("counts relationships even when one target is missing (#515)", async () => {
        expect(
            await mainPartIssues({
                ...MINIMAL_PARTS,
                "_rels/.rels": relationships(
                    `Type="${OFFICE_DOCUMENT}" Target="word/missing.xml"`,
                    `Type="${OFFICE_DOCUMENT}" Target="word/document.xml"`,
                ),
            }),
        ).toEqual([
            expect.objectContaining({ code: CODE, severity: "error", path: "_rels/.rels", message: expect.stringMatching(/rId1.*rId2/) }),
        ]);
    });

    it.each([
        [OFFICE_DOCUMENT, OFFICE_DOCUMENT],
        [STRICT_OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT],
        [OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT],
        [STRICT_OFFICE_DOCUMENT, OFFICE_DOCUMENT],
    ])("counts the two officeDocument spellings together: %s and %s (#515)", async (first, second) => {
        expect(
            await mainPartIssues({
                ...MINIMAL_PARTS,
                "_rels/.rels": relationships(`Type="${first}" Target="word/document.xml"`, `Type="${second}" Target="word/document.xml"`),
            }),
        ).toEqual([
            expect.objectContaining({ code: CODE, severity: "error", path: "_rels/.rels", message: expect.stringMatching(/rId1.*rId2/) }),
        ]);
    });

    it("names every internal officeDocument Id, excluding external and unrelated relationships (#515)", async () => {
        const { "_rels/.rels": _rels, ...parts } = MINIMAL_PARTS;
        const issues = await mainPartIssues({
            ...parts,
            "_RELS/.RELS": relationships(
                `Type="${OFFICE_DOCUMENT}" Target="word/document.xml"`,
                `Type="${CORE}" Target="docProps/core.xml"`,
                `Type="${OFFICE_DOCUMENT}" Target="https://example.com/main.xml" TargetMode="External"`,
                `Type="${STRICT_OFFICE_DOCUMENT}" Target="word/document.xml"`,
                `Type="${OFFICE_DOCUMENT}" Target="word/document.xml"`,
            ),
        });
        expect(issues).toEqual([
            expect.objectContaining({
                code: CODE,
                severity: "error",
                path: "_RELS/.RELS",
                message: expect.stringMatching(/rId1.*rId4.*rId5/),
            }),
        ]);
        expect(issues[0]!.message).not.toMatch(/rId2|rId3/);
    });

    it("ignores additional External officeDocument relationships (#515)", async () => {
        expect(
            await mainPartIssues({
                ...MINIMAL_PARTS,
                "_rels/.rels": relationships(
                    `Type="${OFFICE_DOCUMENT}" Target="word/document.xml"`,
                    `Type="${OFFICE_DOCUMENT}" Target="https://example.com/main.xml" TargetMode="External"`,
                    `Type="${STRICT_OFFICE_DOCUMENT}" Target="https://example.com/other.xml" TargetMode="External"`,
                ),
            }),
        ).toEqual([]);
    });

    it("leaves malformed package relationships to the XML syntax check", async () => {
        expect(await mainPartIssues({ ...MINIMAL_PARTS, "_rels/.rels": "<Relationships" })).toEqual([]);
    });
});
