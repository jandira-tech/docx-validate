/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { glossaryStoryParts } from "./support/glossary-content-type-specimens";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { MISMATCH_STORY_CASES, storyContentType, storyContentTypeParts } from "./support/story-content-type-specimens";

const ID = "word-part-content-type-mismatch";
const diagnose = async (parts: Record<string, string>) =>
    (await createDefaultCheckRegistry().get(ID)?.run(ValidationContext.fromEntries(parts))) ?? [];

describe("story content type mismatch (#600, session 24)", () => {
    it("registers a DOCX content-types check behind xml-syntax", () => {
        expect(createDefaultCheckRegistry().get(ID)).toMatchObject({
            id: ID,
            groups: ["content-types"],
            formats: ["docx"],
            requires: ["xml-syntax"],
        });
    });

    it.each(MISMATCH_STORY_CASES)("rejects $via $segment declared $declared", async ({ segment, via, declared }) => {
        const found = await diagnose(storyContentTypeParts(segment, declared, via));
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({
            code: ID,
            severity: "error",
            path: `word/${via === "glossary" ? "glossary/" : ""}${segment}9.xml`,
        });
        expect(found[0]?.message).toContain(`the ${via === "glossary" ? "glossary" : "main document"}'s ${segment} relationship`);
        expect(found[0]?.message).toContain(declared);
        expect(found[0]?.message).toContain(storyContentType(segment));
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package");
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
        expect(await diagnose(storyContentTypeParts(segment, storyContentType(segment), via))).toEqual([]);
    });

    it.each(["header", "footer"] as const)("ignores unread %s swaps", async (segment) => {
        expect(await diagnose(storyContentTypeParts(segment, undefined, "main", false))).toEqual([]);
        expect(await diagnose(storyContentTypeParts(segment, storyContentType(segment).toUpperCase()))).toEqual([]);
    });

    it.each([storyContentType("styles"), "image/png"])("keeps an unobserved footer type %s silent", async (declared) => {
        expect(await diagnose(storyContentTypeParts("footer", declared))).toEqual([]);
    });

    it.each(["comments", "footnotes", "endnotes"] as const)("keeps %s typed as each other notes/comment part silent", async (segment) => {
        for (const other of ["comments", "footnotes", "endnotes"].filter((value) => value !== segment)) {
            const parts = glossaryStoryParts(segment, true);
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(storyContentType(segment), storyContentType(other));
            expect(await diagnose(parts)).toEqual([]);
        }
    });

    it.each([true, false])("accepts any matching reaching type (header first: %s)", async (headerFirst) => {
        const parts = storyContentTypeParts("footer");
        const rels = parts["word/_rels/document.xml.rels"]!;
        const header =
            '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="footer9.xml"/>';
        parts["word/_rels/document.xml.rels"] = headerFirst
            ? rels.replace('<Relationship Id="rIdX"', `${header}<Relationship Id="rIdX"`)
            : rels.replace("</Relationships>", `${header}</Relationships>`);
        expect(await diagnose(parts)).toEqual([]);
    });

    it("keeps generic, inexact and mismatch findings disjoint", async () => {
        for (const [declared, expected] of [
            ["application/xml", "word-part-content-type-generic"],
            [storyContentType("styles") + "; charset=utf-8", "word-part-content-type-inexact"],
            [storyContentType("header"), ID],
        ] as const) {
            const parts =
                expected === "word-part-content-type-inexact" || expected === "word-part-content-type-generic"
                    ? {
                          ...MINIMAL_PARTS,
                          "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(storyContentType("styles"), declared),
                      }
                    : storyContentTypeParts("footer", declared);
            const ctx = ValidationContext.fromEntries(parts);
            const issues = (
                await Promise.all(
                    ["word-part-content-type-generic", "word-part-content-type-inexact", ID].map(
                        async (id) => (await createDefaultCheckRegistry().get(id)?.run(ctx)) ?? [],
                    ),
                )
            ).flat();
            expect(issues.map((issue) => issue.code)).toEqual([expected]);
        }
        expect(await diagnose({ ...MINIMAL_PARTS })).toEqual([]);
    });
});

// Session 24: exact observations and their controls, with profile-policy and direct-path coverage.
describe("mismatch profile scope (#600, #589)", () => {
    it.each(MISMATCH_STORY_CASES)(
        "judges $via $segment in full profiles",
        async ({ segment, via, declared }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "story.docx");
                writeFileSync(file, await buildDocx(storyContentTypeParts(segment, declared, via)));
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const result = await validate(file, { profile });
                    const found = result.issues.filter((issue) => issue.code === ID);
                    expect(found).toHaveLength(1);
                    expect(found[0]?.severity).toBe(profile === "lenient" ? "warning" : "error");
                    if (profile !== "lenient") expect(result.valid).toBe(false);
                    if (profile === "strict") expect(found[0]?.message).toContain("clause 11.3");
                }
                writeFileSync(file, await buildDocx(storyContentTypeParts(segment, storyContentType(segment), via)));
                expect((await validate(file)).valid).toBe(true);
            });
        },
        60_000,
    );

    it.each([
        ["footer", storyContentType("header"), false],
        ["header", storyContentType("footer"), false],
        ["footer", storyContentType("styles"), true],
        ["footer", "image/png", true],
    ] as const)(
        "strict judges %s typed %s (read: %s) without an unobserved Word claim",
        async (segment, declared, read) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "unobserved.docx");
                writeFileSync(file, await buildDocx(storyContentTypeParts(segment, declared, "main", read)));
                expect((await validate(file)).issues.filter((issue) => issue.code === ID)).toEqual([]);
                const strict = await validate(file, { profile: "strict" });
                const found = strict.issues.filter((issue) => issue.code === ID);
                expect(found).toHaveLength(declared === "image/png" ? 0 : 1);
                if (found.length) {
                    expect(found[0]?.message).toContain("clause 11.3");
                    expect(found[0]?.message).not.toContain("Word 16.115");
                }
            });
        },
        60_000,
    );
});

it("leaves missing, malformed and unresolved content types to their owning checks", async () => {
    const parts = storyContentTypeParts("footer");
    const missing = { ...parts };
    delete missing["[Content_Types].xml"];
    expect(await diagnose(missing)).toEqual([]);
    expect(await diagnose({ ...parts, "[Content_Types].xml": "<Types" })).toEqual([]);
    expect(
        await diagnose({
            ...parts,
            "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        }),
    ).toEqual([]);
});

it("reports an uppercase foreign media type supplied by a Default", async () => {
    const parts = storyContentTypeParts("footer");
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
        `<Override PartName="/word/footer9.xml" ContentType="${storyContentType("header")}"/>`,
        "",
    ).replace('ContentType="application/xml"', `ContentType=" ${storyContentType("header").toUpperCase()}; charset=utf-8"`);
    const found = await diagnose(parts);
    expect(found).toHaveLength(1);
    expect(found[0]?.message).toContain("Default");
});

it("strict judges every mapped related type on the main document and glossary", async () => {
    const segments = [
        "styles",
        "settings",
        "fontTable",
        "webSettings",
        "numbering",
        "theme",
        "header",
        "footer",
        "footnotes",
        "endnotes",
        "comments",
        "glossaryDocument",
        "officeDocument",
    ];
    await withTempDir(async (dir) => {
        for (const via of ["main", "glossary"] as const) {
            for (const segment of segments) {
                if (via === "glossary" && segment === "officeDocument") continue;
                const parts = storyContentTypeParts("footer", storyContentType("header"), via, false);
                const target = via === "main" ? "word/footer9.xml" : "word/glossary/footer9.xml";
                if (segment === "officeDocument") {
                    parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
                        storyContentType("document.main"),
                        storyContentType("header"),
                    );
                    parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
                        `PartName="/${target}" ContentType="${storyContentType("header")}"`,
                        `PartName="/${target}" ContentType="${storyContentType("footer")}"`,
                    );
                } else {
                    const rels = via === "main" ? "word/_rels/document.xml.rels" : "word/glossary/_rels/document.xml.rels";
                    parts[rels] = parts[rels]!.replace('/footer"', `/${segment}"`);
                    if (segment === "header")
                        parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
                            `PartName="/${target}" ContentType="${storyContentType("header")}"`,
                            `PartName="/${target}" ContentType="${storyContentType("footer")}"`,
                        );
                }
                const file = path.join(dir, "mapped.docx");
                writeFileSync(file, await buildDocx(parts));
                const found = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
                expect(found, `${via} ${segment}`).toHaveLength(1);
                expect(found[0]?.message).toContain(segment === "theme" ? "clause 14.2.7" : "clause 11.3");
                expect(found[0]?.message).not.toContain("Word 16.115");
            }
        }
    });
}, 60_000);
