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
import { wordPartContentTypeMismatchFindings } from "../src/scripts/office/checks/word-part-content-type-mismatch";
import { wordPartContentTypeInexactFindings } from "../src/scripts/office/checks/word-part-content-type-inexact";
import { wordPartContentTypeGenericFindings } from "../src/scripts/office/checks/word-part-content-type-generic";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { glossaryStoryParts, glossaryStylesParts } from "./support/glossary-content-type-specimens";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { RELATED_PART_SPECIMENS, relatedContentTypeParts } from "./support/related-content-type-specimens";
import {
    footerWithHeaderRelationship,
    MISMATCH_STORY_CASES,
    storyContentType,
    storyContentTypeParts,
} from "./support/story-content-type-specimens";

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

    it.each([true, false])(
        "accepts a matching type when both footer and header references are USED (header first: %s)",
        async (headerFirst) => {
            const parts = footerWithHeaderRelationship(storyContentType("header"), true, headerFirst);
            expect(await diagnose(parts)).toEqual([]);
            expect(await wordPartContentTypeMismatchFindings(ValidationContext.fromEntries(parts), "spec")).toEqual([]);
        },
    );

    it.each([true, false])(
        "rejects a read footer despite an UNUSED matching header relationship (header first: %s) (#607)",
        async (headerFirst) => {
            const found = await diagnose(footerWithHeaderRelationship(storyContentType("header"), false, headerFirst));
            expect(found).toHaveLength(1);
            expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/footer9.xml" });
            expect(found[0]?.message).toContain("the main document's footer relationship");
            expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 24).");
            expect(await diagnose(footerWithHeaderRelationship(storyContentType("footer"), false, headerFirst))).toEqual([]);
        },
    );

    it("keeps competing unused types disjoint under observed and spec scopes (#607)", async () => {
        for (const scope of ["observed", "spec"] as const) {
            for (const [declared, expected] of [
                ["application/xml", scope === "spec" ? "word-part-content-type-generic" : undefined],
                [storyContentType("footer") + "; charset=utf-8", "word-part-content-type-inexact"],
                [storyContentType("header"), ID],
                [storyContentType("header") + "; charset=utf-8", ID],
                [storyContentType("footer"), undefined],
            ] as const) {
                const ctx = ValidationContext.fromEntries(footerWithHeaderRelationship(declared));
                const findings = (
                    await Promise.all([
                        wordPartContentTypeGenericFindings(ctx, scope),
                        wordPartContentTypeInexactFindings(ctx, scope),
                        wordPartContentTypeMismatchFindings(ctx, scope),
                    ])
                )
                    .flat()
                    .filter(({ issue }) => issue.path === "word/footer9.xml");
                expect(
                    findings.map(({ issue }) => issue.code),
                    `${scope}: ${declared}`,
                ).toEqual(expected === undefined ? [] : [expected]);
            }
        }
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
                // Sessions 49 and 50 made the main document and the main
                // document's styles, settings and numbering reaches observed
                // (#722), so their strict findings keep the Word sentence the
                // swaps already carry (#610); every other shape stays spec-only.
                const session =
                    via === "main" && ["officeDocument", "styles", "settings", "numbering"].includes(segment)
                        ? segment === "styles"
                            ? 49
                            : 50
                        : undefined;
                if (session === undefined) {
                    expect(found[0]?.message).not.toContain("Word 16.115");
                } else {
                    expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (session ${session}).`);
                }
            }
        }
    });
}, 60_000);

// File-backed profile checks are integration tests; the regressions above use in-memory package entries.
it("full, quick and named runs reject the unused-header exemption and accept its footer control (#607)", async () => {
    await withTempDir(async (dir) => {
        const file = path.join(dir, "unused-header.docx");
        for (const declared of [storyContentType("header"), storyContentType("footer")]) {
            writeFileSync(file, await buildDocx(footerWithHeaderRelationship(declared)));
            const full = await validate(file);
            const found = full.issues.filter(({ code }) => code === ID);
            expect(full.valid).toBe(declared === storyContentType("footer"));
            expect(found).toHaveLength(declared === storyContentType("header") ? 1 : 0);
            for (const selection of [{ quick: true }, { checks: [ID] }, { checks: ["content-types"] }]) {
                const partial = await validate(file, selection);
                expect(partial.valid).toBe(full.valid);
                expect(partial.issues.filter(({ code }) => code === ID)).toEqual(found);
            }
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter(({ code }) => code === ID)).toHaveLength(found.length);
            if (!full.valid) expect(strict.valid).toBe(false);
        }
    });
}, 60_000);

it.each([
    { segment: "theme", clause: "14.2.7" },
    { segment: "header", clause: "11.3" },
])(
    "strict mismatch issues rank $segment before footer regardless of relationship order (#608)",
    async ({ segment, clause }) => {
        await withTempDir(async (dir) => {
            const issues = [];
            for (const otherFirst of [true, false]) {
                const parts = storyContentTypeParts("footer", storyContentType("styles"), "main", false);
                const rels = parts["word/_rels/document.xml.rels"]!;
                const other = `<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${segment}" Target="footer9.xml"/>`;
                parts["word/_rels/document.xml.rels"] = otherFirst
                    ? rels.replace('<Relationship Id="rIdX"', `${other}<Relationship Id="rIdX"`)
                    : rels.replace("</Relationships>", `${other}</Relationships>`);
                const file = path.join(dir, "ranked.docx");
                writeFileSync(file, await buildDocx(parts));
                issues.push((await validate(file, { profile: "strict" })).issues.filter(({ code }) => code === ID));
            }
            expect(issues[0]).toHaveLength(1);
            expect(issues[1]).toEqual(issues[0]);
            expect(issues[0]?.[0]).toMatchObject({ code: ID, path: "word/footer9.xml" });
            expect(issues[0]?.[0]?.message).toContain(`the main document's ${segment} relationship`);
            expect(issues[0]?.[0]?.message).toContain(`clause ${clause}`);
            expect(issues[0]?.[0]?.message).not.toContain("Word 16.115");
        });
    },
    60_000,
);

it("strict names the Main Document part and its clause 11.3.10 (#609)", async () => {
    const parts = {
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(storyContentType("document.main"), storyContentType("styles")),
    };
    const found = await wordPartContentTypeMismatchFindings(ValidationContext.fromEntries(parts), "spec");
    expect(found).toHaveLength(1);
    expect(found[0]?.issue).toMatchObject({ code: ID, path: "word/document.xml" });
    expect(found[0]?.issue.message).toContain("the package officeDocument relationship");
    expect(found[0]?.issue.message).toContain("Main Document part");
    expect(found[0]?.issue.message).toContain("clause 11.3.10");
    expect(found[0]?.issue.message).not.toContain("officeDocument part");
});

it.each([
    { segment: "footer", via: "main", session: 24 },
    { segment: "header", via: "main", session: 24 },
    { segment: "footer", via: "glossary", session: 24 },
    { segment: "header", via: "glossary", session: 26 },
] as const)("attributes the $via $segment swap to session $session (#610)", async ({ segment, via, session }) => {
    for (const scope of ["observed", "spec"] as const) {
        const found = await wordPartContentTypeMismatchFindings(
            ValidationContext.fromEntries(storyContentTypeParts(segment, undefined, via)),
            scope,
        );
        expect(found).toHaveLength(1);
        expect(found[0]?.issue.message).toContain(`Word 16.115 offers to recover the package (session ${session}).`);
        expect(found[0]?.issue.message).not.toContain(`session ${session === 24 ? 26 : 24}`);
    }
});

// Sessions 49 and 50 (2026-10-26): Word 16.115 offered recovery for the main
// styles part typed with the settings type (s49b), and for a related settings
// part typed with the styles type, a related numbering part typed with the
// settings type and the main document typed with the styles type (s50d).
describe("observed eager mismatches (#722, sessions 49 and 50)", () => {
    const retypedParts = (part: string, own: string, type: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            `<Override PartName="/${part}" ContentType="${own}"/>`,
            `<Override PartName="/${part}" ContentType="${type}"/>`,
        ),
    });
    const CASES = [
        {
            label: "the main styles part typed with the settings type",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("settings")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: 49,
        },
        {
            label: "a related settings part typed with the styles type",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("styles")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: 50,
        },
        {
            label: "a related numbering part typed with the settings type",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("settings")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: 50,
        },
        {
            label: "the main document typed with the styles type",
            parts: retypedParts("word/document.xml", storyContentType("document.main"), storyContentType("styles")),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            session: 50,
        },
    ] as const;

    it.each(CASES)("rejects $label under word-valid with its session", async ({ parts, path, by, session }) => {
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
        expect(found[0]?.message).toContain(by);
        expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (session ${session}).`);
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
        expect(await diagnose({ ...MINIMAL_PARTS })).toEqual([]);
    });

    it("invalidates the full word-valid run and keeps its control valid", async () => {
        await withTempDir(async (dir) => {
            for (const { label, parts } of CASES) {
                const file = path.join(dir, "observed.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file);
                expect(result.valid, label).toBe(false);
                expect(result.issues.filter((issue) => issue.code === ID && issue.severity === "error")).toHaveLength(1);
            }
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS }));
            expect((await validate(file)).valid).toBe(true);
        });
    }, 60_000);

    it("keeps the header/footer swaps on their sessions and an unshown glossary styles reach silent", async () => {
        const found = await diagnose(storyContentTypeParts("footer"));
        expect(found).toHaveLength(1);
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 24).");
        const glossary = glossaryStylesParts(true);
        glossary["[Content_Types].xml"] = glossary["[Content_Types].xml"]!.replace(
            `<Override PartName="/word/glossary/styles.xml" ContentType="${storyContentType("styles")}"/>`,
            `<Override PartName="/word/glossary/styles.xml" ContentType="${storyContentType("settings")}"/>`,
        );
        expect(await diagnose(glossary)).toEqual([]);
        const spec = await wordPartContentTypeMismatchFindings(ValidationContext.fromEntries(glossary), "spec");
        expect(spec).toHaveLength(1);
        expect(spec[0]?.issue.message).toContain("the glossary's styles relationship");
        expect(spec[0]?.issue.message).not.toContain("Word 16.115");
    });
});
