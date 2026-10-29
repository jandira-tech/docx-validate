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
import { climbingFooterTyped } from "./support/climbing-target-specimens";
import { RELATED_PART_SPECIMENS, relatedContentTypeParts } from "./support/related-content-type-specimens";
import {
    footerWithHeaderRelationship,
    MISMATCH_STORY_CASES,
    storyContentType,
    storyContentTypeParts,
    uniformlyStrict,
} from "./support/story-content-type-specimens";

const ID = "word-part-content-type-mismatch";
const diagnose = async (parts: Record<string, string>) =>
    (await createDefaultCheckRegistry().get(ID)?.run(ValidationContext.fromEntries(parts))) ?? [];

// #1016: sessions 24, 26, 49, 50, 53, 56, 57 and 59 stored Transitional
// packages, so in a uniformly Strict package, and in a Transitional one with
// a Strict docProps/app.xml, the swap and eager-pair findings stay without
// the recovery sentence that cites them.
const withStrictApp = (parts: Record<string, string>): Record<string, string> => ({
    ...parts,
    "docProps/app.xml": '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"/>',
});
const strictVariants = (parts: Record<string, string>): Record<string, string>[] => [uniformlyStrict(parts), withStrictApp(parts)];
/** `parts` with only `name` in the Strict WordprocessingML namespace. */
const withStrictPart = (parts: Record<string, string>, name: string): Record<string, string> => ({
    ...parts,
    [name]: parts[name]!.replaceAll(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ),
});
/** `parts` with only `word/styles.xml` in the Strict WordprocessingML namespace. */
const withStrictStyles = (parts: Record<string, string>): Record<string, string> => withStrictPart(parts, "word/styles.xml");
/** The co-reach finding of `parts`, kept with its verdict but without a citation (#1016). */
const expectUncited = async (parts: Record<string, string>, path: string): Promise<void> => {
    const found = await diagnose(parts);
    expect(found).toHaveLength(1);
    expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
    expect(found[0]?.message).not.toContain("Word 16.115");
    expect(isWordBlockingIssue(found[0]!)).toBe(true);
};

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

    it.each(MISMATCH_STORY_CASES)(
        "keeps $via $segment declared $declared without its session in a package with a Strict part (#1016)",
        async ({ segment, via, declared }) => {
            for (const parts of strictVariants(storyContentTypeParts(segment, declared, via))) {
                const found = await diagnose(parts);
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({
                    code: ID,
                    severity: "error",
                    path: `word/${via === "glossary" ? "glossary/" : ""}${segment}9.xml`,
                });
                expect(found[0]?.message).toContain(`the ${via === "glossary" ? "glossary" : "main document"}'s ${segment} relationship`);
                expect(found[0]?.message).not.toContain("Word 16.115");
                expect(isWordBlockingIssue(found[0]!)).toBe(true);
            }
        },
    );

    it.each(["header", "footer"] as const)("ignores unread %s swaps", async (segment) => {
        expect(await diagnose(storyContentTypeParts(segment, undefined, "main", false))).toEqual([]);
        expect(await diagnose(storyContentTypeParts(segment, storyContentType(segment).toUpperCase()))).toEqual([]);
    });

    it.each(["image/png"] as const)("keeps an unobserved footer type %s silent", async (declared) => {
        expect(await diagnose(storyContentTypeParts("footer", declared))).toEqual([]);
    });

    it("rejects a read footer typed with the styles type, citing sessions 53 and 57 (#739, #778)", async () => {
        const found = await diagnose(storyContentTypeParts("footer", storyContentType("styles")));
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/footer9.xml" });
        expect(found[0]?.message).toContain("the main document's footer relationship");
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (sessions 53 and 57).");
        // The unread twin stays silent, as #516 observed for the header.
        expect(await diagnose(storyContentTypeParts("footer", storyContentType("styles"), "main", false))).toEqual([]);
    });

    // Sessions 53 and 57 stored Transitional main documents only (#861).
    it("keeps the read footer typed styles without a citation in a uniformly Strict package", async () => {
        const found = await diagnose(uniformlyStrict(storyContentTypeParts("footer", storyContentType("styles"))));
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/footer9.xml" });
        expect(found[0]?.message).toContain("the main document's footer relationship");
        expect(found[0]?.message).not.toContain("session");
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
        ["footer", storyContentType("styles"), false],
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
                // (#722), and every pair cites the session that showed it
                // (#779) — here the header type, session 53 (s53d–s53h) — so
                // those strict findings keep the Word sentence the swaps
                // already carry (#610); every other shape stays spec-only.
                const citation =
                    via === "main" && ["officeDocument", "styles", "settings", "numbering"].includes(segment) ? "session 53" : undefined;
                if (citation === undefined) {
                    expect(found[0]?.message).not.toContain("Word 16.115");
                } else {
                    expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (${citation}).`);
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

// Session 58's s58b rows were own-root parts, so #785's pairs co-reach only
// a part holding the story's own root (#803): a comments, endnotes or
// footnotes part whose root is `w:styles` keeps the verdict it had before
// #785 — silent under word-valid, the spec-path finding under strict —
// with no session-58 citation.
describe("session 58 pairs need the story's own root (#803)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /** The body's reference into each story. */
    const REFS = {
        comments: `<w:r><w:commentReference w:id="0"/></w:r>`,
        endnotes: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
        footnotes: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
    } as const;
    type Story = keyof typeof REFS;
    /** A used `word/<story>.xml` holding a `w:styles` root, typed with the styles content type. */
    const foreignRootStoryParts = (story: Story): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "</w:r></w:p>",
            `</w:r>${REFS[story]}</w:p>`,
        ),
        [`word/${story}.xml`]: MINIMAL_PARTS["word/styles.xml"]!,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdS" Type="${R}/${story}" Target="${story}.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/${story}.xml" ContentType="${storyContentType("styles")}"/></Types>`,
        ),
    });

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "reports no mismatch for a used %s part holding a w:styles root under word-valid",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "foreign-root.docx");
                writeFileSync(file, await buildDocx(foreignRootStoryParts(story)));
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(true);
                const result = await validate(file);
                expect(result.valid, "the reference dangles, so the full word-valid run fails").toBe(false);
                expect(result.issues.filter((issue) => issue.code === ID)).toEqual([]);
            });
        },
        60_000,
    );

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "cites the spec, not session 58, for a strict %s part holding a w:styles root",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "foreign-root.docx");
                writeFileSync(file, await buildDocx(foreignRootStoryParts(story)));
                const found = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
                expect(found).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        path: `word/${story}.xml`,
                        message: expect.stringContaining("ECMA-376 Part 1 (4th edition), clause 11.3"),
                    }),
                ]);
                expect(found[0]?.message).not.toContain("session");
            });
        },
        60_000,
    );
});

// Session 62 (s62c) refused the session-58 pairs' shape with a
// parameterized styles spelling: a used own-root `word/comments.xml`,
// `word/endnotes.xml` and `word/footnotes.xml` typed
// `…styles+xml; charset=utf-8` each got the recovery offer, while
// word-valid passed all three — three false negatives. Those pairs now
// also co-reach the styles media type with a parameter, citing session 62,
// while the exact spelling keeps session 58. A padded spelling with no
// parameter, and the header and footer pairs' inexact spellings, are
// unprobed and keep their verdict (#802).
describe("parameterized styles spelling on the session 58 pairs (#802, session 62)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const TEXT_P = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    /** Session 62's parameterized spelling of the styles type (s62c). */
    const STYLES_CHARSET = `${storyContentType("styles")}; charset=utf-8`;
    /** Each story's own root holding the item the body references, and that reference run. */
    const STORIES = {
        comments: {
            root: `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="a">${TEXT_P}</w:comment></w:comments>`,
            ref: `<w:r><w:commentReference w:id="0"/></w:r>`,
        },
        endnotes: {
            root: `<w:endnotes xmlns:w="${W}"><w:endnote w:id="1">${TEXT_P}</w:endnote></w:endnotes>`,
            ref: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
        },
        footnotes: {
            root: `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${TEXT_P}</w:footnote></w:footnotes>`,
            ref: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
        },
    } as const;
    type Story = keyof typeof STORIES;
    /** A `word/<story>.xml`, its own root holding the referenced item when `referenced`, related by the main document and typed `type`. */
    const storyParts = (story: Story, type: string, referenced = true): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "</w:r></w:p>",
            `</w:r>${referenced ? STORIES[story].ref : ""}</w:p>`,
        ),
        [`word/${story}.xml`]: STORIES[story].root,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdS" Type="${R}/${story}" Target="${story}.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
            "</Types>",
            `<Override PartName="/word/${story}.xml" ContentType="${type}"/></Types>`,
        ),
    });
    /** The same part holding a `w:styles` root instead, as #803's foreign-root specimens do. */
    const foreignRootStoryParts = (story: Story, type: string): Record<string, string> => ({
        ...storyParts(story, type),
        [`word/${story}.xml`]: MINIMAL_PARTS["word/styles.xml"]!,
    });

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "rejects a used %s part typed styles with a charset parameter, citing session 62",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "typed-styles-charset.docx");
                writeFileSync(file, await buildDocx(storyParts(story, STYLES_CHARSET)));
                const runs: [string, Awaited<ReturnType<typeof validate>>][] = [
                    ["the full word-valid run", await validate(file)],
                    ["the selected run", await validate(file, { checks: [ID] })],
                ];
                for (const [label, result] of runs) {
                    expect(result.valid, `${label} fails`).toBe(false);
                    const found = result.issues.filter((issue) => issue.code === ID && issue.severity === "error");
                    expect(found, label).toHaveLength(1);
                    expect(found[0]).toMatchObject({ path: `word/${story}.xml` });
                    expect(found[0]?.message).toContain(`the main document's ${story} relationship`);
                    expect(found[0]?.message).toContain(`is typed "${STYLES_CHARSET}" by an Override`);
                    expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 62).");
                    expect(isWordBlockingIssue(found[0]!)).toBe(true);
                }
                // The inexact check reports nothing: the part's reaches expect
                // its own type, so the parameterized spelling is the mismatch's
                // finding alone and nothing is reported twice.
                expect(runs[0]![1].issues.filter((issue) => issue.code === "word-part-content-type-inexact")).toEqual([]);
            });
        },
        60_000,
    );

    // #1016: sessions 58 and 62 stored Transitional packages, so a Strict
    // part beside the story drops the citation.
    it.each(["comments", "endnotes", "footnotes"] as const)(
        "keeps a used %s part typed styles, exactly or with a charset, without its session beside a Strict part (#1016)",
        async (story) => {
            for (const type of [storyContentType("styles"), STYLES_CHARSET]) {
                await expectUncited(withStrictApp(storyParts(story, type)), `word/${story}.xml`);
                await expectUncited(withStrictStyles(storyParts(story, type)), `word/${story}.xml`);
            }
        },
    );

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "still cites session 58 for a used %s part typed with the exact styles spelling",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "typed-styles.docx");
                writeFileSync(file, await buildDocx(storyParts(story, storyContentType("styles"))));
                const result = await validate(file);
                expect(result.valid, "the full word-valid run fails").toBe(false);
                const found = result.issues.filter((issue) => issue.code === ID && issue.severity === "error");
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({ path: `word/${story}.xml` });
                expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 58).");
            });
        },
        60_000,
    );

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "reports no mismatch for a used %s part holding a w:styles root typed with the charset parameter",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "foreign-root-charset.docx");
                writeFileSync(file, await buildDocx(foreignRootStoryParts(story, STYLES_CHARSET)));
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(true);
                const result = await validate(file);
                expect(result.valid, "the reference dangles, so the full word-valid run fails").toBe(false);
                expect(result.issues.filter((issue) => issue.code === ID)).toEqual([]);
            });
        },
        60_000,
    );

    it("keeps a used own-root comments part typed with a padded styles spelling and no parameter without a mismatch", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "padded-styles.docx");
            writeFileSync(file, await buildDocx(storyParts("comments", `${storyContentType("styles")} `)));
            expect((await validate(file, { checks: [ID] })).valid).toBe(true);
            expect((await validate(file)).issues.filter((issue) => issue.code === ID)).toEqual([]);
        });
    }, 60_000);

    it("keeps an own-root comments part typed with the charset parameter that nothing references without a mismatch", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unreferenced-charset.docx");
            writeFileSync(file, await buildDocx(storyParts("comments", STYLES_CHARSET, false)));
            expect((await validate(file, { checks: [ID] })).valid).toBe(true);
            expect((await validate(file)).issues.filter((issue) => issue.code === ID)).toEqual([]);
        });
    }, 60_000);
});

// A stored `mc:AlternateContent` root unwraps to the selected branch's
// element in the effective view, so a selected `w14` Choice holding the
// story's own root passes #803's guard, while session 58's rows (s58b) and
// session 62's parameterized spelling (s62c) stored the story's own root.
// Session 63 (s63b, two runs, all readings agree) offered recovery for the
// comments part in that shape typed with the styles type exactly, and
// opened the same part typed comments clean, so that finding cites session
// 63. The endnotes and footnotes parts in it, the charset spelling, and the
// other stored roots that unwrap to the comments root (an ignorable wrapper
// that processes its content, alone or inside the Choice, and a selected
// Fallback holding the root) are unprobed and
// keep the verdict they had before #785, as the foreign root does (#803):
// silent under word-valid, the spec-path finding under strict, and no
// session citation (#818).
describe("a story part whose stored root is an AlternateContent (#818, session 63 s63b)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    const TEXT_P = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    /** Session 62's parameterized spelling of the styles type (s62c). */
    const STYLES_CHARSET = `${storyContentType("styles")}; charset=utf-8`;
    /** Each story's own root holding the item `id`, and the body's reference to the held one. */
    const STORIES = {
        comments: {
            root: (id: string) => `<w:comments xmlns:w="${W}"><w:comment w:id="${id}" w:author="a">${TEXT_P}</w:comment></w:comments>`,
            ref: `<w:r><w:commentReference w:id="0"/></w:r>`,
            held: "0",
        },
        endnotes: {
            root: (id: string) => `<w:endnotes xmlns:w="${W}"><w:endnote w:id="${id}">${TEXT_P}</w:endnote></w:endnotes>`,
            ref: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
            held: "1",
        },
        footnotes: {
            root: (id: string) => `<w:footnotes xmlns:w="${W}"><w:footnote w:id="${id}">${TEXT_P}</w:footnote></w:footnotes>`,
            ref: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
            held: "1",
        },
    } as const;
    type Story = keyof typeof STORIES;
    /** An `mc:AlternateContent` holding `content` in a selected `w14` Choice, with an empty Fallback. */
    const alternateContent = (content: string): string =>
        `<mc:AlternateContent xmlns:mc="${MC}" xmlns:w="${W}" xmlns:w14="${W14}">` +
        `<mc:Choice Requires="w14">${content}</mc:Choice><mc:Fallback/>` +
        `</mc:AlternateContent>`;
    /** An ignorable element that names itself in `mc:ProcessContent`, so the effective view keeps `content` without it. */
    const ignorableWrapper = (content: string): string =>
        `<x:wrap xmlns:x="urn:example:wrap" xmlns:mc="${MC}" mc:Ignorable="x" mc:ProcessContent="x:wrap">${content}</x:wrap>`;
    /** A used `word/<story>.xml` stored as `stored`, the body referencing the held item, typed `type`. */
    const usedStoryParts = (story: Story, type: string, stored: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "</w:r></w:p>",
            `</w:r>${STORIES[story].ref}</w:p>`,
        ),
        [`word/${story}.xml`]: stored,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdS" Type="${R}/${story}" Target="${story}.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
            "</Types>",
            `<Override PartName="/word/${story}.xml" ContentType="${type}"/></Types>`,
        ),
    });
    /**
     * A used `word/<story>.xml` whose stored root is `mc:AlternateContent`, a
     * selected `w14` Choice holding the story's own root with the referenced
     * item and an empty Fallback, typed `type` — session 63's s63b specimen
     * for the comments part.
     */
    const alternateRootStoryParts = (story: Story, type: string): Record<string, string> =>
        usedStoryParts(story, type, alternateContent(STORIES[story].root(STORIES[story].held)));

    it("rejects the used comments part typed styles, citing session 63 (s63b)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "alternate-root.docx");
            writeFileSync(file, await buildDocx(alternateRootStoryParts("comments", storyContentType("styles"))));
            const expected = [
                expect.objectContaining({
                    severity: "error",
                    path: "word/comments.xml",
                    message: expect.stringContaining("Word 16.115 offers to recover the package (session 63)."),
                }),
            ];
            const partial = await validate(file, { checks: [ID] });
            expect(partial.valid).toBe(false);
            expect(partial.issues.filter((issue) => issue.code === ID)).toEqual(expected);
            const result = await validate(file);
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.code === ID)).toEqual(expected);
            const strict = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
            expect(strict).toEqual([
                expect.objectContaining({
                    severity: "error",
                    path: "word/comments.xml",
                    message: expect.stringContaining(
                        "clause 11.3 assigns the comments part. Word 16.115 offers to recover the package (session 63).",
                    ),
                }),
            ]);
        });
    }, 60_000);

    it("keeps session 63 for XML whitespace, comments and processing instructions around the roots", async () => {
        const padding = " \t\r\n<!-- comment --><?probe content?>";
        const stored = alternateContent(`${padding}${STORIES.comments.root("0")}${padding}`)
            .replace("<mc:Choice", `${padding}<mc:Choice`)
            .replace("<mc:Fallback/>", `${padding}<mc:Fallback>${padding}</mc:Fallback>${padding}`);
        const found = await diagnose(usedStoryParts("comments", storyContentType("styles"), stored));
        expect(found).toEqual([
            expect.objectContaining({
                severity: "error",
                path: "word/comments.xml",
                message: expect.stringContaining("Word 16.115 offers to recover the package (session 63)."),
            }),
        ]);
    });

    // Session 65 (s65b–s65d) refused these styles-typed shapes and opened
    // each matching comments-typed calibration clean (#847).
    const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const SESSION_65_ROOTS = [
        ["text before the root in the Choice", alternateContent(`NOT-A-ROOT${STORIES.comments.root("0")}`)],
        [
            "text in the Fallback",
            alternateContent(STORIES.comments.root("0")).replace("<mc:Fallback/>", "<mc:Fallback>NOT-EMPTY</mc:Fallback>"),
        ],
        ["a Choice requiring w", alternateContent(STORIES.comments.root("0")).replace('Requires="w14"', 'Requires="w"')],
        ["a Strict root in the Choice", alternateContent(STORIES.comments.root("0").replace(W, STRICT_W))],
        ["a Strict root stored directly", STORIES.comments.root("0").replace(W, STRICT_W)],
    ] as const;
    const RUNS = [
        ["full word-valid", {}],
        ["selected word-valid", { checks: [ID] }],
        ["strict", { profile: "strict" }],
    ] as const;
    describe.each(RUNS)("session 65 citations in %s", (_mode, options) => {
        it.each(SESSION_65_ROOTS)(
            "cites session 65 for %s",
            async (_shape, stored) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, "session-65.docx");
                    writeFileSync(file, await buildDocx(usedStoryParts("comments", storyContentType("styles"), stored)));
                    const result = await validate(file, options);
                    expect(result.valid).toBe(false);
                    expect(result.issues.filter((issue) => issue.code === ID)).toEqual([
                        expect.objectContaining({
                            severity: "error",
                            path: "word/comments.xml",
                            message: expect.stringContaining("Word 16.115 offers to recover the package (session 65)."),
                        }),
                    ]);
                });
            },
            60_000,
        );
    });
    // #1016: sessions 63 and 65 stored Transitional parts outside the
    // comments part; s65c and s65d stored a Strict comments root, in the
    // Choice or directly, which is the shape itself.
    it.each([["the s63b root", alternateContent(STORIES.comments.root("0"))] as const, ...SESSION_65_ROOTS])(
        "keeps %s typed styles without its session beside a Strict part (#1016)",
        async (_shape, stored) => {
            const parts = usedStoryParts("comments", storyContentType("styles"), stored);
            await expectUncited(withStrictApp(parts), "word/comments.xml");
            await expectUncited(withStrictStyles(parts), "word/comments.xml");
        },
    );
    it.each(SESSION_65_ROOTS.filter(([shape]) => shape.startsWith("a Strict root")))(
        "keeps session 65 for %s, the Strict part its own shape names (#1016)",
        async (_shape, stored) => {
            const found = await diagnose(usedStoryParts("comments", storyContentType("styles"), stored));
            expect(found).toHaveLength(1);
            expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 65).");
        },
    );

    it.each(SESSION_65_ROOTS)(
        "keeps the comments-typed calibration for %s without a mismatch",
        async (_shape, stored) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "session-65-calibration.docx");
                writeFileSync(file, await buildDocx(usedStoryParts("comments", storyContentType("comments"), stored)));
                for (const [mode, options] of RUNS) {
                    const result = await validate(file, options);
                    expect(
                        result.issues.filter((issue) => issue.code === ID),
                        mode,
                    ).toEqual([]);
                    if (mode !== "strict") expect(result.valid, mode).toBe(true);
                }
            });
        },
        60_000,
    );

    it("cites no content-type session for text between the Choice and the Fallback (#851)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "between-text.docx");
            const stored = alternateContent(STORIES.comments.root("0")).replace("</mc:Choice>", "</mc:Choice>NOT-A-BRANCH");
            writeFileSync(file, await buildDocx(usedStoryParts("comments", storyContentType("styles"), stored)));
            for (const [mode, options] of RUNS) {
                const found = (await validate(file, options)).issues.filter((issue) => issue.code === ID);
                if (mode === "strict") {
                    expect(found).toEqual([
                        expect.objectContaining({
                            severity: "error",
                            path: "word/comments.xml",
                            message: expect.stringContaining("ECMA-376 Part 1 (4th edition), clause 11.3"),
                        }),
                    ]);
                    expect(found[0]?.message).not.toContain("session");
                } else {
                    expect(found, mode).toEqual([]);
                }
            }
        });
    }, 60_000);

    it.each(["endnotes", "footnotes"] as const)(
        "reports no mismatch for a used %s part whose stored root is an AlternateContent under word-valid",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "alternate-root.docx");
                writeFileSync(file, await buildDocx(alternateRootStoryParts(story, storyContentType("styles"))));
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(true);
                expect(partial.issues.filter((issue) => issue.code === ID)).toEqual([]);
                const result = await validate(file);
                expect(result.valid, "the selected Choice holds the referenced item, so only warnings remain").toBe(true);
                expect(result.issues.filter((issue) => issue.code === ID)).toEqual([]);
            });
        },
        60_000,
    );

    it.each(["endnotes", "footnotes"] as const)(
        "cites the spec, not session 58, for a strict %s part whose stored root is an AlternateContent",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "alternate-root.docx");
                writeFileSync(file, await buildDocx(alternateRootStoryParts(story, storyContentType("styles"))));
                const found = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
                expect(found).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        path: `word/${story}.xml`,
                        message: expect.stringContaining("ECMA-376 Part 1 (4th edition), clause 11.3"),
                    }),
                ]);
                expect(found[0]?.message).not.toContain("session");
            });
        },
        60_000,
    );

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "reports no mismatch for a used %s part whose stored root is an AlternateContent typed with the charset parameter",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "alternate-root-charset.docx");
                writeFileSync(file, await buildDocx(alternateRootStoryParts(story, STYLES_CHARSET)));
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(true);
                expect(partial.issues.filter((issue) => issue.code === ID)).toEqual([]);
                expect((await validate(file)).issues.filter((issue) => issue.code === ID)).toEqual([]);
                const strict = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
                expect(strict).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        path: `word/${story}.xml`,
                        message: expect.stringContaining("ECMA-376 Part 1 (4th edition), clause 11.3"),
                    }),
                ]);
                expect(strict[0]?.message).not.toContain("session");
            });
        },
        60_000,
    );

    // The rescope preserves verdicts for shapes no session stored. They
    // retain the existing finding, without borrowing an observation (#847).
    it.each([
        ["a Strict root with the parameterized styles spelling", STORIES.comments.root("0").replace(W, STRICT_W), STYLES_CHARSET],
        [
            "an AlternateContent whose Choice requires the relationships namespace",
            alternateContent(STORIES.comments.root("0"))
                .replace("xmlns:w14=", `xmlns:r="${R}" xmlns:w14=`)
                .replace('Requires="w14"', 'Requires="r"'),
        ],
        ["an AlternateContent whose Choice has text after the root", alternateContent(`${STORIES.comments.root("0")}UNPROBED`)],
        [
            "an AlternateContent whose Choice and Fallback both hold text",
            alternateContent(`NOT-A-ROOT${STORIES.comments.root("0")}`).replace("<mc:Fallback/>", "<mc:Fallback>NOT-EMPTY</mc:Fallback>"),
        ],
        // No session stored CDATA; session 65 stored ordinary text (#857).
        [
            "an AlternateContent with CDATA between the Choice and the Fallback",
            alternateContent(STORIES.comments.root("0")).replace("</mc:Choice>", "</mc:Choice><![CDATA[NOT-A-BRANCH]]>"),
        ],
        [
            "an AlternateContent whose Choice has CDATA before the root",
            alternateContent(`<![CDATA[NOT-A-ROOT]]>${STORIES.comments.root("0")}`),
        ],
        [
            "an AlternateContent whose Fallback holds CDATA",
            alternateContent(STORIES.comments.root("0")).replace("<mc:Fallback/>", "<mc:Fallback><![CDATA[NOT-EMPTY]]></mc:Fallback>"),
        ],
    ])(
        "keeps the verdict without a citation for %s",
        async (_shape, stored, type = storyContentType("styles")) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "unprobed-root.docx");
                writeFileSync(file, await buildDocx(usedStoryParts("comments", type, stored)));
                for (const [mode, options] of RUNS) {
                    const result = await validate(file, options);
                    expect(result.valid, mode).toBe(false);
                    const found = result.issues.filter((issue) => issue.code === ID);
                    expect(found, mode).toEqual([expect.objectContaining({ severity: "error", path: "word/comments.xml" })]);
                    expect(found[0]?.message, mode).not.toContain("session");
                    if (mode === "strict") expect(found[0]?.message).toContain("ECMA-376 Part 1 (4th edition), clause 11.3");
                }
            });
        },
        60_000,
    );

    // Session 65 stored its Strict comments roots in a Transitional package
    // (s65c, s65d); a uniformly Strict package is unprobed (#857).
    const UNIFORMLY_STRICT_ROOTS = [
        ["stored directly", STORIES.comments.root("0").replace(W, STRICT_W)],
        ["in the Choice", alternateContent(STORIES.comments.root("0").replace(W, STRICT_W))],
    ] as const;
    it.each(UNIFORMLY_STRICT_ROOTS)(
        "keeps the verdict without a citation for a Strict root %s in a uniformly Strict package",
        async (_shape, stored) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "uniformly-strict.docx");
                writeFileSync(file, await buildDocx(uniformlyStrict(usedStoryParts("comments", storyContentType("styles"), stored))));
                for (const [mode, options] of RUNS) {
                    const result = await validate(file, options);
                    expect(result.valid, mode).toBe(false);
                    const found = result.issues.filter((issue) => issue.code === ID);
                    expect(found, mode).toEqual([expect.objectContaining({ severity: "error", path: "word/comments.xml" })]);
                    expect(found[0]?.message, mode).not.toContain("session");
                }
            });
        },
        60_000,
    );
    it.each(UNIFORMLY_STRICT_ROOTS)(
        "keeps the comments-typed calibration for a Strict root %s in a uniformly Strict package without a mismatch",
        async (_shape, stored) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "uniformly-strict-calibration.docx");
                writeFileSync(file, await buildDocx(uniformlyStrict(usedStoryParts("comments", storyContentType("comments"), stored))));
                for (const [mode, options] of RUNS) {
                    const result = await validate(file, options);
                    expect(
                        result.issues.filter((issue) => issue.code === ID),
                        mode,
                    ).toEqual([]);
                    if (mode !== "strict") expect(result.valid, mode).toBe(true);
                }
            });
        },
        60_000,
    );

    // Sessions 58, 62, 63 and 65 stored every shape they cite in a package
    // whose main document is Transitional, so a uniformly Strict package
    // keeps the verdict without borrowing their citations (#861).
    it.each([
        ["s63b's AlternateContent root", "comments", alternateContent(STORIES.comments.root("0")), storyContentType("styles")],
        ["text before the root in the Choice", "comments", SESSION_65_ROOTS[0][1], storyContentType("styles")],
        ["text in the Fallback", "comments", SESSION_65_ROOTS[1][1], storyContentType("styles")],
        ["a Choice requiring w", "comments", SESSION_65_ROOTS[2][1], storyContentType("styles")],
        ["a Transitional comments root stored directly", "comments", STORIES.comments.root("0"), storyContentType("styles")],
        ["a Transitional footnotes root stored directly", "footnotes", STORIES.footnotes.root("1"), storyContentType("styles")],
        ["a Transitional comments root typed with the parameterized spelling", "comments", STORIES.comments.root("0"), STYLES_CHARSET],
    ] as const)(
        "keeps the verdict without a citation for %s in a uniformly Strict package",
        async (_shape, story, stored, type) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "uniformly-strict-transitional-root.docx");
                writeFileSync(file, await buildDocx(uniformlyStrict(usedStoryParts(story, type, stored))));
                for (const [mode, options] of RUNS) {
                    const result = await validate(file, options);
                    expect(result.valid, mode).toBe(false);
                    const found = result.issues.filter((issue) => issue.code === ID);
                    expect(found, mode).toEqual([expect.objectContaining({ severity: "error", path: `word/${story}.xml` })]);
                    expect(found[0]?.message, mode).not.toContain("session");
                }
            });
        },
        60_000,
    );

    // The effective view also unwraps an ignorable wrapper that processes
    // its content, a selected Fallback, and a Choice with no Fallback or
    // beside a second Choice, so the effective root alone does not show
    // session 63's stored shape (#831, #837, #843).
    it.each([
        ["an ignorable wrapper", ignorableWrapper(STORIES.comments.root("0"))],
        ["an AlternateContent whose Choice holds an ignorable wrapper", alternateContent(ignorableWrapper(STORIES.comments.root("0")))],
        [
            "an AlternateContent whose unselected Choice is empty and whose Fallback holds the root",
            `<mc:AlternateContent xmlns:mc="${MC}" xmlns:w="${W}" xmlns:u="urn:example:unknown">` +
                `<mc:Choice Requires="u"/><mc:Fallback>${STORIES.comments.root("0")}</mc:Fallback>` +
                `</mc:AlternateContent>`,
        ],
        [
            "an AlternateContent whose Choice holds the root and which has no Fallback",
            `<mc:AlternateContent xmlns:mc="${MC}" xmlns:w="${W}" xmlns:w14="${W14}">` +
                `<mc:Choice Requires="w14">${STORIES.comments.root("0")}</mc:Choice>` +
                `</mc:AlternateContent>`,
        ],
        [
            "an AlternateContent whose two Choices each hold the root",
            `<mc:AlternateContent xmlns:mc="${MC}" xmlns:w="${W}" xmlns:w14="${W14}">` +
                `<mc:Choice Requires="w14">${STORIES.comments.root("0")}</mc:Choice>` +
                `<mc:Choice Requires="w14">${STORIES.comments.root("0")}</mc:Choice><mc:Fallback/>` +
                `</mc:AlternateContent>`,
        ],
    ])(
        "cites no session for a used comments part typed styles whose stored root is %s",
        async (_shape, stored) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "wrapped-root.docx");
                writeFileSync(file, await buildDocx(usedStoryParts("comments", storyContentType("styles"), stored)));
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(true);
                expect(partial.issues.filter((issue) => issue.code === ID)).toEqual([]);
                expect((await validate(file)).issues.filter((issue) => issue.code === ID)).toEqual([]);
                const strict = (await validate(file, { profile: "strict" })).issues.filter((issue) => issue.code === ID);
                expect(strict).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        path: "word/comments.xml",
                        message: expect.stringContaining("ECMA-376 Part 1 (4th edition), clause 11.3"),
                    }),
                ]);
                expect(strict[0]?.message).not.toContain("session");
            });
        },
        60_000,
    );
});

// Sessions 49 and 50 (2026-10-26): Word 16.115 offered recovery for the main
// styles part typed with the settings type (s49b), and for a related settings
// part typed with the styles type, a related numbering part typed with the
// settings type and the main document typed with the styles type (s50d);
// session 53 (s53d–s53h), session 56 (s56a), session 57 (s57a–s57c) and
// session 59 (s59a) showed the other pairs on those reaches, each citation
// naming its own session (#779, #789).
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
            citation: "session 49",
        },
        {
            label: "a related settings part typed with the styles type",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("styles")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            citation: "session 50",
        },
        {
            label: "a related numbering part typed with the settings type",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("settings")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            citation: "session 50",
        },
        {
            label: "the main document typed with the styles type",
            parts: retypedParts("word/document.xml", storyContentType("document.main"), storyContentType("styles")),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            citation: "session 50",
        },
    ] as const;

    it.each(CASES)("rejects $label under word-valid with its sessions", async ({ parts, path, by, citation }) => {
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
        expect(found[0]?.message).toContain(by);
        expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (${citation}).`);
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
        expect(await diagnose({ ...MINIMAL_PARTS })).toEqual([]);
    });

    it.each(CASES)("keeps $label without its session in a package with a Strict part (#1016)", async ({ parts, path, by }) => {
        for (const variant of strictVariants(parts)) {
            const found = await diagnose(variant);
            expect(found).toHaveLength(1);
            expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
            expect(found[0]?.message).toContain(by);
            expect(found[0]?.message).not.toContain("Word 16.115");
            expect(isWordBlockingIssue(found[0]!)).toBe(true);
        }
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

// Session 53 (2026-10-26, chunks s53a and s53b; Word 16.115, two runs, all
// readings agree) refused every co-reach shape: an eagerly read part — the
// main document, a related numbering part or the styles part — that a read
// footer or a second eager relationship of another kind also reaches,
// whatever the Override says. The calibrations carry the part's own type.
describe("observed co-reach (#739, session 53)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MAIN_TYPE = storyContentType("document.main");
    const retypedParts = (part: string, own: string, type: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            `<Override PartName="/${part}" ContentType="${own}"/>`,
            `<Override PartName="/${part}" ContentType="${type}"/>`,
        ),
    });
    /** `parts` plus a footer relationship (rIdF) to `target`, referenced by the body's section. */
    const withReadFooter = (parts: Record<string, string>, target: string): Record<string, string> => ({
        ...parts,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "<w:sectPr>",
            `<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/>`,
        ),
        "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdF" Type="${R}/footer" Target="${target}"/></Relationships>`,
        ),
    });
    /** The styles part also the target of the main document's settings relationship (review S5). */
    const stylesSettingsParts = (type: string): Record<string, string> => ({
        ...retypedParts("word/styles.xml", storyContentType("styles"), type),
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rId2" Type="${R}/settings" Target="styles.xml"/></Relationships>`,
        ),
    });
    const CASES = [
        {
            label: "a read footer targeting the main document, typed main (calibration)",
            parts: withReadFooter({ ...MINIMAL_PARTS }, "document.xml"),
            path: "word/document.xml",
            by: "the main document's footer relationship",
            citation: "sessions 53 and 57",
        },
        {
            label: "a read footer targeting the main document typed with the styles type",
            parts: withReadFooter(retypedParts("word/document.xml", MAIN_TYPE, storyContentType("styles")), "document.xml"),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            citation: "session 50",
        },
        {
            label: "a read footer targeting the related numbering part, typed numbering (calibration)",
            parts: withReadFooter(
                relatedContentTypeParts(RELATED_PART_SPECIMENS[2], RELATED_PART_SPECIMENS[2].contentType),
                "numbering.xml",
            ),
            path: "word/numbering.xml",
            by: "the main document's footer relationship",
            citation: "sessions 53 and 57",
        },
        {
            label: "a read footer targeting the related numbering part typed with the settings type",
            parts: withReadFooter(relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("settings")), "numbering.xml"),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            citation: "session 50",
        },
        {
            label: "a read footer targeting the styles part, typed styles (calibration)",
            parts: withReadFooter({ ...MINIMAL_PARTS }, "styles.xml"),
            path: "word/styles.xml",
            by: "the main document's footer relationship",
            citation: "sessions 53 and 57",
        },
        {
            label: "a read footer targeting the styles part typed with the settings type",
            parts: withReadFooter(retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("settings")), "styles.xml"),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            citation: "session 49",
        },
        {
            label: "the styles part also the settings target, typed styles (calibration)",
            parts: stylesSettingsParts(storyContentType("styles")),
            path: "word/styles.xml",
            by: "the main document's settings relationship",
            citation: "session 50",
        },
        {
            label: "the styles part also the settings target, typed settings",
            parts: stylesSettingsParts(storyContentType("settings")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            citation: "session 49",
        },
    ] as const;

    it.each(CASES)("rejects $label under word-valid with its session", async ({ parts, path, by, citation }) => {
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
        expect(found[0]?.message).toContain(by);
        expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (${citation}).`);
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
    });

    // #1016: session 53 stored Transitional packages.
    it.each(CASES)("keeps $label without its session beside a Strict docProps/app.xml (#1016)", async ({ parts, path }) => {
        await expectUncited(withStrictApp(parts), path);
    });

    // The footer co-reach names no shape part, so a Strict root on the judged
    // part itself drops the citation too (#1030).
    it.each(CASES)("keeps $label without its session when its own root is Strict (#1030)", async ({ parts, path }) => {
        await expectUncited(withStrictPart(parts, path), path);
    });

    // Session 56 (s56a): a footer-typed styles part that a second footer
    // relationship also targets gets the recovery offer whether no section
    // references that relationship or the body does; the matching footer
    // reach no longer exempts the styles reach.
    it.each([
        { label: "no section references", read: false },
        { label: "the body references", read: true },
    ])("rejects a footer-typed styles part a footer relationship $label also targets (session 56)", async ({ read }) => {
        const typed = retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("footer"));
        const parts = read
            ? withReadFooter(typed, "styles.xml")
            : {
                  ...typed,
                  "word/_rels/document.xml.rels": typed["word/_rels/document.xml.rels"]!.replace(
                      "</Relationships>",
                      `<Relationship Id="rIdF" Type="${R}/footer" Target="styles.xml"/></Relationships>`,
                  ),
              };
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/styles.xml" });
        expect(found[0]?.message).toContain("the main document's styles relationship");
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "styles-typed-footer.docx");
            writeFileSync(file, await buildDocx(parts));
            expect((await validate(file)).valid).toBe(false);
            expect((await validate(file, { checks: [ID] })).valid).toBe(false);
        });
    });

    it("invalidates the full run and a named run on every row, keeping the controls' verdicts", async () => {
        await withTempDir(async (dir) => {
            for (const { label, parts } of CASES) {
                const file = path.join(dir, "coreach.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file);
                expect(result.valid, label).toBe(false);
                expect(result.issues.filter((issue) => issue.code === ID && issue.severity === "error")).toHaveLength(1);
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid, label).toBe(false);
                expect(partial.issues.filter((issue) => issue.code === ID)).toHaveLength(1);
            }
            // The controls: the same rows without the second relationship
            // keep their current verdicts.
            const controls: [string, Record<string, string>, boolean][] = [
                ["the minimal package", { ...MINIMAL_PARTS }, true],
                ["the styles part typed styles", { ...MINIMAL_PARTS }, true],
                [
                    "the main document typed with the styles type",
                    retypedParts("word/document.xml", MAIN_TYPE, storyContentType("styles")),
                    false,
                ],
                [
                    "a related numbering part typed with the settings type",
                    relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("settings")),
                    false,
                ],
                [
                    "the styles part typed with the settings type",
                    retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("settings")),
                    false,
                ],
            ];
            for (const [label, parts, valid] of controls) {
                const file = path.join(dir, "control.docx");
                writeFileSync(file, await buildDocx(parts));
                expect((await validate(file)).valid, label).toBe(valid);
            }
        });
    }, 60_000);
});

// Session 57 (2026-10-26; Word 16.115, two runs, all readings agree) showed
// the thirteen mapped-type pairs no earlier session had, and session 59
// (2026-10-26, s59a) the remaining three — the styles part typed endnotes
// and a related numbering part typed footnotes or endnotes — so every
// recovery sentence names the session that showed its pair — the reach plus
// the declared type — instead of its reach's sessions: 49 the settings type
// on the styles part (s49b), 50 the styles type on the main document and a
// related settings part and the settings type on a related numbering part
// (s50d), 53 the other mapped types it ran (s53d–s53h), 56 the footer type
// on the styles part (s56a), 57 the thirteen pairs of s57a–s57c and 59 the
// remaining three (s59a).
describe("per-pair recovery citations (#779)", () => {
    const THEME_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
    const retypedParts = (part: string, own: string, type: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            `<Override PartName="/${part}" ContentType="${own}"/>`,
            `<Override PartName="/${part}" ContentType="${type}"/>`,
        ),
    });
    const CASES = [
        {
            label: "the main document typed webSettings (s57a)",
            parts: retypedParts("word/document.xml", storyContentType("document.main"), storyContentType("webSettings")),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            session: "session 57",
        },
        {
            label: "a related settings part typed webSettings (s57a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("webSettings")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 57",
        },
        {
            label: "a related settings part typed glossary (s57a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("document.glossary")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 57",
        },
        {
            label: "a related settings part typed footer (s57a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("footer")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 57",
        },
        {
            label: "a related settings part typed footnotes (s57a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("footnotes")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 57",
        },
        {
            label: "a related settings part typed endnotes (s57a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("endnotes")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 57",
        },
        {
            label: "a related numbering part typed webSettings (s57b)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("webSettings")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 57",
        },
        {
            label: "a related numbering part typed glossary (s57b)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("document.glossary")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 57",
        },
        {
            label: "a related numbering part typed theme (s57b)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], THEME_TYPE),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 57",
        },
        {
            label: "a related numbering part typed footer (s57b)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("footer")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 57",
        },
        {
            label: "the styles part typed webSettings (s57b)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("webSettings")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 57",
        },
        {
            label: "the styles part typed glossary (s57b)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("document.glossary")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 57",
        },
        {
            label: "the styles part typed footnotes (s57c)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("footnotes")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 57",
        },
        {
            label: "the styles part typed footer (s56a)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("footer")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 56",
        },
        {
            label: "the styles part typed settings (s49b)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("settings")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 49",
        },
        {
            label: "the main document typed styles (s50d)",
            parts: retypedParts("word/document.xml", storyContentType("document.main"), storyContentType("styles")),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            session: "session 50",
        },
        {
            label: "a related settings part typed styles (s50d)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("styles")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 50",
        },
        {
            label: "a related numbering part typed settings (s50d)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("settings")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 50",
        },
        {
            label: "the main document typed theme (s53d)",
            parts: retypedParts("word/document.xml", storyContentType("document.main"), THEME_TYPE),
            path: "word/document.xml",
            by: "the package officeDocument relationship",
            session: "session 53",
        },
        {
            label: "a related settings part typed main (s53f)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[0], storyContentType("document.main")),
            path: "word/settings.xml",
            by: "the main document's settings relationship",
            session: "session 53",
        },
        {
            label: "a related numbering part typed main (s53g)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("document.main")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 53",
        },
        {
            label: "the styles part typed main (s53g)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("document.main")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 53",
        },
        {
            label: "the styles part typed endnotes (s59a)",
            parts: retypedParts("word/styles.xml", storyContentType("styles"), storyContentType("endnotes")),
            path: "word/styles.xml",
            by: "the main document's styles relationship",
            session: "session 59",
        },
        {
            label: "a related numbering part typed footnotes (s59a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("footnotes")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 59",
        },
        {
            label: "a related numbering part typed endnotes (s59a)",
            parts: relatedContentTypeParts(RELATED_PART_SPECIMENS[2], storyContentType("endnotes")),
            path: "word/numbering.xml",
            by: "the main document's numbering relationship",
            session: "session 59",
        },
    ] as const;

    it.each(CASES)("cites $session for $label", async ({ parts, path, by, session }) => {
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
        expect(found[0]?.message).toContain(by);
        expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (${session}).`);
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
    });

    it.each(CASES)("keeps $label without its session in a package with a Strict part (#1016)", async ({ parts, path, by }) => {
        for (const variant of strictVariants(parts)) {
            const found = await diagnose(variant);
            expect(found).toHaveLength(1);
            expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
            expect(found[0]?.message).toContain(by);
            expect(found[0]?.message).not.toContain("Word 16.115");
            expect(isWordBlockingIssue(found[0]!)).toBe(true);
        }
    });

    it("invalidates the full word-valid run on the session 57 pairs with exactly one error", async () => {
        await withTempDir(async (dir) => {
            for (const { label, parts } of CASES.filter((entry) => entry.session === "session 57")) {
                const file = path.join(dir, "pair.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file);
                expect(result.valid, label).toBe(false);
                expect(result.issues.filter((issue) => issue.code === ID && issue.severity === "error")).toHaveLength(1);
            }
        });
    }, 120_000);
});

// Session 57 (2026-10-26, chunks s57c–s57e; Word 16.115, two runs, all
// readings agree) asked Word about the header/footer reaches with no eager
// reach: every row gets the recovery offer. The false negatives — a footer
// the body references typed with the settings type, a header the body
// references typed with the main document or styles type, and the
// exact-typed styles part a read header relationship also targets (the
// s57e calibration) — join the footer rows session 53 already showed; a
// header typed with a type no session showed stays valid.
describe("read header/footer co-reach (#778, session 57)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /** `word/styles.xml`, exact styles type, also the target of a header relationship the body references (s57e). */
    const stylesWithReadHeader = (): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "<w:sectPr>",
            `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`,
        ),
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdH" Type="${R}/header" Target="styles.xml"/></Relationships>`,
        ),
    });
    /** A reached glossary's own styles part that a footer relationship the main body references also targets (s57d). */
    const glossaryStylesWithReadFooter = (): Record<string, string> => {
        const parts = glossaryStylesParts(true);
        return {
            ...parts,
            "word/document.xml": parts["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
                "<w:sectPr>",
                `<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/>`,
            ),
            "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdF" Type="${R}/footer" Target="glossary/styles.xml"/></Relationships>`,
            ),
        };
    };
    const FALSE_NEGATIVES = [
        {
            label: "footer1-referenced-typed-settings (s57c)",
            parts: storyContentTypeParts("footer", storyContentType("settings")),
            path: "word/footer9.xml",
            by: "the main document's footer relationship",
        },
        {
            label: "header1-referenced-typed-main (s57c)",
            parts: storyContentTypeParts("header", storyContentType("document.main")),
            path: "word/header9.xml",
            by: "the main document's header relationship",
        },
        {
            label: "header1-referenced-typed-styles (s57d)",
            parts: storyContentTypeParts("header", storyContentType("styles")),
            path: "word/header9.xml",
            by: "the main document's header relationship",
        },
        {
            label: "styles-exact-read-header-calibration (s57e)",
            parts: stylesWithReadHeader(),
            path: "word/styles.xml",
            by: "the main document's header relationship",
        },
    ] as const;

    // #1016: sessions 53 and 57 stored Transitional packages.
    it.each([
        ...FALSE_NEGATIVES,
        {
            label: "footer1-referenced-typed-styles (s57c)",
            parts: storyContentTypeParts("footer", storyContentType("styles")),
            path: "word/footer9.xml",
        },
        { label: "glossary-styles-main-footer-coreach (s57d)", parts: glossaryStylesWithReadFooter(), path: "word/glossary/styles.xml" },
    ])("keeps $label without its session beside a Strict part (#1016)", async ({ parts, path }) => {
        await expectUncited(withStrictApp(parts), path);
        if (!path.endsWith("styles.xml")) await expectUncited(withStrictStyles(parts), path);
        // Sessions 53 and 57 stored the judged header or footer Transitional,
        // so its own Strict root drops the citation too (#1030).
        await expectUncited(withStrictPart(parts, path), path);
    });

    it.each(FALSE_NEGATIVES)(
        "rejects $label under word-valid citing session 57",
        async ({ parts, path: expectedPath, by }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "coreach.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file);
                expect(result.valid, "the full word-valid run fails").toBe(false);
                const found = result.issues.filter((issue) => issue.code === ID && issue.severity === "error");
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({ path: expectedPath });
                expect(found[0]?.message).toContain(by);
                expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 57).");
                expect(isWordBlockingIssue(found[0]!)).toBe(true);
                const partial = await validate(file, { checks: [ID] });
                expect(partial.valid).toBe(false);
            });
        },
        60_000,
    );

    it.each([
        {
            label: "footer1-referenced-typed-styles (s57c)",
            parts: storyContentTypeParts("footer", storyContentType("styles")),
            path: "word/footer9.xml",
        },
        {
            label: "footer1-referenced-typed-main (s57c)",
            parts: storyContentTypeParts("footer", storyContentType("document.main")),
            path: "word/footer9.xml",
        },
        {
            label: "footer1-referenced-typed-numbering (s57c)",
            parts: storyContentTypeParts("footer", storyContentType("numbering")),
            path: "word/footer9.xml",
        },
        {
            label: "glossary-styles-main-footer-coreach (s57d)",
            parts: glossaryStylesWithReadFooter(),
            path: "word/glossary/styles.xml",
        },
    ] as const)("keeps $label rejected, citing sessions 53 and 57", async ({ parts, path }) => {
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path });
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (sessions 53 and 57).");
        expect(isWordBlockingIssue(found[0]!)).toBe(true);
    });

    it("keeps a header typed with a type no session showed valid (header typed numbering)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unshown.docx");
            writeFileSync(file, await buildDocx(storyContentTypeParts("header", storyContentType("numbering"))));
            expect((await validate(file)).valid).toBe(true);
        });
    }, 60_000);
});

// Session 58 (2026-10-26, chunk s58b; Word 16.115, two runs, all readings
// agree) asked Word about the comments, endnotes and footnotes stories with
// no eager reach: a used part — its own root, the body's reference matching
// — typed by its Override with the styles content type gets the recovery
// offer on all three. The styles part itself, exact styles type, also the
// target of a comments, endnotes or footnotes relationship nothing
// references opens clean (s58a/s58b), so the observed shape is the read
// story typed styles, not the styles part co-reached.
describe("read story typed styles (#785, session 58)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const TEXT_P = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    /** Each story's own root holding the item the body references, and that reference run. */
    const STORIES = {
        comments: {
            root: `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="a">${TEXT_P}</w:comment></w:comments>`,
            ref: `<w:r><w:commentReference w:id="0"/></w:r>`,
        },
        endnotes: {
            root: `<w:endnotes xmlns:w="${W}"><w:endnote w:id="1">${TEXT_P}</w:endnote></w:endnotes>`,
            ref: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
        },
        footnotes: {
            root: `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${TEXT_P}</w:footnote></w:footnotes>`,
            ref: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
        },
    } as const;
    type Story = keyof typeof STORIES;
    /** A used `word/<story>.xml`, its own root holding the referenced item, typed `type` (its own by default). */
    const usedStoryParts = (story: Story, type: string = storyContentType(story)): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
            "</w:r></w:p>",
            `</w:r>${STORIES[story].ref}</w:p>`,
        ),
        [`word/${story}.xml`]: STORIES[story].root,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdS" Type="${R}/${story}" Target="${story}.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/${story}.xml" ContentType="${type}"/></Types>`,
        ),
    });
    /** `word/styles.xml`, exact styles type, also the target of a `story` relationship nothing references. */
    const stylesCoReachNoReference = (story: Story): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdX" Type="${R}/${story}" Target="styles.xml"/></Relationships>`,
        ),
    });

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "rejects a used %s part typed with the styles type, citing session 58",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "typed-styles.docx");
                writeFileSync(file, await buildDocx(usedStoryParts(story, storyContentType("styles"))));
                const result = await validate(file);
                expect(result.valid, "the full word-valid run fails").toBe(false);
                const found = result.issues.filter((issue) => issue.code === ID && issue.severity === "error");
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({ path: `word/${story}.xml` });
                expect(found[0]?.message).toContain(`the main document's ${story} relationship`);
                expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 58).");
                expect(isWordBlockingIssue(found[0]!)).toBe(true);
                // The own-type twin stays valid; for comments and endnotes it is
                // session 58's matching-reference calibration, which opened clean (s58a).
                writeFileSync(file, await buildDocx(usedStoryParts(story)));
                expect((await validate(file)).valid).toBe(true);
            });
        },
        60_000,
    );

    it.each(["comments", "endnotes", "footnotes"] as const)(
        "keeps an exact-typed styles part a %s relationship nothing references also targets valid",
        async (story) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "coreach.docx");
                writeFileSync(file, await buildDocx(stylesCoReachNoReference(story)));
                expect((await validate(file)).valid).toBe(true);
            });
        },
        60_000,
    );
});

// #996, Word session 81 (s81a, two runs, all readings agree): a read footer
// named by a target climbing past the package root is read at the root, as
// every read-model walk reads it since #986, so its styles typing got the
// recovery offer session 53 gave the in-package twin, and the footer typing
// opened clean with `Foot` rendered.
describe("a read footer named by a climbing target (#996, session 81)", () => {
    const judged = (parts: Record<string, string>) =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "climbing-footer.docx");
            writeFileSync(file, await buildDocx(parts));
            return validate(file, { profile: "word-valid" });
        });

    it("blocks /footer1.xml named by ../../footer1.xml and typed styles, as Word refused it", async () => {
        const result = await judged(climbingFooterTyped("styles"));
        expect(result.valid).toBe(false);
        expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => [issue.code, issue.path])).toEqual([
            [ID, "footer1.xml"],
        ]);
    });

    it("accepts the same footer typed footer, which Word opened clean", async () => {
        const result = await judged(climbingFooterTyped("footer"));
        expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        expect(result.valid).toBe(true);
    });
});
