/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A quick or named run applies its profile's policy, as a full run does
 * (#381): lenient never runs the Word rules or the ZIP-envelope checks, and
 * word-valid reports an error Word tolerates as a warning. So a partial run
 * never fails a package that the full run under the same profile passes.
 */

import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { poisonEntry } from "./support/poison-entry";
import { buildRawZip } from "./support/raw-zip";

const ROW_WITHOUT_CELL = path.join(
    path.dirname(fileURLToPath(import.meta.url)),
    "fixtures",
    "broken",
    "package.word-row-without-cell.docx",
);

const APP_CONTENT_TYPE = `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
const MALFORMED_APP = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace("</Types>", APP_CONTENT_TYPE),
    "docProps/app.xml": "<Properties",
};
const MALFORMED_MAIN = { ...MINIMAL_PARTS, "word/document.xml": "<w:document" };

const COMMENT_CONTENT_TYPES = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
    "</Types>",
    `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
        `<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>`,
);
const GOOD_COMMENTS = `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0"><w:p w14:paraId="11111111"/></w:comment></w:comments>`;
const DOCUMENT_WITH_MARKERS = MINIMAL_PARTS["word/document.xml"]!.replace(
    "<w:r><w:t>Hello</w:t></w:r>",
    `<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`,
);
const COMMENT_PARTS = {
    ...MINIMAL_PARTS,
    "[Content_Types].xml": COMMENT_CONTENT_TYPES,
    // The relationships a package whose document carries comment markers
    // really has: Word reads the comments part through the commentReference
    // (#437), and the commentsExtended part through its Microsoft extension
    // type, so both stay Word-blocking when malformed.
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
            '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' +
            "</Relationships>",
    ),
    "word/document.xml": DOCUMENT_WITH_MARKERS,
    "word/comments.xml": GOOD_COMMENTS,
};
const MALFORMED_COMMENT_PACKAGES = [
    { name: "malformed-comments", part: "word/comments.xml", parts: { ...COMMENT_PARTS, "word/comments.xml": "<w:comments" } },
    { name: "malformed-document", part: "word/document.xml", parts: { ...COMMENT_PARTS, "word/document.xml": "<w:document" } },
    {
        name: "malformed-extended",
        part: "word/commentsExtended.xml",
        parts: { ...COMMENT_PARTS, "word/commentsExtended.xml": "<w15:commentsEx" },
    },
];

// The comments part stored as word/COMMENTS.XML: the same part name, in
// another case (#390).
const CASE_COMMENT_PARTS: Readonly<Record<string, string>> = {
    ...Object.fromEntries(Object.entries(COMMENT_PARTS).filter(([name]) => name !== "word/comments.xml")),
    "[Content_Types].xml": COMMENT_CONTENT_TYPES.replace("/word/comments.xml", "/word/COMMENTS.XML"),
    "word/COMMENTS.XML": "<w:comments",
};
// A malformed comments part beside a well-formed document with an orphan
// comment-range end, which only the comment rules report (#391).
const ORPHAN_BESIDE_MALFORMED_PARTS = {
    ...COMMENT_PARTS,
    "word/comments.xml": "<w:comments",
    "word/document.xml": DOCUMENT_WITH_MARKERS.replace('<w:commentRangeStart w:id="0"/>', ""),
};
// Every Word part under WORD/, with the main document part malformed (#392).
const UPPER_WORD_PARTS = Object.fromEntries(
    Object.entries({ ...MINIMAL_PARTS, "word/document.xml": "<w:document" }).map(([name, xml]) => [name.replace(/^word\//, "WORD/"), xml]),
);
// A malformed content-types stream stored as [CONTENT_TYPES].XML (review Q).
const UPPER_CONTENT_TYPES_PARTS = {
    ...Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== "[Content_Types].xml")),
    "[CONTENT_TYPES].XML": "<Types",
};
// Package relationships Word cannot read: it offers recovery for each (#388).
// The xml-syntax check reports both codes.
const UNREADABLE_PACKAGE_RELS = [
    { name: "zero-byte", code: "xml-syntax", xml: "" },
    { name: "truncated", code: "xml-syntax", xml: MINIMAL_PARTS["_rels/.rels"]!.slice(0, 120) },
    {
        name: "undeclared-prefix",
        code: "xml-undeclared-prefix",
        xml: MINIMAL_PARTS["_rels/.rels"]!.replace("<Relationships ", "<x:Relationships ").replace(
            "</Relationships>",
            "</x:Relationships>",
        ),
    },
] as const;
// A main part declaring an encoding no decoder knows: Word refuses the file,
// and opens the windows-1252 control clean (#397).

describe("a partial run reports a refused item name as the full run does (#413)", { timeout: 120_000 }, () => {
    // Stored item names the package view refuses (#372, #378). JSZip rewrites
    // `..`, `.` and empty segments at load, so both runs judge the stored
    // spelling, through `entryNameRefusal`.
    const REFUSED: readonly (readonly [string, string])[] = [
        ["dotdot", "word/a/../x.xml"],
        ["dot-only", "word/./x.xml"],
        ["empty-segment", "word//x.xml"],
        ["drive", "c:doc/x.xml"],
    ];
    const SELECTIONS = [{ checks: ["xml-syntax"] }, { checks: ["package"] }, { quick: true }] as const;

    const issue = (result: { issues: readonly ValidationIssue[] }, code: string): ValidationIssue | undefined =>
        result.issues.find((entry) => entry.severity === "error" && entry.code === code);

    describe.each(REFUSED)("the stored item name %j", (label, name) => {
        it.each(["word-valid", "lenient"] as const)(
            "under %s, every partial selection reports the full run's failure, at the same code and message",
            async (profile) => {
                await withTempDir(async (dir) => {
                    const file = await specimen(dir, `${label}.docx`, await buildDocx({ ...MINIMAL_PARTS, [name]: "<x/>" }));
                    const full = await validate(file, { profile });
                    const fullFailure = issue(full, "package-extract-failed");

                    expect(full.valid).toBe(false);
                    expect(fullFailure).toBeDefined();
                    for (const selection of SELECTIONS) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${label} ${profile} ${JSON.stringify(selection)}`;

                        // #381: partial errors are a subset of full errors as
                        // code@path, and a valid full run means a valid partial
                        // run (both are invalid here).
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                        expect(partial.valid, tag).toBe(false);
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        expect(issue(partial, "package-extract-failed")?.message, tag).toBe(fullFailure!.message);
                    }
                });
            },
        );

        it("strict still refuses a partial run of it", async () => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${label}-strict.docx`, await buildDocx({ ...MINIMAL_PARTS, [name]: "<x/>" }));
                for (const selection of SELECTIONS) {
                    await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toBeInstanceOf(
                        ModeConflictError,
                    );
                }
            });
        });
    });

    it("agrees on a directory holding a drive-relative name", async () => {
        // The other three names cannot be staged as directories: a filesystem
        // resolves `word/a/../x.xml`, `word/./x.xml` and `word//x.xml` to
        // `word/x.xml`, so the directory never holds them. `c:doc` is a plain
        // directory name on POSIX; Windows forbids it.
        if (process.platform === "win32") return;
        await withTempDir(async (dir) => {
            const root = unpacked(dir, "drive.docx", { ...MINIMAL_PARTS, "c:doc/x.xml": "<x/>" });
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(root, { profile });
                for (const selection of SELECTIONS) {
                    const partial = await validate(root, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;

                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            }
        });
    });

    it("reports an unreadable entry as the same extract failure in full and partial runs", async () => {
        // Control: a stored entry whose compressed stream is damaged opens as
        // a ZIP but cannot be read, so both views refuse it (#369).
        await withTempDir(async (dir) => {
            const poisoned = poisonEntry(await buildDocx({ ...MINIMAL_PARTS, "word/extra.xml": "<x/>" }), "word/extra.xml");
            const file = await specimen(dir, "poisoned.docx", Buffer.from(poisoned));
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                const fullFailure = issue(full, "package-extract-failed");

                expect(fullFailure, profile).toBeDefined();
                for (const selection of SELECTIONS) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;

                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(partial.valid, tag).toBe(false);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    expect(issue(partial, "package-extract-failed")?.message, tag).toBe(fullFailure!.message);
                }
            }
        });
    });

    it("keeps package-open-failed, in full and partial runs, for bytes that are not a ZIP", async () => {
        // Control: the archive itself cannot be opened, which stays
        // package-open-failed in every run — only the package-view refusal
        // changed.
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "notzip.docx", Buffer.from("this is not a zip archive at all"));
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                const fullFailure = issue(full, "package-open-failed");

                expect(fullFailure, profile).toBeDefined();
                for (const selection of SELECTIONS) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;

                    expect(partial.valid, tag).toBe(false);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    expect(issue(partial, "package-open-failed")?.message, tag).toBe(fullFailure!.message);
                }
            }
        });
    });
});

const declaringEncoding = (encoding: string): Readonly<Record<string, string>> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace('encoding="UTF-8"', `encoding="${encoding}"`),
});
// The comments part named with an escaped `x` in its extension (review Q).

const ESCAPED_COMMENT_PARTS: Readonly<Record<string, string>> = {
    ...Object.fromEntries(Object.entries(COMMENT_PARTS).filter(([name]) => name !== "word/comments.xml")),
    "word/comments.%78ml": "<w:comments",
};

const errorCodes = (issues: readonly ValidationIssue[]): string[] => [
    ...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "")),
];

const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    [...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code ?? ""}@${issue.path ?? ""}`))].sort();

const specimen = async (dir: string, name: string, bytes: Buffer): Promise<string> => {
    const file = path.join(dir, name);
    writeFileSync(file, bytes);
    return file;
};

const unpacked = (dir: string, name: string, parts: Readonly<Record<string, string>>): string => {
    const root = path.join(dir, name);
    for (const [part, xml] of Object.entries(parts)) {
        mkdirSync(path.dirname(path.join(root, part)), { recursive: true });
        writeFileSync(path.join(root, part), xml);
    }
    return root;
};

const trailingBytes = async (): Promise<Buffer> => Buffer.concat([await buildDocx(), Buffer.from("trailing junk")]);

describe("a partial run of a refused package view still runs the stored-name checks (#420)", { timeout: 120_000 }, () => {
    // A stored name that places no part refuses the package view (#378), but
    // the stored-name checks (#351, #408) read only the names — the full run
    // reports them before extraction, and a partial run that selects them must
    // report the same findings.
    const STORED_NAME_CHECKS = [
        "opc-equivalent-part-names",
        "opc-part-name-escaped-unreserved",
        "opc-part-name-escaped-slash",
        "opc-part-name-word-refused",
    ] as const;
    const SELECTIONS = [
        { checks: ["opc-equivalent-part-names"] },
        { checks: ["opc-part-name-escaped-unreserved"] },
        { checks: [...STORED_NAME_CHECKS] },
        { checks: ["package"] },
        { quick: true },
    ] as const;
    const APPLICABLE = createDefaultCheckRegistry()
        .list("docx")
        .map((definition) => definition.id);

    const rawPackage = (entries: readonly { name: string; content?: string }[], drop?: string): Buffer =>
        buildRawZip([
            ...Object.entries(MINIMAL_PARTS)
                .filter(([part]) => part !== drop)
                .map(([part, text]) => ({ name: part, content: text })),
            ...entries,
        ]);

    /** The stored-name check ids a selection names — directly or via the package group / quick preset. */
    const selectedStoredNameChecks = (selection: (typeof SELECTIONS)[number]): readonly string[] =>
        STORED_NAME_CHECKS.filter((id) => !("checks" in selection) || selection.checks.some((name) => name === id || name === "package"));

    const SPECIMENS: readonly {
        readonly name: string;
        readonly code: string;
        readonly drop?: string;
        readonly entries: readonly { readonly name: string; readonly content: string }[];
    }[] = [
        {
            name: "escaped-dotdot",
            code: "opc-part-name-escaped-unreserved",
            entries: [
                { name: "word/%78tra.xml", content: "<x/>" },
                { name: "word/a/../b.xml", content: "<x/>" },
            ],
        },
        {
            name: "escaped-slash",
            code: "opc-part-name-escaped-unreserved",
            entries: [
                { name: "word/%78tra.xml", content: "<x/>" },
                { name: "/abs.xml", content: "<x/>" },
            ],
        },
        {
            name: "equivalent-dotdot",
            code: "opc-equivalent-part-names",
            entries: [
                { name: "word/a/../b.xml", content: "<x/>" },
                { name: "word/STYLES.xml", content: MINIMAL_PARTS["word/styles.xml"]! },
            ],
        },
        {
            name: "identical-dotdot",
            code: "opc-equivalent-part-names",
            drop: "word/document.xml",
            entries: [
                { name: "word/document.xml", content: MINIMAL_PARTS["word/document.xml"]! },
                { name: "word/document.xml", content: "<x/>" },
                { name: "word/a/../b.xml", content: "<x/>" },
            ],
        },
    ] as const;

    describe.each(SPECIMENS)("$name", ({ name, code, drop, entries }) => {
        it.each(["word-valid", "lenient"] as const)(
            "reports the stored-name findings and the extract failure as the full %s run does",
            async (profile) => {
                await withTempDir(async (dir) => {
                    const file = await specimen(dir, `${name}.docx`, rawPackage(entries, drop));
                    const full = await validate(file, { profile });
                    const findings = full.issues.filter((issue) => issue.code === code);
                    const failure = full.issues.find((issue) => issue.code === "package-extract-failed");

                    expect(findings.length, `${name}: full run reports the stored-name finding`).toBeGreaterThan(0);
                    expect(failure, `${name}: full run reports the refused package view`).toBeDefined();
                    for (const selection of SELECTIONS) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${name} ${profile} ${JSON.stringify(selection)}`;

                        // The stored-name checks the selection names report
                        // exactly what the full run reports: code, path,
                        // message and severity.
                        const named = selectedStoredNameChecks(selection);
                        expect(
                            partial.issues.filter((issue) => named.includes(issue.code ?? "")),
                            tag,
                        ).toEqual(full.issues.filter((issue) => named.includes(issue.code ?? "")));
                        expect(partial.issues.find((issue) => issue.code === "package-extract-failed")?.message, tag).toBe(
                            failure!.message,
                        );
                        // A selected id is never in neither list (#420): every
                        // applicable check is either executed or omitted.
                        expect(
                            new Set([...(partial.metadata?.executedChecks ?? []), ...(partial.metadata?.omittedChecks ?? [])]),
                            tag,
                        ).toEqual(new Set(APPLICABLE));
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
                });
            },
        );
    });

    it.each(["word-valid", "lenient"] as const)(
        "lists every stored-name check as executed under %s when the selection names it",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, "executed.docx", rawPackage([{ name: "word/a/../b.xml", content: "<x/>" }]));
                for (const selection of SELECTIONS) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    const named = selectedStoredNameChecks(selection);
                    expect(partial.metadata?.executedChecks, tag).toEqual(expect.arrayContaining([...named]));
                    // A stored-name check the selection did not name stays unrun.
                    expect(
                        STORED_NAME_CHECKS.filter((id) => !named.includes(id) && (partial.metadata?.executedChecks ?? []).includes(id)),
                        tag,
                    ).toEqual([]);
                }
            });
        },
    );

    it.each(["word-valid", "lenient", "strict"] as const)(
        "the full %s run diagnoses exactly the checks flagged as needing only stored names",
        async (profile) => {
            // One flag decides both paths, so a check cannot run on a refused
            // view in a partial run and be skipped by the full run, or the reverse.
            const registry = createDefaultCheckRegistry();
            const flagged = registry
                .list("docx")
                .filter((definition) => definition.storedNamesOnly === true)
                .map((definition) => definition.id);
            const envelope = new Set(
                registry
                    .list("docx")
                    .filter((definition) => definition.archiveOnly === true)
                    .map((definition) => definition.id),
            );
            await withTempDir(async (dir) => {
                const file = await specimen(dir, "flagged.docx", rawPackage([{ name: "word/a/../b.xml", content: "<x/>" }]));
                const full = await validate(file, { profile });
                expect(full.issues.some((issue) => issue.code === "package-extract-failed")).toBe(true);
                expect((full.metadata?.executedChecks ?? []).filter((id) => !envelope.has(id)).sort()).toEqual(
                    [...STORED_NAME_CHECKS].sort(),
                );
                expect([...flagged].sort()).toEqual([...STORED_NAME_CHECKS].sort());
            });
        },
    );

    it("runs a selected envelope check on a refused view, as the full run does", async () => {
        // An archive-only check reads the archive bytes, not the package view,
        // so a refused view does not omit it (CodeRabbit on #440).
        await withTempDir(async (dir) => {
            const bytes = Buffer.concat([rawPackage([{ name: "word/a/../b.xml", content: "<x/>" }]), Buffer.from("trailing junk")]);
            const file = await specimen(dir, "refused-trailing.docx", bytes);
            const partial = await validate(file, { checks: ["zip-bytes-after-eocd"] });
            const full = await validate(file, {});

            expect(partial.metadata?.executedChecks).toEqual(["zip-bytes-after-eocd"]);
            expect(partial.metadata?.omittedChecks).not.toContain("zip-bytes-after-eocd");
            expect(partial.issues.map((issue) => issue.code)).toEqual(
                expect.arrayContaining(["zip-bytes-after-eocd", "package-extract-failed"]),
            );
            expect(full.issues.map((issue) => issue.code)).toEqual(
                expect.arrayContaining(["zip-bytes-after-eocd", "package-extract-failed"]),
            );
        });
    });
});

describe("a partial run applies its profile's policy (#381)", { timeout: 120_000 }, () => {
    it("lenient quick runs no Word rule, as full lenient does not", async () => {
        const result = await validate(ROW_WITHOUT_CELL, { profile: "lenient", quick: true });

        expect(result.valid).toBe(true);
        expect(result.issues.filter((issue) => issue.code === "word-row-without-cell")).toEqual([]);
        // ct-duplicate-declaration is an OPC content-types check, not a Word rule:
        // full lenient reports duplicates as warnings, so lenient quick runs it (#431).
        expect(result.metadata?.executedChecks).toEqual([
            "opc-equivalent-part-names",
            "opc-part-name-escaped-unreserved",
            "opc-part-name-escaped-slash",
            "opc-part-name-word-refused",
            "xml-syntax",
            "ct-duplicate-declaration",
            // The content-types group's Override PartName grammar (#414).
            "ct-override-partname-invalid",
            "ct-override-partname-word-refused",
            // Its Markup Compatibility refusal (#459).
            "ct-markup-compatibility",
            "word-part-content-type-generic",
        ]);
        expect(result.metadata?.omittedChecks).toEqual(expect.arrayContaining(["word-row-without-cell", "zip-bytes-after-eocd"]));
        expect(errorCodes((await validate(ROW_WITHOUT_CELL, { profile: "lenient" })).issues)).not.toContain("word-row-without-cell");
    });

    it("lenient quick runs no ZIP-envelope check, as full lenient does not", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "trailing.docx", await trailingBytes());
            const result = await validate(file, { profile: "lenient", quick: true });

            expect(result.issues.filter((issue) => issue.code === "zip-bytes-after-eocd")).toEqual([]);
            expect(result.metadata?.executedChecks).not.toContain("zip-bytes-after-eocd");
        });
    });

    it("word-valid quick reports an error Word tolerates as a warning, as the full run does", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "malformed-app.docx", await buildDocx(MALFORMED_APP));
            const quick = await validate(file, { quick: true });

            expect(quick.metadata?.profile).toBe("word-valid");
            expect(quick.valid).toBe(true);
            expect(quick.issues.filter((issue) => issue.code === "xml-syntax")).toEqual([
                expect.objectContaining({ severity: "warning", path: "docProps/app.xml" }),
            ]);
            expect((await validate(file, {})).valid).toBe(true);
        });
    });

    it("word-valid quick keeps the errors Word rejects", async () => {
        await withTempDir(async (dir) => {
            const main = await validate(await specimen(dir, "malformed-main.docx", await buildDocx(MALFORMED_MAIN)), { quick: true });
            expect(main.issues.filter((issue) => issue.code === "xml-syntax")).toEqual([
                expect.objectContaining({ severity: "error", path: "word/document.xml" }),
            ]);

            const row = await validate(ROW_WITHOUT_CELL, { quick: true });
            expect(errorCodes(row.issues)).toContain("word-row-without-cell");
        });
    });

    it("word-valid quick keeps the ZIP-envelope errors the full run reports", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "trailing.docx", await trailingBytes());

            expect(errorCodes((await validate(file, { quick: true })).issues)).toContain("zip-bytes-after-eocd");
            expect(errorCodes((await validate(file, {})).issues)).toContain("zip-bytes-after-eocd");
        });
    });

    it.each([[["word-row-without-cell"]], [["xml-syntax", "word-row-without-cell"]], [["zip-bytes-after-eocd"]], [["tables"]]])(
        "refuses a lenient run of %j, which names checks lenient never runs",
        async (checks) => {
            const run = validate(ROW_WITHOUT_CELL, { profile: "lenient", checks });

            await expect(run).rejects.toBeInstanceOf(ModeConflictError);
            await expect(validate(ROW_WITHOUT_CELL, { profile: "lenient", checks })).rejects.toThrow(/lenient/);
        },
    );

    it("refuses a selection lenient never runs before touching the filesystem", async () => {
        // A contradictory selection is a configuration error, not a verdict
        // about a document: it wins over a missing file, as the other mode
        // conflicts do.
        await withTempDir(async (dir) => {
            const missing = path.join(dir, "missing.docx");
            await expect(validate(missing, { profile: "lenient", checks: ["tables"] })).rejects.toBeInstanceOf(ModeConflictError);
        });
    });

    it.each(["word-valid", "lenient"] as const)("runs the shared comment checks under %s", async (profile) => {
        const result = await validate(ROW_WITHOUT_CELL, { profile, checks: ["comments"] });

        expect(result.valid).toBe(true);
        expect(result.metadata?.executedChecks).toEqual(["xml-syntax", "comment-markers", "comment-threading"]);
        expect(result.metadata?.omittedChecks).not.toContain("xml-syntax");
    });

    describe.each(["word-valid", "lenient"] as const)("comment parse diagnostics under %s (#385)", (profile) => {
        it.each(MALFORMED_COMMENT_PACKAGES)("reports $name once through xml-syntax", async ({ name, part, parts }) => {
            await withTempDir(async (dir) => {
                const result = await validate(await specimen(dir, `${name}.docx`, await buildDocx(parts)), {
                    profile,
                    checks: ["comments"],
                });

                expect(result.valid).toBe(false);
                expect(result.issues).toContainEqual(expect.objectContaining({ code: "xml-syntax", severity: "error", path: part }));
                expect(result.issues.filter((issue) => issue.path === part && /^comment-.*-parse$/.test(issue.code ?? ""))).toEqual([]);
                // Lenient stops at the malformed part, as its full run does (#391).
                expect(result.metadata?.executedChecks).toEqual(
                    profile === "lenient" ? ["xml-syntax"] : ["xml-syntax", "comment-markers", "comment-threading"],
                );
                expect(result.metadata?.omittedChecks).not.toContain("xml-syntax");
            });
        });
    });

    it("keeps judging the comment surface alongside a tolerated malformed app part (#385)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...COMMENT_PARTS,
                "[Content_Types].xml": COMMENT_CONTENT_TYPES.replace("</Types>", APP_CONTENT_TYPE),
                "docProps/app.xml": MALFORMED_APP["docProps/app.xml"],
                "word/document.xml": DOCUMENT_WITH_MARKERS.replace('<w:commentRangeEnd w:id="0"/>', ""),
            };
            const result = await validate(await specimen(dir, "app-with-comments.docx", await buildDocx(parts)), { checks: ["comments"] });

            expect(result.valid).toBe(true);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "xml-syntax", severity: "warning", path: "docProps/app.xml" }),
            );
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "comment-orphan-start", severity: "warning", path: "word/document.xml" }),
            );
            expect(result.issues).toContainEqual(expect.objectContaining({ code: "comment-thread-count-mismatch", severity: "warning" }));
            expect(result.metadata?.executedChecks).toEqual(["xml-syntax", "comment-markers", "comment-threading"]);
        });
    });

    it("keeps the direct marker API's @xmldom parse message (#385)", async () => {
        await withTempDir(async (dir) => {
            const parts = MALFORMED_COMMENT_PACKAGES[1]!.parts;
            for (const [part, xml] of Object.entries(parts)) {
                mkdirSync(path.dirname(path.join(dir, part)), { recursive: true });
                writeFileSync(path.join(dir, part), xml);
            }
            const parsed = ValidationContext.fromEntries(parts).xml("word/document.xml");
            expect(parsed?.ok).toBe(false);
            const result = await new DOCXSchemaValidator({ unpackedDir: dir }).validateCommentMarkers();
            expect(result.issues).toEqual([
                {
                    code: "comment-marker-parse",
                    severity: "error",
                    path: "word/document.xml",
                    message: `Error parsing XML: ${parsed && !parsed.ok ? parsed.message : ""}`,
                },
            ]);
        });
    });

    it("runs the checks of a named group that lenient runs, and omits the rest", async () => {
        const result = await validate(ROW_WITHOUT_CELL, { profile: "lenient", checks: ["xml", "tables"] });

        expect(result.metadata?.executedChecks).toEqual(["xml-syntax"]);
        expect(result.metadata?.omittedChecks).toContain("word-row-without-cell");
    });

    it.each([["comments"], ["comment-markers"], ["comment-threading"]])(
        "stops a lenient %s run at a malformed part, as the full lenient run does (#391)",
        async (name) => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, "orphan-beside-malformed.docx", await buildDocx(ORPHAN_BESIDE_MALFORMED_PARTS));
                const full = await validate(file, { profile: "lenient" });
                const partial = await validate(file, { profile: "lenient", checks: [name] });

                expect(errorPairs(partial.issues)).toEqual(["xml-syntax@word/comments.xml", "xml-undeclared-prefix@word/comments.xml"]);
                expect(errorPairs(partial.issues)).toEqual(errorPairs(full.issues));
                expect(partial.metadata?.executedChecks).toEqual(["xml-syntax"]);
                // A selected check the stop kept from running is not an omission.
                expect(partial.metadata?.omittedChecks).not.toContain(name === "comments" ? "comment-markers" : name);
                // Word-valid keeps judging the parts that parsed (#307).
                const wordValid = await validate(file, { checks: [name] });
                expect(wordValid.metadata?.executedChecks).toEqual(
                    expect.arrayContaining(["xml-syntax", name === "comments" ? "comment-markers" : name]),
                );
            });
        },
    );

    it("never fails a package that the full run under the same profile passes", async () => {
        await withTempDir(async (dir) => {
            const packages = [
                ROW_WITHOUT_CELL,
                await specimen(dir, "minimal.docx", await buildDocx()),
                await specimen(dir, "malformed-app.docx", await buildDocx(MALFORMED_APP)),
                await specimen(dir, "malformed-main.docx", await buildDocx(MALFORMED_MAIN)),
                await specimen(dir, "trailing.docx", await trailingBytes()),
                ...(await Promise.all(
                    MALFORMED_COMMENT_PACKAGES.map(async ({ name, parts }) => specimen(dir, `${name}.docx`, await buildDocx(parts))),
                )),
                await specimen(dir, "case-comments.docx", await buildDocx(CASE_COMMENT_PARTS)),
                unpacked(dir, "case-comments-unpacked.docx", CASE_COMMENT_PARTS),
                await specimen(dir, "orphan-beside-malformed.docx", await buildDocx(ORPHAN_BESIDE_MALFORMED_PARTS)),
                await specimen(dir, "upper-word.docx", await buildDocx(UPPER_WORD_PARTS)),
                await specimen(dir, "upper-content-types.docx", await buildDocx(UPPER_CONTENT_TYPES_PARTS)),
                await specimen(dir, "escaped-comments.docx", await buildDocx(ESCAPED_COMMENT_PARTS)),
            ];
            for (const file of packages) {
                for (const profile of ["lenient", "word-valid"] as const) {
                    const label = `${path.basename(file)} under ${profile}`;
                    const full = await validate(file, { profile });
                    for (const selection of [{ quick: true }, { checks: ["comments"] }] as const) {
                        const partial = await validate(file, { profile, ...selection });
                        // Never fails what the full run passes; never reports an
                        // error, at a path, that the full run does not (#381).
                        if (full.valid) expect(partial.valid, `${label}: ${JSON.stringify(selection)}`).toBe(true);
                        expect(errorPairs(full.issues), `${label}: ${JSON.stringify(selection)}`).toEqual(
                            expect.arrayContaining(errorPairs(partial.issues)),
                        );
                    }
                }
            }
        });
    });
});

describe("a part named in another case gets the verdict its lower-case name gets (#390, #392)", { timeout: 120_000 }, () => {
    it.each(["lenient", "word-valid", "strict"] as const)(
        "the full run rejects a malformed word/COMMENTS.XML under %s",
        async (profile) => {
            await withTempDir(async (dir) => {
                const result = await validate(await specimen(dir, "case-comments.docx", await buildDocx(CASE_COMMENT_PARTS)), { profile });

                expect(result.valid).toBe(false);
                expect(result.issues).toContainEqual(
                    expect.objectContaining({ code: "xml-syntax", severity: "error", path: "word/COMMENTS.XML" }),
                );
                expect(
                    result.issues.filter((issue) => issue.path === "word/COMMENTS.XML" && /^comment-.*-parse$/.test(issue.code ?? "")),
                ).toEqual([]);
            });
        },
    );

    it.each(["packed", "unpacked"] as const)("a partial run over a %s package agrees with the full run on it", async (input) => {
        await withTempDir(async (dir) => {
            const target =
                input === "packed"
                    ? await specimen(dir, "case-comments.docx", await buildDocx(CASE_COMMENT_PARTS))
                    : unpacked(dir, "case-comments.docx", CASE_COMMENT_PARTS);
            for (const profile of ["lenient", "word-valid"] as const) {
                const full = await validate(target, { profile });
                for (const selection of [{ quick: true }, { checks: ["comments"] }] as const) {
                    const partial = await validate(target, { profile, ...selection });
                    const label = `${profile} ${JSON.stringify(selection)}`;

                    expect(partial.valid, label).toBe(full.valid);
                    expect(errorPairs(full.issues), label).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            }
        });
    });

    it.each([{}, { quick: true }] as const)("word-valid rejects a malformed main part stored under WORD/ (%j)", async (selection) => {
        await withTempDir(async (dir) => {
            const result = await validate(await specimen(dir, "upper-word.docx", await buildDocx(UPPER_WORD_PARTS)), {
                profile: "word-valid",
                ...selection,
            });

            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "xml-syntax", severity: "error", path: "WORD/document.xml" }),
            );
        });
    });

    it.each([{}, { quick: true }] as const)("word-valid rejects a malformed [CONTENT_TYPES].XML (%j)", async (selection) => {
        await withTempDir(async (dir) => {
            const result = await validate(await specimen(dir, "upper-content-types.docx", await buildDocx(UPPER_CONTENT_TYPES_PARTS)), {
                profile: "word-valid",
                ...selection,
            });

            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "xml-syntax", severity: "error", path: "[CONTENT_TYPES].XML" }),
            );
        });
    });

    it.skipIf(process.platform === "win32")(
        "word-valid warns on a directory's malformed word\\foo.xml in full and quick runs alike (review Q, #437)",
        async () => {
            await withTempDir(async (dir) => {
                const root = unpacked(dir, "backslash.docx", { ...MINIMAL_PARTS, "word\\foo.xml": "<w:foo" });
                for (const selection of [{}, { quick: true }] as const) {
                    const result = await validate(root, { profile: "word-valid", ...selection });

                    // No relationship names word/foo.xml, so Word 16.115 does
                    // not read it (#437); the run still judges the part, as the
                    // lower-case spelling, in full and quick runs alike.
                    expect(result.valid, JSON.stringify(selection)).toBe(true);
                    expect(result.issues).toContainEqual(
                        expect.objectContaining({ code: "xml-syntax", severity: "warning", path: "word/foo.xml" }),
                    );
                }
            });
        },
    );

    it.each(["lenient", "word-valid"] as const)(
        "judges a malformed comments part with an escaped extension in full and partial runs under %s",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, "escaped-comments.docx", await buildDocx(ESCAPED_COMMENT_PARTS));
                for (const selection of [{}, { quick: true }, { checks: ["comments"] }] as const) {
                    const result = await validate(file, { profile, ...selection });
                    const label = JSON.stringify(selection);

                    expect(result.valid, label).toBe(false);
                    expect(result.issues, label).toContainEqual(
                        expect.objectContaining({ code: "xml-syntax", severity: "error", path: "word/comments.%78ml" }),
                    );
                }
            });
        },
    );
});

describe("word-valid rejects package relationships Word cannot read (#388)", { timeout: 120_000 }, () => {
    it.each(UNREADABLE_PACKAGE_RELS)("rejects an unreadable _rels/.rels ($name) in full and partial runs", async ({ name, code, xml }) => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, `${name}-rels.docx`, await buildDocx({ ...MINIMAL_PARTS, "_rels/.rels": xml }));
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(file, { profile: "word-valid", ...selection });
                const label = JSON.stringify(selection);

                expect(result.valid, label).toBe(false);
                expect(result.issues, label).toContainEqual(expect.objectContaining({ code, severity: "error", path: "_rels/.rels" }));
            }
        });
    });

    it("passes the readable _rels/.rels control in full and partial runs", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "control.docx", await buildDocx());
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(file, { profile: "word-valid", ...selection });

                expect(result.valid, JSON.stringify(selection)).toBe(true);
            }
        });
    });
});

describe("word-valid rejects a main part Word cannot decode (#397)", { timeout: 120_000 }, () => {
    it("rejects an unknown declared encoding in full and partial runs", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "unknown-encoding.docx", await buildDocx(declaringEncoding("X-UNKNOWN-ENCODING")));
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(file, { profile: "word-valid", ...selection });
                const label = JSON.stringify(selection);

                expect(result.valid, label).toBe(false);
                expect(result.issues, label).toContainEqual(
                    expect.objectContaining({ code: "xml-encoding", severity: "error", path: "word/document.xml" }),
                );
            }
        });
    });

    it("passes the windows-1252 control in full and partial runs", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "windows-1252.docx", await buildDocx(declaringEncoding("windows-1252")));
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(file, { profile: "word-valid", ...selection });

                expect(result.valid, JSON.stringify(selection)).toBe(true);
            }
        });
    });
});

describe("M1.8 selection parity (#408, #381)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "lenient"] as const)("keeps named, package and quick findings equal to full %s findings", async (profile) => {
        await withTempDir(async (dir) => {
            const code = "opc-part-name-escaped-unreserved";
            const file = path.join(dir, "escaped.docx");
            writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "word/%78tra.xml": "<root/>" }));
            const full = await validate(file, { profile });
            const issues = full.issues.filter((issue) => issue.code === code);
            expect(issues).toHaveLength(1);
            for (const selection of [{ checks: [code] }, { checks: ["package"] }, { quick: true }]) {
                const partial = await validate(file, { profile, ...selection });
                expect(partial.issues.filter((issue) => issue.code === code)).toEqual(issues);
                if (full.valid) expect(partial.valid).toBe(true);
                expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            }
        });
    });
});

describe("M1.12 selection parity for entries ZIP loading hides (#411)", { timeout: 120_000 }, () => {
    // JSZip keeps only the last entry under a given name, so two entries
    // stored under the identical name need raw archive bytes (JSZip cannot
    // write them). The stored-name diagnosis must see both anyway.
    const identical = (first: string, second: string): Buffer =>
        buildRawZip([
            ...Object.entries(MINIMAL_PARTS)
                .filter(([part]) => part !== "word/document.xml")
                .map(([part, text]) => ({ name: part, content: text })),
            { name: "word/document.xml", content: first },
            { name: "word/document.xml", content: second },
        ]);
    const identicalFinding = (severity: string) => ({
        code: "opc-equivalent-part-names",
        severity,
        path: "word/document.xml",
        message:
            'OPC [M1.12] forbids equivalent stored names: "word/document.xml", "word/document.xml". The last stored entry is the one read.',
    });

    it.each(["word-valid", "lenient"] as const)("keeps named, package and quick findings equal to full %s findings", async (profile) => {
        await withTempDir(async (dir) => {
            const bytes = identical(MINIMAL_PARTS["word/document.xml"]!, MINIMAL_PARTS["word/document.xml"]!.replace(">Hello<", ">Other<"));
            const file = await specimen(dir, "identical.docx", bytes);
            const full = await validate(file, { profile });
            const issues = full.issues.filter((issue) => issue.code === "opc-equivalent-part-names");
            expect(issues).toEqual([identicalFinding(profile === "lenient" ? "warning" : "error")]);
            if (profile === "word-valid") expect(full.valid).toBe(false);
            for (const selection of [{ checks: ["opc-equivalent-part-names"] }, { checks: ["package"] }, { quick: true }] as const) {
                const partial = await validate(file, { profile, ...selection });
                const tag = `${profile} ${JSON.stringify(selection)}`;
                expect(
                    partial.issues.filter((issue) => issue.code === "opc-equivalent-part-names"),
                    tag,
                ).toEqual(issues);
                if (full.valid) expect(partial.valid, tag).toBe(true);
                expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            }
        });
    });

    // A stored `..` segment names no part (#378). Stored first, JSZip's
    // resolution hides it behind the later `word/x.xml`; full and partial
    // runs must report the same extract failure in either order (#413).
    it.each(["word-valid", "lenient"] as const)(
        "reports the hidden refused name as the same extract failure in full and partial %s runs, in either order",
        async (profile) => {
            await withTempDir(async (dir) => {
                const messages: string[] = [];
                for (const [order, entries] of [
                    [
                        { name: "word/deep/../x.xml", content: "<x/>" },
                        { name: "word/x.xml", content: "<x/>" },
                    ],
                    [
                        { name: "word/x.xml", content: "<x/>" },
                        { name: "word/deep/../x.xml", content: "<x/>" },
                    ],
                ].entries()) {
                    const file = await specimen(
                        dir,
                        `order-${order}.docx`,
                        buildRawZip([...Object.entries(MINIMAL_PARTS).map(([part, text]) => ({ name: part, content: text })), ...entries]),
                    );
                    const full = await validate(file, { profile });
                    const fullFailure = full.issues.find((entry) => entry.severity === "error" && entry.code === "package-extract-failed");
                    expect(fullFailure, `order ${order}`).toBeDefined();
                    expect(full.valid, `order ${order}`).toBe(false);
                    for (const selection of [
                        { checks: ["opc-equivalent-part-names"] },
                        { checks: ["package"] },
                        { quick: true },
                    ] as const) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${profile} ${JSON.stringify(selection)} order ${order}`;
                        expect(partial.valid, tag).toBe(false);
                        expect(
                            partial.issues.find((entry) => entry.code === "package-extract-failed")?.message.replace(file, "FILE"),
                            tag,
                        ).toBe(fullFailure!.message.replace(file, "FILE"));
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
                    messages.push(fullFailure!.message.replace(file, "FILE"));
                }
                expect(messages[0]).toBe(messages[1]);
            });
        },
    );
});
