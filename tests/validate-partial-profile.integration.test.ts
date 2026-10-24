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
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

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

const errorCodes = (issues: readonly ValidationIssue[]): string[] => [
    ...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "")),
];

const specimen = async (dir: string, name: string, bytes: Buffer): Promise<string> => {
    const file = path.join(dir, name);
    writeFileSync(file, bytes);
    return file;
};

const trailingBytes = async (): Promise<Buffer> => Buffer.concat([await buildDocx(), Buffer.from("trailing junk")]);

describe("a partial run applies its profile's policy (#381)", { timeout: 120_000 }, () => {
    it("lenient quick runs no Word rule, as full lenient does not", async () => {
        const result = await validate(ROW_WITHOUT_CELL, { profile: "lenient", quick: true });

        expect(result.valid).toBe(true);
        expect(result.issues.filter((issue) => issue.code === "word-row-without-cell")).toEqual([]);
        expect(result.metadata?.executedChecks).toEqual(["xml-syntax"]);
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

    it.each([[["word-row-without-cell"]], [["xml-syntax", "word-row-without-cell"]], [["package"]], [["tables"]]])(
        "refuses a lenient run of %j, which names checks lenient never runs",
        async (checks) => {
            const run = validate(ROW_WITHOUT_CELL, { profile: "lenient", checks });

            await expect(run).rejects.toBeInstanceOf(ModeConflictError);
            await expect(validate(ROW_WITHOUT_CELL, { profile: "lenient", checks })).rejects.toThrow(/lenient/);
        },
    );

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
                expect(result.metadata?.executedChecks).toEqual(["xml-syntax", "comment-markers", "comment-threading"]);
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
            ];
            for (const file of packages) {
                for (const profile of ["lenient", "word-valid"] as const) {
                    const label = `${path.basename(file)} under ${profile}`;
                    const full = errorCodes((await validate(file, { profile })).issues);
                    for (const selection of [{ quick: true }, { checks: ["comments"] }] as const) {
                        const partial = errorCodes((await validate(file, { profile, ...selection })).issues);
                        expect(full, `${label}: ${JSON.stringify(selection)}`).toEqual(expect.arrayContaining(partial));
                    }
                }
            }
        });
    });
});
