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

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
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
        expect(result.metadata?.executedChecks).toEqual(["comment-markers", "comment-threading"]);
        expect(result.metadata?.omittedChecks).toContain("xml-syntax");
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
            ];
            for (const file of packages) {
                for (const profile of ["lenient", "word-valid"] as const) {
                    const label = `${path.basename(file)} under ${profile}`;
                    const full = errorCodes((await validate(file, { profile })).issues);
                    const quick = errorCodes((await validate(file, { profile, quick: true })).issues);
                    expect(full, label).toEqual(expect.arrayContaining(quick));
                }
            }
        });
    });
});
