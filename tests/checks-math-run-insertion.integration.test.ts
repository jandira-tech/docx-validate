/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { mathRunInsertionCheck } from "../src/scripts/office/checks/math-run-insertion";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const HERE = path.dirname(fileURLToPath(import.meta.url));
const validateBytes = async (bytes: Uint8Array, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        const result = await validate(file, { profile });
        expect(readFileSync(file)).toEqual(Buffer.from(bytes));
        return result;
    });
const packageWithMath = (math: string) =>
    buildDocx({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("</w:p>", `<m:oMath xmlns:m="${M}">${math}</m:oMath></w:p>`),
    });

describe("math-run-insertion integration", { timeout: 120000 }, () => {
    it("keeps a single-defect package advisory under Word-valid, rejects it under strict, and leaves lenient unchanged", async () => {
        const bytes = await packageWithMath(`<m:r><w:ins w:id="1" w:author="A"><m:t>x</m:t></w:ins></m:r>`);
        const word = await validateBytes(bytes, "word-valid");
        expect(word.valid).toBe(true);
        expect(word.issues.filter((issue) => issue.code === mathRunInsertionCheck.id)).toEqual([
            expect.objectContaining({ severity: "warning", path: "word/document.xml" }),
        ]);
        const strict = await validateBytes(bytes, "strict");
        expect(strict.valid).toBe(false);
        expect(strict.issues.some((issue) => issue.severity === "error" && issue.code?.startsWith("xsd-"))).toBe(true);
        expect(strict.issues.filter((issue) => issue.code === mathRunInsertionCheck.id)).toHaveLength(1);
        const lenient = await validateBytes(bytes, "lenient");
        expect(lenient.issues.some((issue) => issue.code === mathRunInsertionCheck.id)).toBe(false);
    });

    it("proves the nearby math revision wrappers are schema-valid", async () => {
        const mark = `w:id="1" w:author="A"`;
        const run = `<m:r><w:rPr/><m:t>x</m:t></m:r>`;
        const bytes = await packageWithMath(`<w:ins ${mark}>${run}</w:ins><w:del ${mark}>${run}</w:del>
            <m:sSup><m:sSupPr><m:ctrlPr><w:ins ${mark}><w:del ${mark}><w:rPr/></w:del></w:ins></m:ctrlPr></m:sSupPr>
            <m:e><w:ins ${mark}>${run}</w:ins><m:ctrlPr><w:del ${mark}><w:rPr/></w:del></m:ctrlPr></m:e>
            <m:sup><w:del ${mark}>${run}</w:del></m:sup></m:sSup>`);
        const result = await validateBytes(bytes, "strict");
        expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        expect(result.valid).toBe(true);
        expect(result.issues.some((issue) => issue.code === mathRunInsertionCheck.id)).toBe(false);
    });

    // Counts independently recounted from all XML parts before implementing the check.
    for (const specimen of ["k02_1", "k02_2", "k03_1", "k03_2", "k05_1", "k05_2"]) {
        it(`${specimen}: reports exactly 11 document insertions`, async () => {
            const bytes = readFileSync(path.join(HERE, "lane", "jubarte-0.13.0", `${specimen}.redline.docx`));
            const { context, openError } = await contextFromArchive(bytes);
            expect(openError).toBeUndefined();
            const issues = await mathRunInsertionCheck.run(context);
            expect(issues).toHaveLength(11);
            expect(issues.map((issue) => [issue.path, issue.code, issue.severity])).toEqual(
                Array.from({ length: 11 }, () => ["word/document.xml", "word-math-run-insertion", "warning"]),
            );
        });
    }
});
