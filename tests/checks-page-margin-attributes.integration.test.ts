/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Integration tests for `pageMarginAttributesCheck` (issue #214, jubarte
 * lane family k10): a synthetic single-defect package through the full
 * `validate()` profiles, then the lane specimens through the named check
 * directly.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { pageMarginAttributesCheck } from "../src/scripts/office/checks/page-margin-attributes";
import { contextFromArchive } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

const CHECK_ID = "page-margin-attributes";

// The k10 defect: the live pgMar carries only top/right/bottom/left.
const DEFECTIVE_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace(
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`,
    `<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>`,
);

async function validateBytes(bytes: Uint8Array, profile: Profile) {
    return withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
}

describe("page-margin-attributes: synthetic single-defect package", { timeout: 120000 }, () => {
    it("word-valid stays valid and carries the advisory", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "word-valid");
        expect(result.valid).toBe(true);
        const findings = result.issues.filter((issue) => issue.code === CHECK_ID);
        expect(findings).toHaveLength(3);
        for (const finding of findings) {
            expect(finding).toMatchObject({ severity: "warning", path: "word/document.xml" });
        }
        expect(findings.map((finding) => finding.message)).toEqual([
            expect.stringContaining("w:header"),
            expect.stringContaining("w:footer"),
            expect.stringContaining("w:gutter"),
        ]);
    });

    it("strict rejects the package", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "strict");
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.code === CHECK_ID && issue.severity === "warning")).toBe(true);
    });

    it("lenient does not carry the check's code", async () => {
        const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": DEFECTIVE_DOCUMENT }), "lenient");
        expect(result.issues.map((issue) => issue.code)).not.toContain(CHECK_ID);
    });
});

describe("page-margin-attributes: jubarte 0.13.0 lane specimens", () => {
    const findingsByPath = async (fileName: string) => {
        const { context } = await contextFromArchive(readFileSync(path.join(LANE, fileName)));
        const issues = await pageMarginAttributesCheck.run(context);
        const byPath = new Map<string, number>();
        for (const issue of issues) {
            expect(issue).toMatchObject({ severity: "warning", code: CHECK_ID });
            byPath.set(issue.path ?? "", (byPath.get(issue.path ?? "") ?? 0) + 1);
        }
        return byPath;
    };

    it("k10_1.redline.docx: three missing attributes on the final body-level w:pgMar", async () => {
        const byPath = await findingsByPath("k10_1.redline.docx");
        // The pgMar at document.xml:919 lacks w:header, w:footer and w:gutter;
        // the sectPrChange copy is complete.
        expect(byPath).toEqual(new Map([["word/document.xml", 3]]));
    });

    it("k10_2.redline.docx: two missing attributes inside the w:sectPrChange w:sectPr", async () => {
        const byPath = await findingsByPath("k10_2.redline.docx");
        // The sectPrChange copy's pgMar lacks w:header and w:gutter; the live
        // pgMar is complete.
        expect(byPath).toEqual(new Map([["word/document.xml", 2]]));
    });
});
