/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Validation contract (docs/plans/2026-10-23-fixing-validator.md, Task 1).
 *
 * `strict` is a superset of `word-valid`: every defect that makes Microsoft
 * Word refuse or repair a package is also a strict failure. Each case
 * compares the error codes of a defect specimen against its control, so a
 * pass means the named check fired, not that some unrelated schema or
 * fixture error happened to make the result invalid.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken");
const WORKING = path.join(HERE, "fixtures", "working");

const errorCodes = (issues: readonly { severity: string; code?: string }[]): Set<string> =>
    new Set(issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? ""));

async function validateBytes(bytes: Uint8Array, profile: Profile) {
    return withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
}

describe("validation contract: strict includes every Word-valid package check", { timeout: 120000 }, () => {
    for (const profile of ["word-valid", "strict"] as const) {
        it(`${profile}: the schema-clean control passes`, async () => {
            const result = await validateBytes(await buildDocx(), profile);
            expect([...errorCodes(result.issues)]).toEqual([]);
            expect(result.valid).toBe(true);
        });

        it(`${profile}: one undeclared byte after the ZIP EOCD record is rejected as zip-bytes-after-eocd`, async () => {
            const control = await buildDocx();
            const result = await validateBytes(Buffer.concat([control, Buffer.from([0x0a])]), profile);
            expect([...errorCodes(result.issues)]).toEqual(["zip-bytes-after-eocd"]);
            expect(result.valid).toBe(false);
        });

        it(`${profile}: an EOCD comment length past the end of the file is rejected as zip-eocd-comment-overrun`, async () => {
            const control = await buildDocx();
            const overrun = Buffer.from(control);
            overrun.writeUInt16LE(8, overrun.length - 2);
            const result = await validateBytes(overrun, profile);
            expect(errorCodes(result.issues).has("zip-eocd-comment-overrun")).toBe(true);
            expect(result.valid).toBe(false);
        });
    }
});

/**
 * Every Word-proven rejection rule, with the fixture pair its own suite uses:
 * the specimen Word refuses and the repaired control Word opens.
 */
const WORD_RULES: readonly { code: string; specimen: string; control: string }[] = [
    { code: "word-anchor-missing-attribute", specimen: "package.word-anchor-open.docx", control: "package.word-anchor-open-repaired.docx" },
    {
        code: "word-attached-toolbar-utf8",
        specimen: "package.word-attached-toolbar-utf8.docx",
        control: "package.word-attached-toolbar-utf8-repaired.docx",
    },
    {
        code: "word-cell-ends-with-table",
        specimen: "package.word-cell-ends-with-table.docx",
        control: "package.word-cell-ends-with-table-repaired.docx",
    },
    {
        code: "word-choice-requires-undeclared",
        specimen: "package.word-choice-requires-undeclared.docx",
        control: "package.word-choice-requires-declared.docx",
    },
    {
        code: "word-drawing-extent-negative",
        specimen: "package.word-drawing-extent-negative.docx",
        control: "package.word-drawing-extent-positive.docx",
    },
    {
        code: "word-duplicate-font-table-relationship",
        specimen: "package.word-duplicate-font-table.docx",
        control: "package.word-duplicate-font-table-removed.docx",
    },
    {
        code: "word-duplicate-settings-relationship",
        specimen: "package.word-duplicate-settings.docx",
        control: "package.word-duplicate-settings-removed.docx",
    },
    {
        code: "word-duplicate-theme-relationship",
        specimen: "package.word-duplicate-theme.docx",
        control: "package.word-duplicate-theme-removed.docx",
    },
    {
        code: "word-duplicate-web-settings-relationship",
        specimen: "package.word-duplicate-web-settings.docx",
        control: "package.word-duplicate-web-settings-removed.docx",
    },
    { code: "word-empty-drawing", specimen: "package.word-empty-drawing.docx", control: "package.word-empty-drawing-removed.docx" },
    {
        code: "word-footnote-reference-missing",
        specimen: "package.word-footnote-reference-missing.docx",
        control: "package.word-footnote-reference-present.docx",
    },
    {
        code: "word-hyperlink-relationship",
        specimen: "package.word-hyperlink-relationship.docx",
        control: "package.word-hyperlink-relationship-repaired.docx",
    },
    {
        code: "word-move-from-deleted-text",
        specimen: "package.word-move-from-deleted-text.docx",
        control: "package.word-move-from-text.docx",
    },
    {
        code: "word-numbering-picture-bullet",
        specimen: "package.word-numbering-picture-bullet.docx",
        control: "package.word-numbering-picture-bullet-removed.docx",
    },
    {
        code: "word-row-without-cell",
        specimen: "package.word-row-without-cell.docx",
        control: "package.word-row-without-cell-repaired.docx",
    },
    { code: "word-sdt-id-empty", specimen: "package.word-sdt-id-empty.docx", control: "package.word-sdt-id-empty-repaired.docx" },
    {
        code: "word-text-contains-paragraph",
        specimen: "package.word-text-contains-paragraph.docx",
        control: "package.word-text-contains-paragraph-lifted.docx",
    },
    {
        code: "word-textbox-deleted-text",
        specimen: "package.word-textbox-deleted-text.docx",
        control: "package.word-textbox-deleted-text-removed.docx",
    },
    { code: "word-drawing-arc-radius", specimen: "package.drawing-arc-radius.docx", control: "package.drawing-arc-radius-integer.docx" },
    { code: "ct-uncovered-part", specimen: "package.word-undeclared-part.docx", control: "package.word-undeclared-part-removed.docx" },
];

describe("validation contract: every Word rejection rule runs under strict", { timeout: 120000 }, () => {
    for (const rule of WORD_RULES) {
        it(`strict reports ${rule.code} on ${rule.specimen} and not on its control`, async () => {
            const [specimen, control] = await Promise.all([
                validate(path.join(BROKEN, rule.specimen), { profile: "strict" }),
                validate(path.join(WORKING, rule.control), { profile: "strict" }),
            ]);
            expect(errorCodes(specimen.issues).has(rule.code)).toBe(true);
            expect(errorCodes(control.issues).has(rule.code)).toBe(false);
            expect(specimen.valid).toBe(false);
        });
    }
});

describe("validation contract: default DOCX profile", () => {
    // Flipped on in Task 10, after the structural checks of Task 7 replace
    // the XSD-message heuristics. Until then the default stays `lenient`.
    it.todo("default DOCX validation returns the same result as an explicit word-valid run");
});
