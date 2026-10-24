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
import { defaultProfileForFormat } from "../src/lib/mode-policy";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

/** The profile a `.docx` is validated under when the caller names none. */
const DEFAULT_DOCX_PROFILE = defaultProfileForFormat("docx");

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

describe("validation contract: strict keeps checking after an XML failure", { timeout: 120000 }, () => {
    // A row with no cell in the body, and styles.xml cut off before its end tag.
    const ROW_WITHOUT_CELL =
        `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr/></w:tbl>` +
        `<w:p w14:paraId="00000002" w14:textId="00000002"/>`;

    it("strict reports every error code word-valid reports when another part is malformed", async () => {
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${ROW_WITHOUT_CELL}<w:sectPr>`),
            "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace("</w:styles>", ""),
        });
        const [wordValid, strict] = await Promise.all([validateBytes(bytes, "word-valid"), validateBytes(bytes, "strict")]);
        const wordValidCodes = errorCodes(wordValid.issues);
        expect(wordValidCodes.has("word-row-without-cell")).toBe(true);
        const strictCodes = errorCodes(strict.issues);
        expect([...wordValidCodes].filter((code) => !strictCodes.has(code))).toEqual([]);
        expect(strict.valid).toBe(false);
    });

    it("reports the malformed part once, through its XML diagnostics", async () => {
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace("</w:styles>", ""),
        });
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(bytes, profile);
            const onPart = result.issues.filter((issue) => issue.path === "word/styles.xml").map((issue) => issue.code);
            expect(new Set(onPart), profile).toEqual(new Set(["xml-syntax"]));
            expect(result.valid, profile).toBe(false);
        }
    });

    it("reports the malformed part's syntax error exactly once", async () => {
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace("</w:styles>", ""),
        });
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(bytes, profile);
            expect(
                result.issues.filter((issue) => issue.code === "xml-syntax"),
                profile,
            ).toHaveLength(1);
        }
    });

    it("keeps judging the well-formed parts when a part Word tolerates is malformed", async () => {
        // A run directly in a table cell makes Word refuse the file
        // (word-rejects-misplaced-run.docx); a malformed part outside word/
        // is not a Word rejection on its own.
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:sectPr>",
                `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr w14:paraId="00000003" w14:textId="00000003"><w:tc><w:r><w:t>x</w:t></w:r></w:tc></w:tr></w:tbl><w:p w14:paraId="00000004" w14:textId="00000004"/><w:sectPr>`,
            ),
            "customXml/item1.xml": "<root><unclosed></root>",
        });
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(bytes, profile);
            expect(result.issues, profile).toContainEqual(
                expect.objectContaining({ code: "word-run-misplaced", severity: "error", path: "word/document.xml" }),
            );
            expect(result.valid, profile).toBe(false);
        }
    });
});

describe("validation contract: Word rules judge the content Word consumes (#291)", { timeout: 120000 }, () => {
    // A cell-less row inside one AlternateContent branch, with a valid paragraph in the Fallback.
    const alternateRow = (requires: string) =>
        `<mc:AlternateContent xmlns:e="urn:future" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
        `<mc:Choice Requires="${requires}"><w:tbl><w:tblPr/><w:tblGrid/><w:tr/></w:tbl></mc:Choice>` +
        `<mc:Fallback><w:p w14:paraId="00000002" w14:textId="00000002"/></mc:Fallback>` +
        `</mc:AlternateContent>`;
    const withAlternate = (requires: string) =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${alternateRow(requires)}<w:sectPr>`),
        });

    it("word-valid ignores a defect inside a Choice Word does not select", async () => {
        const result = await validateBytes(await withAlternate("e"), "word-valid");
        expect([...errorCodes(result.issues)]).toEqual([]);
        expect(result.valid).toBe(true);
    });

    it("word-valid rejects the same defect inside a Choice Word selects", async () => {
        const result = await validateBytes(await withAlternate("wps"), "word-valid");
        expect(errorCodes(result.issues).has("word-row-without-cell")).toBe(true);
        expect(result.valid).toBe(false);
    });
});

describe("validation contract: default DOCX profile", { timeout: 120000 }, () => {
    // Flipped in Task 10: the default DOCX profile is `word-valid`, so a run
    // without a profile must agree with an explicit `word-valid` run.
    it("default DOCX validation returns the same result as an explicit word-valid run", async () => {
        const specimen = await validateBytes(await buildDocx(), DEFAULT_DOCX_PROFILE);
        const explicit = await validateBytes(await buildDocx(), "word-valid");
        expect(specimen.valid).toBe(explicit.valid);
        expect(specimen.metadata?.profile).toBe("word-valid");
        expect(specimen.issues.map((i) => [i.severity, i.code, i.path, i.message])).toEqual(
            explicit.issues.map((i) => [i.severity, i.code, i.path, i.message]),
        );
    });
});
