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

import { existsSync, mkdirSync, promises as fs, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { defaultProfileForFormat } from "../src/lib/mode-policy";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { walkFiles } from "../src/scripts/office/validators/base";
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

        it(`${profile}: a schema-clean package behind a 4096-byte prefix is rejected as zip-bytes-before-archive`, async () => {
            const control = await buildDocx();
            const result = await validateBytes(Buffer.concat([Buffer.alloc(4096, 0x20), control]), profile);
            expect([...errorCodes(result.issues)]).toEqual(["zip-bytes-before-archive"]);
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

describe("validation contract: Word-refused names also fail strict", { timeout: 120000 }, () => {
    for (const code of ["opc-part-name-word-refused", "ct-override-partname-word-refused"]) {
        it(`${code} is an error under both profiles`, async () => {
            const parts = code.startsWith("opc-")
                ? { ...MINIMAL_PARTS, "word/a%21b.xml": "<x/>" }
                : {
                      ...MINIMAL_PARTS,
                      "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                          "</Types>",
                          '<Override PartName="/word/a%21b.xml" ContentType="application/xml"/></Types>',
                      ),
                  };
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(await buildDocx(parts), profile);
                expect(errorCodes(result.issues).has(code)).toBe(true);
                expect(result.valid).toBe(false);
            }
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

describe("validation contract: Markup Compatibility in the content-types stream (#459)", { timeout: 120000 }, () => {
    // Session 14a (docs/plans/evidence/word-probes-2026-10-24.md): Word 16.115
    // offers recovery for `mc:Ignorable` in [Content_Types].xml, so word-valid
    // must fail it and strict — the superset — must report the same check.
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const withIgnorable = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        /<Types ([^>]*)>/,
        `<Types $1 xmlns:mc="${MC}" xmlns:x="urn:x:unknown" mc:Ignorable="x">`,
    );

    for (const profile of ["word-valid", "strict"] as const) {
        it(`${profile}: mc:Ignorable in the content-types stream is rejected as ct-markup-compatibility`, async () => {
            const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "[Content_Types].xml": withIgnorable }), profile);
            expect(errorCodes(result.issues).has("ct-markup-compatibility")).toBe(true);
            expect(result.valid).toBe(false);
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

describe("validation contract: a full run judges the entries a quick run judges (#368)", { timeout: 120000 }, () => {
    it.each(["word-valid", "lenient", "strict"] as const)(
        "%s judges only the entry the package view keeps when two entries name one part",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "word\\document.xml": "<w:document" }));
                const quick = await validate(file, { checks: ["xml-syntax"] });
                const full = await validate(file, { profile });
                const syntax = (issues: readonly { code?: string; path?: string }[]): (string | undefined)[] =>
                    issues.filter((issue) => issue.code === "xml-syntax").map((issue) => issue.path);

                expect(syntax(quick.issues)).toEqual([]);
                expect(syntax(full.issues)).toEqual([]);
                expect(full.issues.filter((issue) => issue.path?.includes("\\"))).toEqual([]);
            });
        },
    );
});

describe("validation contract: a full run judges what the package view holds (#371, #372, #373)", { timeout: 120000 }, () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const withSpecimen = <T>(
        parts: Readonly<Record<string, string | Uint8Array>>,
        run: (file: string, dir: string) => Promise<T>,
    ): Promise<T> =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(parts));
            return run(file, dir);
        });
    const withoutMain = (): Record<string, string> => {
        const { "word/document.xml": _main, ...rest } = MINIMAL_PARTS;
        return rest;
    };

    it.each(["/word/bonus.xml", "\\word\\bonus.xml", "C:/bonus.xml"])(
        "a quick and a full run both refuse the item name %j, which places no part (#372)",
        async (name) => {
            await withSpecimen({ ...MINIMAL_PARTS, [name]: "<b/>" }, async (file) => {
                const quick = await validate(file, { checks: ["xml-syntax"] });
                const full = await validate(file, { profile: "lenient" });
                expect(errorCodes(quick.issues)).toContain("package-extract-failed");
                expect(errorCodes(full.issues)).toContain("package-extract-failed");
                expect([quick.valid, full.valid]).toEqual([false, false]);
            });
        },
    );

    it("schema-validates a main part stored only as word\\document.xml (#371)", async () => {
        const main = MINIMAL_PARTS["word/document.xml"].replace("<w:body>", "<w:body><w:bogus/>");
        await withSpecimen({ ...withoutMain(), "word\\document.xml": main }, async (file) => {
            const full = await validate(file, { profile: "strict" });
            expect(full.xsdWalk.parts).toContain("word/document.xml");
            expect(errorCodes(full.issues)).toContain("xsd-error");
        });
    });

    it("finds the sidecar of a part stored only as word\\header1.xml (#371)", async () => {
        const header =
            `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p><w:r><w:drawing>` +
            `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId9"/>` +
            `</w:drawing></w:r></w:p></w:hdr>`;
        const sidecar =
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId9" Type="${R}/image" Target="media/i.png"/></Relationships>`;
        await withSpecimen(
            {
                ...MINIMAL_PARTS,
                "word\\header1.xml": header,
                "word/_rels/header1.xml.rels": sidecar,
                "word/media/i.png": new Uint8Array([1, 2, 3]),
            },
            async (file) => {
                const full = await validate(file, { profile: "lenient" });
                expect(full.issues.filter((issue) => issue.code === "rels-missing-sidecar")).toEqual([]);
            },
        );
    });

    it.each([
        ["the malformed entry first", "cafe\u0301.xml", "caf\u00e9.xml"],
        ["the malformed entry last", "caf\u00e9.xml", "cafe\u0301.xml"],
    ])("never judges bytes the filesystem merged from two entries (%s, #373)", async (_order, malformed, wellFormed) => {
        // Both parts are referenced, so the malformed entry is the package's only defect.
        const rels = MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
            "</Relationships>",
            `<Relationship Id="rId2" Type="${R}/customXml" Target="../customXml/${malformed}"/>` +
                `<Relationship Id="rId3" Type="${R}/customXml" Target="../customXml/${wellFormed}"/></Relationships>`,
        );
        await withSpecimen(
            {
                ...MINIMAL_PARTS,
                "word/_rels/document.xml.rels": rels,
                [`customXml/${malformed}`]: "<broken",
                [`customXml/${wellFormed}`]: "<x/>",
            },
            async (file, dir) => {
                mkdirSync(path.join(dir, "probe"));
                writeFileSync(path.join(dir, "probe", malformed), "");
                const merges = existsSync(path.join(dir, "probe", wellFormed));
                const quick = await validate(file, { checks: ["xml-syntax"] });
                const full = await validate(file, { profile: "lenient" });
                expect(quick.issues.filter((issue) => issue.code === "xml-syntax").map((issue) => issue.path)).toEqual([
                    `customXml/${malformed}`,
                ]);
                expect(full.valid).toBe(false);
                expect(errorCodes(full.issues)).toContain(merges ? "package-extract-failed" : "xml-syntax");
            },
        );
    });
});

describe("validation contract: a full run keeps every part's own name (#376, #377, #378)", { timeout: 120000 }, () => {
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const withSpecimen = <T>(
        parts: Readonly<Record<string, string | Uint8Array>>,
        run: (file: string, dir: string) => Promise<T>,
    ): Promise<T> =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(parts));
            return run(file, dir);
        });

    it("never passes a toolbar it renamed through a folder the filesystem folds (#376)", async () => {
        const composed = "caf\u00e9";
        const decomposed = "cafe\u0301";
        const parts = {
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
                "</Types>",
                '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
            ),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
                "</Relationships>",
                `<Relationship Id="rId2" Type="${R}/settings" Target="settings.xml"/>` +
                    `<Relationship Id="rId3" Type="${R}/customXml" Target="../${composed}/other.bin"/></Relationships>`,
            ),
            "word/settings.xml":
                `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
                `xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:r="${R}"><wne:toolbarData r:id="rId1"/></w:settings>`,
            "word/_rels/settings.xml.rels":
                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId1" Type="${R}/attachedToolbars" Target="../${decomposed}/toolbars.bin"/></Relationships>`,
            [`${composed}/other.bin`]: new Uint8Array([1]),
            [`${decomposed}/toolbars.bin`]: new Uint8Array([0xef, 0xbf, 0xbd]),
        };
        await withSpecimen(parts, async (file, dir) => {
            mkdirSync(path.join(dir, "probe"));
            writeFileSync(path.join(dir, "probe", composed), "");
            const merges = existsSync(path.join(dir, "probe", decomposed));
            const quick = await validate(file, { checks: ["word-attached-toolbar-utf8"] });
            const full = await validate(file, { profile: "word-valid" });
            expect(errorCodes(quick.issues)).toContain("word-attached-toolbar-utf8");
            expect(full.valid).toBe(false);
            expect(errorCodes(full.issues)).toContain(merges ? "package-extract-failed" : "word-attached-toolbar-utf8");
        });
    });

    it("schema-validates a package whose parts sit under WORD/ (#377)", async () => {
        const upper = Object.fromEntries(
            Object.entries(MINIMAL_PARTS).map(([name, text]) => [
                name.replace(/^word\//, "WORD/"),
                text
                    .replaceAll('PartName="/word/', 'PartName="/WORD/')
                    .replaceAll('Target="word/', 'Target="WORD/')
                    .replace('w:w="12240"', 'w:w="BOGUS"'),
            ]),
        );
        await withSpecimen(upper, async (file) => {
            const full = await validate(file, { profile: "strict" });
            expect(full.xsdWalk.parts).toContain("WORD/document.xml");
            expect(errorCodes(full.issues)).toContain("xsd-error");
        });
    });

    it("applies markup compatibility to WORD/ parts before their schema (#377)", async () => {
        // The main part carries w14 attributes under mc:Ignorable="w14".
        const upper = Object.fromEntries(
            Object.entries(MINIMAL_PARTS).map(([name, text]) => [
                name.replace(/^word\//, "WORD/"),
                text.replaceAll('PartName="/word/', 'PartName="/WORD/').replaceAll('Target="word/', 'Target="WORD/'),
            ]),
        );
        expect(upper["WORD/document.xml"]).toContain('mc:Ignorable="w14"');
        await withSpecimen(upper, async (file) => {
            const full = await validate(file, { profile: "strict" });
            expect(full.xsdWalk.parts).toContain("WORD/document.xml");
            expect(full.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
        });
    });

    it.each([
        ["a main entry stored as /../word/document.xml", "/../word/document.xml", "word/document.xml"],
        ["custom properties stored as docProps\\nested\\..\\custom.xml", "docProps\\nested\\..\\custom.xml", undefined],
    ])("a quick and a full run both refuse %s (#378)", async (_label, stored, replaces) => {
        const parts: Record<string, string> = { ...MINIMAL_PARTS };
        if (replaces !== undefined) {
            delete parts[replaces];
        }
        parts[stored] = replaces === undefined ? "<x/>" : MINIMAL_PARTS[replaces];
        await withSpecimen(parts, async (file) => {
            const quick = await validate(file, { checks: ["xml-syntax"] });
            const full = await validate(file, { profile: "lenient" });
            expect(errorCodes(quick.issues)).toContain("package-extract-failed");
            expect(errorCodes(full.issues)).toContain("package-extract-failed");
        });
    });
});

describe("validation contract: a Word-read part typed only by the generic Default (#432)", { timeout: 120000 }, () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const CT_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    /** The five parts Word 16.115 offers to recover when only the xml Default types them (session 8a). */
    const GENERIC_DEFAULT_PARTS: readonly (readonly [string, string, string])[] = [
        ["word/document.xml", "officeDocument", CT_MAIN],
        ["word/styles.xml", "styles", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"],
        ["word/settings.xml", "settings", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"],
        ["word/fontTable.xml", "fontTable", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"],
        ["word/webSettings.xml", "webSettings", "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"],
    ];
    const ROOTS: Readonly<Record<string, string>> = {
        "word/settings.xml": "w:settings",
        "word/fontTable.xml": "w:fonts",
        "word/webSettings.xml": "w:webSettings",
    };

    const specimen = async (part: string, segment: string, contentType: string, override: boolean): Promise<Uint8Array> => {
        const extra: Record<string, string> =
            ROOTS[part] === undefined
                ? {}
                : { [part]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${ROOTS[part]} xmlns:w="${W}"/>` };
        const rels =
            part === "word/document.xml"
                ? MINIMAL_PARTS["word/_rels/document.xml.rels"]
                : MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                      "</Relationships>",
                      `<Relationship Id="rId9" Type="${R}/${segment}" Target="${part.replace("word/", "")}"/></Relationships>`,
                  );
        const stream = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            new RegExp(`<Override PartName="/${part.replaceAll(".", "\\.")}" ContentType="[^"]+"/>`),
            "",
        ).replace("</Types>", override ? `<Override PartName="/${part}" ContentType="${contentType}"/></Types>` : "</Types>");
        return buildDocx({ ...MINIMAL_PARTS, ...extra, "word/_rels/document.xml.rels": rels, "[Content_Types].xml": stream });
    };

    it.each(GENERIC_DEFAULT_PARTS)(
        "word-valid and strict both fail %s typed only by the generic Default",
        { timeout: 120000 },
        async (part, segment, contentType) => {
            const [wordValid, strict] = await Promise.all([
                validateBytes(await specimen(part, segment, contentType, false), "word-valid"),
                validateBytes(await specimen(part, segment, contentType, false), "strict"),
            ]);
            expect(wordValid.valid).toBe(false);
            expect(strict.valid).toBe(false);
            expect(wordValid.issues).toContainEqual(
                expect.objectContaining({ code: "word-part-content-type-generic", severity: "error", path: part }),
            );
            if (segment === "officeDocument") {
                // Strict already rejects the main document through its
                // declarable root; the defect is not reported twice.
                expect(strict.issues.some((issue) => issue.code === "word-part-content-type-generic")).toBe(false);
                expect(errorCodes(strict.issues).has("ct-undeclared-part")).toBe(true);
            } else {
                expect(strict.issues).toContainEqual(
                    expect.objectContaining({ code: "word-part-content-type-generic", severity: "error", path: part }),
                );
            }
        },
    );

    it.each(GENERIC_DEFAULT_PARTS)(
        "every profile passes %s with its proper Override",
        { timeout: 120000 },
        async (part, segment, contentType) => {
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(await specimen(part, segment, contentType, true), profile);
                expect(errorCodes(result.issues).has("word-part-content-type-generic"), `${part} ${profile}`).toBe(false);
                expect(result.valid, `${part} ${profile}`).toBe(true);
            }
        },
    );
});

// These are filesystem integration tests: literal backslashes are file-name
// characters on POSIX, and separators on Windows.
describe.skipIf(process.platform === "win32")("directory package staging (#374)", { timeout: 120000 }, () => {
    type Parts = Readonly<Record<string, string | Uint8Array>>;
    const storedName = (name: string) => name.replaceAll("/", "\\");
    const writeDirectory = async (root: string, parts: Parts, backslashes: boolean) => {
        await fs.mkdir(root, { recursive: true });
        for (const [part, bytes] of Object.entries(parts)) {
            const file = path.join(root, backslashes ? storedName(part) : part);
            await fs.mkdir(path.dirname(file), { recursive: true });
            await fs.writeFile(file, bytes);
        }
        return root;
    };
    const snapshot = async (root: string) => {
        const files: Record<string, Buffer> = {};
        for (const file of walkFiles(root)) files[path.relative(root, file)] = await fs.readFile(file);
        return files;
    };
    const findings = (result: Awaited<ReturnType<typeof validate>>) => result.issues.map((issue) => JSON.stringify(issue)).sort();
    // validate() must resolve, never reject (#412): a rejection fails here as
    // an assertion, naming the error, instead of escaping the test.
    const validateResolves = async (...args: Parameters<typeof validate>) => {
        const settled = await validate(...args).then(
            (result) => ({ result, error: undefined }),
            (error: unknown) => ({ result: undefined, error: error instanceof Error ? error.message : String(error) }),
        );
        expect(settled.error, "validate() rejected instead of resolving").toBeUndefined();
        return settled.result!;
    };
    const assertSameVerdict = (actual: Awaited<ReturnType<typeof validate>>, control: Awaited<ReturnType<typeof validate>>) => {
        expect(actual.valid).toBe(control.valid);
        expect(findings(actual)).toEqual(findings(control));
        expect(actual.metadata?.schemaCompleted).toBe(control.metadata?.schemaCompleted);
    };

    it.each([false, true])("schema-walks a backslash main part (invalid=%s), preserving every input byte", async (invalid) => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": invalid
                    ? MINIMAL_PARTS["word/document.xml"].replace("<w:body>", "<w:body><w:bogus/>")
                    : MINIMAL_PARTS["word/document.xml"],
            };
            const controlDir = await writeDirectory(path.join(dir, "control.docx"), parts, false);
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const before = await snapshot(input);
            const control = await validate(controlDir, { strict: true });
            const actual = await validate(input, { strict: true });
            expect(actual.xsdWalk.parts).toContain("word/document.xml");
            expect(actual.xsdWalk).toEqual(control.xsdWalk);
            expect(errorCodes(actual.issues).has("xsd-error")).toBe(invalid);
            assertSameVerdict(actual, control);
            expect(await snapshot(input)).toEqual(before);
        });
    });

    it("finds backslash sidecars and image targets with the same findings as a normal directory", async () => {
        await withTempDir(async (dir) => {
            const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const parts = {
                ...MINIMAL_PARTS,
                "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
                    "</Types>",
                    '<Default Extension="png" ContentType="image/png"/></Types>',
                ),
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(
                    "<w:r><w:t>Hello</w:t></w:r>",
                    `<w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${r}" r:embed="rId9"/></w:drawing></w:r>`,
                ),
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
                    "</Relationships>",
                    `<Relationship Id="rId9" Type="${r}/image" Target="media/image1.png"/></Relationships>`,
                ),
                "word/media/image1.png": new Uint8Array([1, 2, 3]),
            };
            const control = await validate(await writeDirectory(path.join(dir, "control.docx"), parts, false));
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const before = await snapshot(input);
            const actual = await validate(input);
            expect(actual.issues.filter((i) => i.code === "rels-missing-sidecar")).toEqual([]);
            assertSameVerdict(actual, control);
            expect(await snapshot(input)).toEqual(before);
        });
    });

    it("judges only the kept duplicate, as the packed package does, and reports the duplicate (#351)", async () => {
        await withTempDir(async (dir) => {
            const parts = { ...MINIMAL_PARTS, "word\\document.xml": "<dropped" };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, false);
            const before = await snapshot(input);
            const packed = path.join(dir, "packed.docx");
            await fs.writeFile(packed, await buildDocx(parts));
            const spy = vi.spyOn(DOCXSchemaValidator.prototype, "validate");
            let actual;
            try {
                actual = await validate(input, { strict: true });
                // Even a dropped backslash entry selects staging. The original
                // directory must not be passed to the path-keyed validators.
                expect(spy).toHaveBeenCalledOnce();
                expect(spy.mock.contexts[0]).not.toHaveProperty("unpackedDir", input);
            } finally {
                spy.mockRestore();
            }
            const archive = await validate(packed, { strict: true });
            const control = await validate(await writeDirectory(path.join(dir, "control.docx"), MINIMAL_PARTS, false), { strict: true });
            expect(control.valid).toBe(true);
            assertSameVerdict(actual, archive);
            // Strict rejects the duplicate itself; every other finding is the
            // forward-slash control's, so the dropped "<dropped" entry is not judged.
            expect(actual.valid).toBe(false);
            expect(actual.issues.filter((i) => i.code === "opc-equivalent-part-names")).toEqual([
                expect.objectContaining({ severity: "error", path: "word/document.xml" }),
            ]);
            const withoutDuplicate = (result: typeof actual) => findings(result).filter((f) => !f.includes('"opc-equivalent-part-names"'));
            expect(withoutDuplicate(actual)).toEqual(findings(control));
            expect(await snapshot(input)).toEqual(before);
        });
    });

    it("writes repaired bytes and deletions back to stored names, leaving every untouched file identical", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
                "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                "word/media/untouched.bin": new Uint8Array([0, 255, 1, 2]),
            };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const normal = await writeDirectory(path.join(dir, "control.docx"), parts, false);
            const before = await snapshot(input);
            const control = await validate(normal, { strict: true, autoRepair: true });
            const actual = await validate(input, { strict: true, autoRepair: true });
            expect(actual.xsdWalk.parts).toContain("word/document.xml");
            expect(actual.repairs).toBe(control.repairs);
            expect(actual.repairs).toBe(2);
            assertSameVerdict(actual, control);
            const after = await snapshot(input);
            const repaired = storedName("word/document.xml");
            const deleted = storedName("word/_rels/unused.xml.rels");
            expect(after[repaired]?.toString()).toContain('xml:space="preserve"');
            expect(after[repaired]).toEqual(await fs.readFile(path.join(normal, "word/document.xml")));
            expect(after).not.toHaveProperty(deleted);
            expect(existsSync(path.join(input, "word/document.xml"))).toBe(false);
            for (const [name, bytes] of Object.entries(before)) {
                if (name !== repaired && name !== deleted) expect(after[name]).toEqual(bytes);
            }
            expect(Object.keys(after).sort()).toEqual(
                Object.keys(before)
                    .filter((name) => name !== deleted)
                    .sort(),
            );
        });
    });

    it("does not stage a normal directory, and stages only the backslash control", async () => {
        await withTempDir(async (dir) => {
            const normal = await writeDirectory(path.join(dir, "normal.docx"), MINIMAL_PARTS, false);
            const backslash = await writeDirectory(path.join(dir, "backslash.docx"), MINIMAL_PARTS, true);
            // Observe the real validator's directory without replacing its behavior.
            const spy = vi.spyOn(DOCXSchemaValidator.prototype, "validate");
            try {
                const control = await validate(normal);
                const actual = await validate(backslash);
                expect(spy).toHaveBeenCalledTimes(2);
                expect(spy.mock.contexts[0]).toHaveProperty("unpackedDir", normal);
                expect(spy.mock.contexts[1]).not.toHaveProperty("unpackedDir", backslash);
                assertSameVerdict(actual, control);
            } finally {
                spy.mockRestore();
            }
        });
    });

    it.each([false, true])("keeps full/partial verdict and error-subset invariants (malformed=%s, #381)", async (malformed) => {
        await withTempDir(async (dir) => {
            const parts = { ...MINIMAL_PARTS, "word/document.xml": malformed ? "<w:document" : MINIMAL_PARTS["word/document.xml"] };
            const normal = await writeDirectory(path.join(dir, "normal.docx"), parts, false);
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const full = await validate(input);
            const control = await validate(normal);
            assertSameVerdict(full, control);
            expect(full.valid).toBe(!malformed);
            const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path}`));
            for (const selection of [{ checks: ["xml-syntax", "relationships", "tables", "drawings"] }, { quick: true }]) {
                const partial = await validate(input, { profile: "word-valid", ...selection });
                const partialControl = await validate(normal, { profile: "word-valid", ...selection });
                assertSameVerdict(partial, partialControl);
                if (full.valid) expect(partial.valid).toBe(true);
                for (const error of errors(partial)) expect(errors(full)).toContain(error);
            }
        });
    });

    it("stages a backslash PPTX slide and finds its sidecar as the normal directory does", async () => {
        await withTempDir(async (dir) => {
            const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            const parts = {
                "[Content_Types].xml":
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/></Types>',
                "ppt/presentation.xml": `<p:presentation xmlns:p="${p}"><p:bogus/></p:presentation>`,
                "ppt/slides/slide1.xml": `<p:sld xmlns:p="${p}" xmlns:r="${r}"><p:cSld><p:spTree><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId9"/></p:spTree></p:cSld></p:sld>`,
                "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${r}/image" Target="../media/image1.png"/></Relationships>`,
                "ppt/media/image1.png": new Uint8Array([1, 2, 3]),
            };
            const control = await validate(await writeDirectory(path.join(dir, "control.pptx"), parts, false));
            const input = await writeDirectory(path.join(dir, "input.pptx"), parts, true);
            const before = await snapshot(input);
            const actual = await validate(input);
            expect(actual.issues.filter((i) => i.code === "rels-missing-sidecar")).toEqual([]);
            expect(actual.xsdWalk.parts).toContain("ppt/presentation.xml");
            expect(actual.xsdWalk).toEqual(control.xsdWalk);
            expect(errorCodes(actual.issues)).toContain("xsd-error");
            assertSameVerdict(actual, control);
            expect(await snapshot(input)).toEqual(before);
        });
    });

    it.each(["word\\..\\escape.xml", "\\word\\bonus.xml"])(
        "refuses unplaceable stored name %j without changing the source",
        async (name) => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), { ...MINIMAL_PARTS, [name]: "<b/>" }, false);
                const before = await snapshot(input);
                const result = await validate(input, { autoRepair: true });
                expect(errorCodes(result.issues)).toContain("package-extract-failed");
                expect(result.metadata?.schemaCompleted).toBe(false);
                expect(result.repairs).toBe(0);
                expect(await snapshot(input)).toEqual(before);
            });
        },
    );

    // chmod has no effect for root, and the whole describe is skipped on win32.
    const runsAsRoot = process.getuid?.() === 0;

    it.skipIf(runsAsRoot)(
        "returns package-extract-failed for a stored file that cannot be read, where the direct path only warns (#412)",
        async () => {
            await withTempDir(async (dir) => {
                const parts: Parts = { ...MINIMAL_PARTS, "word/media/image1.png": new Uint8Array([1, 2, 3]) };
                const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
                const unreadable = path.join(input, storedName("word/media/image1.png"));
                await fs.chmod(unreadable, 0o000);
                try {
                    const actual = await validateResolves(input, { strict: true });
                    expect(actual.valid).toBe(false);
                    expect(actual.issues).toEqual([
                        expect.objectContaining({
                            severity: "error",
                            code: "package-extract-failed",
                            message: expect.stringContaining("image1.png"),
                        }),
                    ]);
                    expect(actual.xsdWalk.started).toBe(false);
                    expect(actual.metadata?.schemaCompleted).toBe(false);
                    expect(actual.repairs).toBe(0);
                } finally {
                    await fs.chmod(unreadable, 0o644);
                }
                // The same unreadable file under slash names takes the direct
                // path: validators read it lazily, so each failed read becomes
                // a check-crashed warning and the run stays valid. The staged
                // path must read the file up front, so it fails the run
                // instead — it cannot name which parts a package view holds.
                const direct = await writeDirectory(path.join(dir, "direct.docx"), parts, false);
                const directUnreadable = path.join(direct, "word", "media", "image1.png");
                await fs.chmod(directUnreadable, 0o000);
                try {
                    const control = await validate(direct);
                    expect(control.valid).toBe(true);
                    expect(control.issues.some((issue) => issue.code === "check-crashed")).toBe(true);
                } finally {
                    await fs.chmod(directUnreadable, 0o644);
                }
            });
        },
    );

    it.skipIf(runsAsRoot)(
        "reports a failed auto-repair write-back, keeps repairing the remaining files, and leaves the read-only source untouched (#412)",
        async () => {
            await withTempDir(async (dir) => {
                const parts = {
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
                    "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                };
                const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
                const normal = await writeDirectory(path.join(dir, "control.docx"), parts, false);
                const readOnly = path.join(input, storedName("word/document.xml"));
                const before = await snapshot(input);
                await fs.chmod(readOnly, 0o444);
                let actual;
                try {
                    actual = await validateResolves(input, { strict: true, autoRepair: true });
                } finally {
                    await fs.chmod(readOnly, 0o644);
                }
                // The control's write-back succeeds; `repairs` counts repairs
                // applied to the package view, on either path.
                const control = await validate(normal, { strict: true, autoRepair: true });
                expect(actual.repairs).toBe(control.repairs);
                expect(actual.valid).toBe(false);
                expect(actual.issues).toEqual(
                    expect.arrayContaining([
                        expect.objectContaining({
                            severity: "error",
                            code: "repair-write-back-failed",
                            path: "word/document.xml",
                            message: expect.stringContaining(storedName("word/document.xml")),
                        }),
                    ]),
                );
                const after = await snapshot(input);
                // The failed write-back left the stored bytes unfixed.
                expect(after[storedName("word/document.xml")]).toEqual(before[storedName("word/document.xml")]);
                // The other repair — deleting the empty .rels — still landed.
                expect(after).not.toHaveProperty(storedName("word/_rels/unused.xml.rels"));
                for (const [name, bytes] of Object.entries(before)) {
                    if (name !== storedName("word/_rels/unused.xml.rels")) expect(after[name]).toEqual(bytes);
                }
            });
        },
    );

    it("treats an already-removed write-back target as done when the repair deleted the part (#412)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const realRm = fs.rm.bind(fs);
            // The write-back removes the stored file for the deleted part; a
            // removal that already happened mid-run makes the second rm throw
            // ENOENT. Only the call into the input directory is doubled — the
            // repair's own rm lands in the temp copy.
            const spy = vi.spyOn(fs, "rm").mockImplementation(async (target, options) => {
                await realRm(target, options);
                if (typeof target === "string" && target.startsWith(input)) await realRm(target, options);
            });
            let actual;
            try {
                actual = await validateResolves(input, { strict: true, autoRepair: true });
            } finally {
                spy.mockRestore();
            }
            expect(actual.repairs).toBe(1);
            expect(actual.issues.filter((issue) => issue.code === "repair-write-back-failed")).toEqual([]);
            expect(existsSync(path.join(input, storedName("word/_rels/unused.xml.rels")))).toBe(false);
        });
    });

    it("reports a write-back removal that fails for a reason other than the source already being gone (#412)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const realRm = fs.rm.bind(fs);
            const spy = vi.spyOn(fs, "rm").mockImplementation(async (target, options) => {
                if (typeof target === "string" && target.startsWith(input)) {
                    throw Object.assign(new Error("EACCES: permission denied, rm"), { code: "EACCES" });
                }
                return realRm(target, options);
            });
            let actual;
            try {
                actual = await validateResolves(input, { strict: true, autoRepair: true });
            } finally {
                spy.mockRestore();
            }
            expect(actual.valid).toBe(false);
            expect(actual.issues).toEqual(
                expect.arrayContaining([
                    expect.objectContaining({
                        severity: "error",
                        code: "repair-write-back-failed",
                        path: "word/_rels/unused.xml.rels",
                    }),
                ]),
            );
            // The staged part was still removed — only the stored file stayed.
            expect(existsSync(path.join(input, storedName("word/_rels/unused.xml.rels")))).toBe(true);
        });
    });

    it("skips the write-back when the staged copy cannot be snapshotted before repair (#412)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
            };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const realRead = fs.readFile;
            let failed = false;
            // The first read of the staged word/document.xml is the pre-repair
            // snapshot in runStaged: the input's stored name uses a backslash
            // and the temp copy holds only forward-slash part names.
            const spy = vi.spyOn(fs, "readFile").mockImplementation(((
                target: Parameters<typeof fs.readFile>[0],
                ...rest: readonly unknown[]
            ) => {
                if (!failed && typeof target === "string" && target.endsWith("word/document.xml")) {
                    failed = true;
                    return Promise.reject(Object.assign(new Error("EPERM: staged read failed"), { code: "EPERM" }));
                }
                return (realRead as (...args: readonly unknown[]) => unknown)(target, ...rest);
            }) as typeof fs.readFile);
            let actual;
            try {
                actual = await validateResolves(input, { strict: true, autoRepair: true });
            } finally {
                spy.mockRestore();
            }
            expect(failed).toBe(true);
            expect(actual.valid).toBe(false);
            expect(actual.issues).toEqual(
                expect.arrayContaining([
                    expect.objectContaining({ severity: "error", code: "repair-write-back-failed", path: "word/document.xml" }),
                ]),
            );
            // Without the pre-repair bytes the write-back cannot prove what
            // changed, so the source is left untouched.
            const back = await fs.readFile(path.join(input, storedName("word/document.xml")), "utf-8");
            expect(back).toContain("> Hello <");
            expect(back).not.toContain('xml:space="preserve"');
        });
    });

    it.skipIf(runsAsRoot)("reports a staged part that turns unreadable before its write-back (#412)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
            };
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, true);
            const realRepair = DOCXSchemaValidator.prototype.repair;
            // Deterministic mid-run hook: once auto-repair has rewritten the
            // staged copy, the staged document turns unreadable, so the
            // write-back re-read fails with EACCES instead of ENOENT.
            const spy = vi.spyOn(DOCXSchemaValidator.prototype, "repair").mockImplementation(async function (this: DOCXSchemaValidator) {
                const repairs = await realRepair.call(this);
                await fs.chmod(path.join(this.unpackedDir, "word", "document.xml"), 0o000);
                return repairs;
            });
            let actual;
            try {
                actual = await validateResolves(input, { strict: true, autoRepair: true });
            } finally {
                spy.mockRestore();
            }
            expect(actual.valid).toBe(false);
            expect(actual.issues).toEqual(
                expect.arrayContaining([
                    expect.objectContaining({ severity: "error", code: "repair-write-back-failed", path: "word/document.xml" }),
                ]),
            );
            expect(actual.repairs).toBe(1);
            const back = await fs.readFile(path.join(input, storedName("word/document.xml")), "utf-8");
            expect(back).toContain("> Hello <");
            expect(back).not.toContain('xml:space="preserve"');
        });
    });
});
