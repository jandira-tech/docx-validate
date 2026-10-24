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

import { execFile } from "node:child_process";
import { chmodSync, existsSync, mkdirSync, promises as fs, symlinkSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { defaultProfileForFormat } from "../src/lib/mode-policy";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { PPTXSchemaValidator } from "../src/scripts/office/validators/pptx";
import { walkFiles } from "../src/scripts/office/validators/base";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";
import { DUPLICATE_MAIN_DOCUMENT_PARTS } from "./support/main-document-specimens";

import { GENERIC_NUMBERING_THEME_SPECIMENS, genericNumberingThemeParts } from "./support/generic-numbering-theme-specimens";

import { BYTE_REJECTED_MAIN_PARTS, GENERIC_RENAMED_MAIN_PARTS } from "./support/word-profile-regression-specimens";

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

it.skipIf(process.platform === "win32" || process.getuid?.() === 0)(
    "validation contract: inaccessible targets throw the same system error across profiles and partial/full runs (#574, #381)",
    async () => {
        await withTempDir(async (dir) => {
            const holder = path.join(dir, "holder");
            const target = path.join(holder, "packed.docx");
            mkdirSync(holder);
            writeFileSync(target, await buildDocx());
            chmodSync(holder, 0o000);
            try {
                let systemError: unknown;
                try {
                    await fs.access(target);
                } catch (err) {
                    systemError = err;
                }
                expect(systemError).toMatchObject({ code: "EACCES" });
                const expected = `Error: Cannot access ${target}: ${(systemError as Error).message}`;
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    await expect(validate(target, { profile })).rejects.toThrow(expected);
                    if (profile === "strict") continue;
                    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }]) {
                        await expect(validate(target, { profile, ...selection })).rejects.toThrow(expected);
                    }
                }
            } finally {
                chmodSync(holder, 0o755);
            }
        });
    },
    60_000,
);

it.skipIf(process.platform === "win32" || process.getuid?.() === 0)(
    "validation contract: unreadable input keeps strict and partial/full messages aligned (#547, #381)",
    async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unreadable.docx");
            writeFileSync(file, await buildDocx());
            chmodSync(file, 0o000);
            try {
                const word = await validate(file, { profile: "word-valid" });
                expect(word.issues).toEqual([
                    expect.objectContaining({
                        code: "package-open-failed",
                        severity: "error",
                        message: expect.stringMatching(/^Cannot read /),
                    }),
                ]);
                for (const options of [{ strict: true }, { quick: true }, { checks: ["xml-syntax"] }]) {
                    const result = await validate(file, options);
                    expect(result.valid).toBe(false);
                    expect(result.issues).toEqual(word.issues);
                }
            } finally {
                chmodSync(file, 0o644);
            }
        });
    },
    60_000,
);

describe("validation contract: multiple main relationships (#515, #381)", () => {
    it.each(DUPLICATE_MAIN_DOCUMENT_PARTS)(
        "$label fails word-valid and strict, with partial/full agreement",
        async ({ parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                const control = path.join(dir, "control.docx");
                writeFileSync(file, await buildDocx(parts));
                writeFileSync(control, await buildDocx());
                const code = "word-main-document-part";
                for (const profile of ["word-valid", "strict"] as const) {
                    const full = await validate(file, { profile });
                    expect(full.valid).toBe(false);
                    expect(errorCodes(full.issues)).toEqual(new Set([code]));
                    expect((await validate(control, { profile })).valid).toBe(true);
                    if (profile === "strict") continue; // Strict rejects partial selectors.
                    for (const selection of [{ quick: true }, { checks: [code] }, { checks: ["relationships"] }]) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(partial.valid).toBe(full.valid);
                        expect(partial.issues.filter((issue) => issue.code === code)).toEqual(
                            full.issues.filter((issue) => issue.code === code),
                        );
                        expect(errorCodes(partial.issues)).toEqual(errorCodes(full.issues));
                        expect((await validate(control, { profile, ...selection })).valid).toBe(true);
                    }
                }
            });
        },
        60_000,
    );
});

describe("validation contract: no prefix claim where JSZip's local read throws (#524, #381)", { timeout: 120_000 }, () => {
    // Review AP A2: the first record's local header points at a local
    // signature appended as the EOCD comment. JSZip reads the signature and
    // then past the buffer (zipEntry.js `readLocalPart`), so the archive
    // never opens and `zip-bytes-before-archive` claims no prefix for it.
    it("word-valid and strict report the refusal alone, and partial runs agree", async () => {
        const bytes = buildRawZip(Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })));
        const cd = bytes.readUInt32LE(bytes.length - 22 + 16);
        bytes.writeUInt32LE(bytes.length, cd + 42);
        bytes.writeUInt16LE(4, bytes.length - 2);
        const specimen = Buffer.concat([bytes, Buffer.from([0x50, 0x4b, 0x03, 0x04])]);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, specimen);
            for (const profile of ["word-valid", "strict"] as const) {
                const full = await validate(file, { profile });
                expect(full.valid, profile).toBe(false);
                expect([...errorCodes(full.issues)], profile).toEqual(["package-open-failed"]);
                if (profile === "strict") continue; // Strict rejects partial selectors.
                for (const selection of [{ quick: true }, { checks: ["zip-bytes-before-archive"] }, { checks: ["package"] }]) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    expect(partial.valid, tag).toBe(false);
                    expect(errorCodes(partial.issues).has("package-open-failed"), tag).toBe(true);
                    expect(
                        [...errorCodes(partial.issues)].every((code) => errorCodes(full.issues).has(code)),
                        `${tag}: partial errors ⊆ full errors`,
                    ).toBe(true);
                }
            }
        });
    });
});

describe("validation contract: no prefix claim where JSZip's reader index goes negative (#555, #381)", { timeout: 120_000 }, () => {
    // Review AT F1: a central compressed size JSZip reads as negative — the
    // 32-bit field 0x80000000, or a ZIP64 0x0001 extra resolving below zero
    // — makes `readLocalPart`'s final `readData(compressedSize)` throw on
    // `checkIndex`'s `newIndex < 0` arm, so the archive never opens and no
    // prefix is claimed beside the refusal, full and partial runs alike.
    it.each([["the 32-bit compressed-size field 0x80000000"], ["a ZIP64 extra resolving the compressed size below zero"]])(
        "word-valid and strict report the refusal alone for %s, and partial runs agree",
        async ([label]) => {
            const plain = buildRawZip(Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })));
            const cd = plain.readUInt32LE(plain.length - 22 + 16);
            let bytes: Buffer;
            if (label.startsWith("the 32-bit")) {
                bytes = plain;
                bytes.writeUInt32LE(0x80000000, cd + 20);
            } else {
                const field = Buffer.alloc(20);
                field.writeUInt16LE(0x0001, 0);
                field.writeUInt16LE(16, 2);
                field.writeBigUInt64LE(BigInt(plain.readUInt32LE(cd + 24)), 4);
                field.writeBigUInt64LE(0xffffffff80000000n, 12);
                const insert = cd + 46 + plain.readUInt16LE(cd + 28);
                bytes = Buffer.concat([plain.subarray(0, insert), field, plain.subarray(insert)]);
                bytes.writeUInt32LE(0xffffffff, cd + 20);
                bytes.writeUInt32LE(0xffffffff, cd + 24);
                bytes.writeUInt16LE(bytes.readUInt16LE(cd + 30) + field.length, cd + 30);
                const eocd = bytes.length - 22;
                bytes.writeUInt32LE(bytes.readUInt32LE(eocd + 12) + field.length, eocd + 12);
            }
            const specimen = Buffer.concat([Buffer.alloc(64, 0x41), bytes]);
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, specimen);
                for (const profile of ["word-valid", "strict"] as const) {
                    const full = await validate(file, { profile });
                    expect(full.valid, `${label} ${profile}`).toBe(false);
                    expect([...errorCodes(full.issues)], `${label} ${profile}`).toEqual(["package-open-failed"]);
                    if (profile === "strict") continue; // Strict rejects partial selectors.
                    for (const selection of [{ quick: true }, { checks: ["zip-bytes-before-archive"] }, { checks: ["package"] }]) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${label} ${profile} ${JSON.stringify(selection)}`;
                        expect(partial.valid, tag).toBe(false);
                        expect(errorCodes(partial.issues).has("package-open-failed"), tag).toBe(true);
                        expect(
                            [...errorCodes(partial.issues)].every((code) => errorCodes(full.issues).has(code)),
                            `${tag}: partial errors ⊆ full errors`,
                        ).toBe(true);
                    }
                }
            });
        },
    );
});

describe("validation contract: no prefix claim where JSZip's local name/extra reads throw (#566, #381)", { timeout: 120_000 }, () => {
    // Review of layer 112: `readLocalPart` bounds the local name and the
    // extra field on their own reads before `readData(compressedSize)`, so
    // a small negative resolved size — a ZIP64 0x0001 extra resolving to
    // -8, or the 32-bit field 0xFFFFFFF8 — still throws on a name/extra
    // overrun of up to |size| bytes; the combined bound #555 added misses
    // exactly that window. The archive never opens and no prefix is
    // claimed beside the refusal, full and partial runs alike.
    it.each([["the 32-bit compressed-size field 0xFFFFFFF8"], ["a ZIP64 extra resolving the compressed size to -8"]])(
        "word-valid and strict report the refusal alone for %s, and partial runs agree",
        async ([label]) => {
            const plain = buildRawZip(Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })));
            const cd = plain.readUInt32LE(plain.length - 22 + 16);
            let bytes: Buffer;
            if (label.startsWith("the 32-bit")) {
                bytes = plain;
                bytes.writeUInt32LE(0xfffffff8, cd + 20);
            } else {
                const field = Buffer.alloc(20);
                field.writeUInt16LE(0x0001, 0);
                field.writeUInt16LE(16, 2);
                field.writeBigUInt64LE(BigInt(plain.readUInt32LE(cd + 24)), 4);
                field.writeBigUInt64LE(0xfffffffffffffff8n, 12);
                const insert = cd + 46 + plain.readUInt16LE(cd + 28);
                bytes = Buffer.concat([plain.subarray(0, insert), field, plain.subarray(insert)]);
                bytes.writeUInt32LE(0xffffffff, cd + 20);
                bytes.writeUInt32LE(0xffffffff, cd + 24);
                bytes.writeUInt16LE(bytes.readUInt16LE(cd + 30) + field.length, cd + 30);
                const eocd = bytes.length - 22;
                bytes.writeUInt32LE(bytes.readUInt32LE(eocd + 12) + field.length, eocd + 12);
            }
            const specimen = Buffer.concat([Buffer.alloc(64, 0x41), bytes]);
            // The first record's local header sits at archive offset 0 —
            // buffer offset 64 — and its declared extra length ends the
            // local name and extra field 4 bytes past the buffer.
            const nameLength = specimen.readUInt16LE(64 + 26);
            specimen.writeUInt16LE(specimen.length + 4 - (64 + 30 + nameLength), 64 + 28);
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, specimen);
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const full = await validate(file, { profile });
                    expect(full.valid, `${label} ${profile}`).toBe(false);
                    expect([...errorCodes(full.issues)], `${label} ${profile}`).toEqual(["package-open-failed"]);
                    if (profile !== "word-valid") continue; // Strict rejects partial selectors; lenient runs no envelope check.
                    for (const selection of [{ quick: true }, { checks: ["zip-bytes-before-archive"] }, { checks: ["package"] }]) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${label} ${profile} ${JSON.stringify(selection)}`;
                        expect(partial.valid, tag).toBe(false);
                        expect(errorCodes(partial.issues).has("package-open-failed"), tag).toBe(true);
                        expect(
                            [...errorCodes(partial.issues)].every((code) => errorCodes(full.issues).has(code)),
                            `${tag}: partial errors ⊆ full errors`,
                        ).toBe(true);
                    }
                }
            });
        },
    );
});

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

    // #455: once the grammar follows ECMA-376 Part 2 §9.1.1.1 and accepts a
    // literal-ucschar PartName, only the Word-specific rule keeps refusing it,
    // so the superset contract runs through that rule alone.
    it("a literal Part IRI PartName Word refuses fails both profiles through ct-override-partname-word-refused", async () => {
        const parts = {
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/café.xml" ContentType="application/xml"/></Types>',
            ),
        };
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(parts), profile);
            expect(errorCodes(result.issues)).toEqual(new Set(["ct-override-partname-word-refused"]));
            expect(result.valid).toBe(false);
        }
    });
});

describe("validation contract: the part-name grammar follows RFC 3987 §2.2 (#522)", { timeout: 120_000 }, () => {
    const withOverride = (partName: string): Readonly<Record<string, string>> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="${partName}" ContentType="application/xml"/></Types>`,
        ),
    });

    // RFC 3987 §2.2 ends `ucschar` with %xE1000-EFFFD, so U+E0000..U+E0FFF
    // is in neither the RFC 3986 `pchar` nor the RFC 3987 `ipchar` set: the
    // Override PartName names no part, and both profiles report it.
    it.each([
        ["/word/x\u{E0000}y.xml", "U+E0000"],
        ["/word/x\u{E0FFF}y.xml", "U+E0FFF"],
    ])("%s is refused as ct-override-partname-invalid under both profiles", async (partName) => {
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(withOverride(partName)), profile);
            expect(errorCodes(result.issues).has("ct-override-partname-invalid"), profile).toBe(true);
            expect(result.valid, profile).toBe(false);
        }
    });

    // The control: U+E1000 is the first code point of the last `ucschar`
    // block, so the grammar accepts the name and only Word's refusal of a
    // literal non-ASCII PartName is left.
    it("the control at U+E1000 draws no part-name finding under either profile", async () => {
        const partName = "/word/x\u{E1000}y.xml";
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(withOverride(partName)), profile);
            expect(errorCodes(result.issues).has("ct-override-partname-invalid"), profile).toBe(false);
            expect(errorCodes(result.issues).has("ct-override-partname-word-refused"), profile).toBe(true);
        }
    });

    it("keeps partial runs in agreement with the full run (#381)", async () => {
        const partName = "/word/x\u{E0000}y.xml";
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(withOverride(partName)));
            const full = await validate(file, { profile: "word-valid" });
            expect(full.valid).toBe(false);
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: ["ct-override-partname-invalid"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(partial.valid, tag).toBe(full.valid);
                expect(
                    partial.issues.filter((issue) => issue.code === "ct-override-partname-invalid"),
                    tag,
                ).toEqual(full.issues.filter((issue) => issue.code === "ct-override-partname-invalid"));
                expect(errorCodes(partial.issues), tag).toEqual(errorCodes(full.issues));
            }
        });
    });
});

describe("validation contract: a malformed escape does not hide a later encoded iunreserved (#523)", { timeout: 120_000 }, () => {
    const withOverride = (partName: string): Readonly<Record<string, string>> => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="${partName}" ContentType="application/xml"/></Types>`,
        ),
    });

    // OPC [M1.8] reads "iunreserved" in the Part IRI form: a literal `é`
    // leaves only that form, and the `%C3%A9` behind a malformed, overlong
    // or truncated escape still encodes an `iunreserved` character.
    it.each([
        ["/word/x%E9%C3%A9-é.xml", "a lone continuation byte"],
        ["/word/x%C0%AF%C3%A9-é.xml", "an overlong lead"],
        ["/word/x%C3%C3%A9-é.xml", "a truncated sequence"],
    ])("%s is refused as ct-override-partname-invalid under both profiles", async (partName) => {
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(withOverride(partName)), profile);
            expect(errorCodes(result.issues).has("ct-override-partname-invalid"), profile).toBe(true);
            expect(result.valid, profile).toBe(false);
        }
    });

    // The control: a valid Part URI, whose escape encodes a non-ASCII
    // character only, so the grammar accepts the name.
    it("the control %s draws no part-name finding under either profile", async () => {
        const partName = "/word/caf%C3%A9.xml";
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(withOverride(partName)), profile);
            expect(errorCodes(result.issues).has("ct-override-partname-invalid"), profile).toBe(false);
        }
    });

    it("keeps partial runs in agreement with the full run (#381)", async () => {
        const partName = "/word/x%E9%C3%A9-é.xml";
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(withOverride(partName)));
            const full = await validate(file, { profile: "word-valid" });
            expect(full.valid).toBe(false);
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: ["ct-override-partname-invalid"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(partial.valid, tag).toBe(full.valid);
                expect(
                    partial.issues.filter((issue) => issue.code === "ct-override-partname-invalid"),
                    tag,
                ).toEqual(full.issues.filter((issue) => issue.code === "ct-override-partname-invalid"));
                // #381: the partial run's errors are a subset of the full
                // run's; a named run judges fewer checks.
                expect([...errorCodes(full.issues)], tag).toEqual(expect.arrayContaining([...errorCodes(partial.issues)]));
            }
        });
    });
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

    // An MC element that carries an MC attribute of its own.
    const withChoiceAttribute = {
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x="urn:x:unknown">` +
                `<mc:Choice Requires="x" mc:MustUnderstand="x"/><mc:Fallback/></mc:AlternateContent></Types>`,
        ),
    };

    for (const profile of ["word-valid", "strict"] as const) {
        it(`${profile}: mc:Ignorable in the content-types stream is rejected as ct-markup-compatibility`, async () => {
            const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "[Content_Types].xml": withIgnorable }), profile);
            expect(errorCodes(result.issues).has("ct-markup-compatibility")).toBe(true);
            expect(result.valid).toBe(false);
        });

        // #526: DOM document order puts an element before its own attributes.
        it(`${profile}: the Markup Compatibility message lists an MC element before its own MC attribute`, async () => {
            const result = await validateBytes(await buildDocx(withChoiceAttribute), profile);
            const findings = result.issues.filter((issue) => issue.code === "ct-markup-compatibility");
            expect(findings).toHaveLength(1);
            expect(findings[0]!.message).toContain(
                "<mc:AlternateContent> element; <mc:Choice> element; mc:MustUnderstand attribute on <mc:Choice>",
            );
        });
    }

    it("the Markup Compatibility message is the same in the full and the partial runs (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(withChoiceAttribute));
            const full = await validate(file, { profile: "word-valid" });
            const message = full.issues.find((issue) => issue.code === "ct-markup-compatibility")?.message;
            expect(message).toBeDefined();
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: ["ct-markup-compatibility"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(partial.issues.find((issue) => issue.code === "ct-markup-compatibility")?.message, tag).toBe(message);
                expect(partial.valid, tag).toBe(full.valid);
            }
        });
    });
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
            async (file) => {
                const quick = await validate(file, { checks: ["xml-syntax"] });
                const full = await validate(file, { profile: "lenient" });
                expect(quick.issues.filter((issue) => issue.code === "xml-syntax")).toEqual([]);
                expect(errorCodes(quick.issues)).toContain("package-extract-failed");
                expect(full.valid).toBe(false);
                expect(errorCodes(full.issues)).toContain("package-extract-failed");
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
        await withSpecimen(parts, async (file) => {
            const quick = await validate(file, { checks: ["word-attached-toolbar-utf8"] });
            const full = await validate(file, { profile: "word-valid" });
            expect(errorCodes(quick.issues)).toContain("package-extract-failed");
            expect(errorCodes(quick.issues)).not.toContain("word-attached-toolbar-utf8");
            expect(full.valid).toBe(false);
            expect(errorCodes(full.issues)).toContain("package-extract-failed");
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

describe("validation contract: a Word-read part typed by a generic XML type (#517)", { timeout: 120000 }, () => {
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const CT = MINIMAL_PARTS["[Content_Types].xml"]!;
    const DOCM = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const specimen = (mutated: string, extra: Record<string, string> = {}, rels?: string) =>
        buildDocx({
            ...MINIMAL_PARTS,
            ...extra,
            "word/_rels/document.xml.rels": rels ?? MINIMAL_PARTS["word/_rels/document.xml.rels"]!,
            "[Content_Types].xml": mutated,
        });

    for (const [label, bytes] of [
        ["the main part's Override to application/xml", specimen(CT.replace(`ContentType="${DOCM}"`, 'ContentType="application/xml"'))],
        ["the main part's Override to text/xml", specimen(CT.replace(`ContentType="${DOCM}"`, 'ContentType="text/xml"'))],
        [
            "styles typed by the Default to text/xml",
            specimen(
                CT.replace(`<Override PartName="/word/styles.xml" ContentType="${STYLES}"/>`, "").replace(
                    '<Default Extension="xml" ContentType="application/xml"/>',
                    '<Default Extension="xml" ContentType="text/xml"/>',
                ),
            ),
        ],
        [
            "styles typed by an Override to application/xml",
            specimen(CT.replace(`ContentType="${STYLES}"`, 'ContentType="application/xml"')),
        ],
        [
            "settings typed by an Override to application/xml",
            specimen(
                CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/xml"/></Types>'),
                { "word/settings.xml": `<w:settings xmlns:w="${W}"/>` },
                MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rId9" Type="${R}/settings" Target="settings.xml"/></Relationships>`,
                ),
            ),
        ],
    ] as const) {
        it(`word-valid and strict both reject ${label}`, async () => {
            const wordValid = await validateBytes(await bytes, "word-valid");
            expect(wordValid.issues).toContainEqual(expect.objectContaining({ code: "word-part-content-type-generic", severity: "error" }));
            expect(wordValid.valid).toBe(false);
            const strict = await validateBytes(await bytes, "strict");
            expect(strict.issues).toContainEqual(expect.objectContaining({ code: "word-part-content-type-generic", severity: "error" }));
            expect(strict.valid).toBe(false);
        });
    }
});

describe("validation contract: an inexact content type on a Word-read part (#550)", { timeout: 120000 }, () => {
    const STREAM = MINIMAL_PARTS["[Content_Types].xml"]!;
    const RELS = MINIMAL_PARTS["word/_rels/document.xml.rels"]!;
    const DOCM = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    const CODE = "word-part-content-type-inexact";

    const specimen = (stream: string, extra: Record<string, string> = {}, rels: string = RELS) =>
        buildDocx({ ...MINIMAL_PARTS, ...extra, "word/_rels/document.xml.rels": rels, "[Content_Types].xml": stream });

    /** A settings, fontTable or webSettings part related from the main document, its Override carrying `contentType` (session 17's `docPart`). */
    const docPart = (segment: string, file: string, root: string, contentType: string) => ({
        stream: STREAM.replace("</Types>", `<Override PartName="/word/${file}" ContentType="${contentType}"/></Types>`),
        extra: { [`word/${file}`]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${root} xmlns:w="${W}"/>` },
        rels: RELS.replace("</Relationships>", `<Relationship Id="rId9" Type="${R}/${segment}" Target="${file}"/></Relationships>`),
        part: `word/${file}`,
    });

    const docPartSpecimen = (segment: string, file: string, root: string, contentType: string) => {
        const typed = docPart(segment, file, root, contentType);
        return specimen(typed.stream, typed.extra, typed.rels);
    };

    /** A header the main document references through `w:headerReference`, its Override carrying `contentType` (session 17's `headerInMain`). */
    const headerInMain = (contentType: string) => ({
        stream: STREAM.replace("</Types>", `<Override PartName="/word/header9.xml" ContentType="${contentType}"/></Types>`),
        extra: {
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
                "<w:sectPr>",
                `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`,
            ),
            "word/header9.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
        },
        rels: RELS.replace("</Relationships>", `<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/></Relationships>`),
        part: "word/header9.xml",
    });

    const headerSpecimen = (contentType: string) => {
        const typed = headerInMain(contentType);
        return specimen(typed.stream, typed.extra, typed.rels);
    };

    // Session 20 (2026-10-24, Word 16.115; sessions 20 and 21 in
    // docs/plans/evidence/word-probes-2026-10-24.md) probed the part types
    // #550 left unprobed: a related numbering or theme part and a referenced
    // footer, footnotes, endnotes or comments part typed `type; charset=utf-8`
    // get the repair warning, a parameterized glossaryDocument opens clean,
    // and every exact-type control opens clean (#570). Session 21 observed
    // a trailing space refused on the same six parts.
    const WMLT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const THEME_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";

    /** The well-formed theme part session 20's specimens carry — `word/theme/theme1.xml` of `package.word-anchor-open-repaired.docx`, copied not modified. */
    let themeXml: string | undefined;
    const theme = async (): Promise<string> => {
        if (themeXml === undefined) {
            const zip = await JSZip.loadAsync(await fs.readFile(path.join(WORKING, "package.word-anchor-open-repaired.docx")));
            themeXml = await zip.file("word/theme/theme1.xml")!.async("string");
        }
        return themeXml;
    };

    const notesXml = (kind: "footnote" | "endnote"): string =>
        `<w:${kind}s xmlns:w="${W}"><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
        `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
        `<w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`;
    const NUMBERING_XML =
        `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
        `<w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const FOOTER_XML = `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`;
    const COMMENTS_XML = `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A" w:initials="A"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`;
    const GLOSSARY_XML =
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

    /** The main document with `xmlns:r` declared, `run` appended after the Hello run and `sect` prepended into `w:sectPr` — the s20 builder's `docWith`. */
    const docUsing = (run: string, sect: string): string =>
        MINIMAL_PARTS["word/document.xml"]!.replace(`<w:document xmlns:w="${W}"`, `<w:document xmlns:w="${W}" xmlns:r="${R}"`)
            .replace("<w:t>Hello</w:t></w:r>", `<w:t>Hello</w:t></w:r>${run}`)
            .replace("<w:sectPr>", `<w:sectPr>${sect}`);

    /** The part shapes session 20's builder produced, keyed by the relationship type's segment. */
    const S20_PARTS = {
        numbering: { file: "numbering.xml", type: `${WMLT}.numbering+xml`, xml: (): string => NUMBERING_XML },
        theme: { file: "theme/theme1.xml", type: THEME_TYPE, xml: theme },
        glossaryDocument: { file: "glossary/document.xml", type: `${WMLT}.document.glossary+xml`, xml: (): string => GLOSSARY_XML },
        footer: {
            file: "footer9.xml",
            type: `${WMLT}.footer+xml`,
            xml: (): string => FOOTER_XML,
            sect: `<w:footerReference w:type="default" r:id="rIdX"/>`,
        },
        footnotes: {
            file: "footnotes.xml",
            type: `${WMLT}.footnotes+xml`,
            xml: (): string => notesXml("footnote"),
            run: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
        },
        endnotes: {
            file: "endnotes.xml",
            type: `${WMLT}.endnotes+xml`,
            xml: (): string => notesXml("endnote"),
            run: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
        },
        comments: {
            file: "comments.xml",
            type: `${WMLT}.comments+xml`,
            xml: (): string => COMMENTS_XML,
            run: `<w:r><w:commentReference w:id="0"/></w:r>`,
        },
    } as const;
    type S20Segment = keyof typeof S20_PARTS;

    /** The parts session 20's `mainPart` added to the minimal package — the related part and, when used, the referencing document. */
    const s20Parts = async (segment: S20Segment, used = true): Promise<Record<string, string>> => {
        const spec = S20_PARTS[segment];
        const run = "run" in spec ? spec.run : "";
        const sect = "sect" in spec ? spec.sect : "";
        return {
            ...(used && (run !== "" || sect !== "") ? { "word/document.xml": docUsing(run, sect) } : {}),
            [`word/${spec.file}`]: await spec.xml(),
        };
    };

    const s20Specimen = async (segment: S20Segment, contentType: string, used = true): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            ...(await s20Parts(segment, used)),
            "word/_rels/document.xml.rels": RELS.replace(
                "</Relationships>",
                `<Relationship Id="rIdX" Type="${R}/${segment}" Target="${S20_PARTS[segment].file}"/></Relationships>`,
            ),
            "[Content_Types].xml": STREAM.replace(
                "</Types>",
                `<Override PartName="/word/${S20_PARTS[segment].file}" ContentType="${contentType}"/></Types>`,
            ),
        });

    it.each([
        {
            label: "the main part's Override carrying the main type plus `; charset=utf-8`",
            bytes: specimen(STREAM.replace(`ContentType="${DOCM}"`, `ContentType="${DOCM}; charset=utf-8"`)),
            code: CODE,
            part: "word/document.xml",
        },
        {
            label: "the styles Override carrying the styles type plus `; charset=utf-8`",
            bytes: specimen(STREAM.replace(`ContentType="${STYLES}"`, `ContentType="${STYLES}; charset=utf-8"`)),
            code: CODE,
            part: "word/styles.xml",
        },
        {
            label: "the settings Override carrying the settings type plus `; charset=utf-8` (s17b)",
            bytes: docPartSpecimen(
                "settings",
                "settings.xml",
                "settings",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml; charset=utf-8",
            ),
            code: CODE,
            part: "word/settings.xml",
        },
        {
            label: "the fontTable Override carrying the fontTable type plus `; charset=utf-8` (s17b)",
            bytes: docPartSpecimen(
                "fontTable",
                "fontTable.xml",
                "fonts",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml; charset=utf-8",
            ),
            code: CODE,
            part: "word/fontTable.xml",
        },
        {
            label: "the webSettings Override carrying the webSettings type plus `; charset=utf-8` (s17b)",
            bytes: docPartSpecimen(
                "webSettings",
                "webSettings.xml",
                "webSettings",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml; charset=utf-8",
            ),
            code: CODE,
            part: "word/webSettings.xml",
        },
        {
            label: "a referenced header's Override carrying the header type plus `; charset=utf-8` (s17b)",
            bytes: headerSpecimen(`${HEADER}; charset=utf-8`),
            code: CODE,
            part: "word/header9.xml",
        },
        {
            label: "the main part's Override carrying the main type plus `;charset=utf-8`, no space (s17c)",
            bytes: specimen(STREAM.replace(`ContentType="${DOCM}"`, `ContentType="${DOCM};charset=utf-8"`)),
            code: CODE,
            part: "word/document.xml",
        },
        {
            label: "the main part's Override carrying the main type plus `; foo=bar` (s17c)",
            bytes: specimen(STREAM.replace(`ContentType="${DOCM}"`, `ContentType="${DOCM}; foo=bar"`)),
            code: CODE,
            part: "word/document.xml",
        },
        {
            label: "the styles Override carrying the styles type plus a trailing space (s17c)",
            bytes: specimen(STREAM.replace(`ContentType="${STYLES}"`, `ContentType="${STYLES} "`)),
            code: CODE,
            part: "word/styles.xml",
        },
        {
            label: "the styles Override to `application/xml; charset=utf-8`",
            bytes: specimen(STREAM.replace(`ContentType="${STYLES}"`, 'ContentType="application/xml; charset=utf-8"')),
            code: "word-part-content-type-generic",
            part: "word/styles.xml",
        },
        {
            label: "the numbering Override carrying the numbering type plus `; charset=utf-8` (s20b)",
            bytes: s20Specimen("numbering", `${WMLT}.numbering+xml; charset=utf-8`),
            code: CODE,
            part: "word/numbering.xml",
        },
        {
            label: "the theme Override carrying the theme type plus `; charset=utf-8` (s20b)",
            bytes: s20Specimen("theme", `${THEME_TYPE}; charset=utf-8`),
            code: CODE,
            part: "word/theme/theme1.xml",
        },
        {
            label: "a referenced footer's Override carrying the footer type plus `; charset=utf-8` (s20b)",
            bytes: s20Specimen("footer", `${WMLT}.footer+xml; charset=utf-8`),
            code: CODE,
            part: "word/footer9.xml",
        },
        {
            label: "a referenced footnotes part's Override carrying the footnotes type plus `; charset=utf-8` (s20b)",
            bytes: s20Specimen("footnotes", `${WMLT}.footnotes+xml; charset=utf-8`),
            code: CODE,
            part: "word/footnotes.xml",
        },
        {
            label: "a referenced endnotes part's Override carrying the endnotes type plus `; charset=utf-8` (s20b)",
            bytes: s20Specimen("endnotes", `${WMLT}.endnotes+xml; charset=utf-8`),
            code: CODE,
            part: "word/endnotes.xml",
        },
        {
            label: "a referenced comments part's Override carrying the comments type plus `; charset=utf-8` (s20c)",
            bytes: s20Specimen("comments", `${WMLT}.comments+xml; charset=utf-8`),
            code: CODE,
            part: "word/comments.xml",
        },
        {
            label: "the numbering Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("numbering", `${WMLT}.numbering+xml `),
            code: CODE,
            part: "word/numbering.xml",
        },
        {
            label: "the theme Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("theme", `${THEME_TYPE} `),
            code: CODE,
            part: "word/theme/theme1.xml",
        },
        {
            label: "a referenced footer's Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("footer", `${WMLT}.footer+xml `),
            code: CODE,
            part: "word/footer9.xml",
        },
        {
            label: "a referenced footnotes part's Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("footnotes", `${WMLT}.footnotes+xml `),
            code: CODE,
            part: "word/footnotes.xml",
        },
        {
            label: "a referenced endnotes part's Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("endnotes", `${WMLT}.endnotes+xml `),
            code: CODE,
            part: "word/endnotes.xml",
        },
        {
            label: "a referenced comments part's Override carrying its own type plus a trailing space (s21a)",
            bytes: s20Specimen("comments", `${WMLT}.comments+xml `),
            code: CODE,
            part: "word/comments.xml",
        },
    ] as const)(
        "word-valid and strict both reject $label",
        async ({ bytes, code, part }) => {
            const wordValid = await validateBytes(await bytes, "word-valid");
            expect(wordValid.issues).toContainEqual(expect.objectContaining({ code, severity: "error", path: part }));
            expect(wordValid.valid).toBe(false);
            const strict = await validateBytes(await bytes, "strict");
            expect(strict.issues).toContainEqual(expect.objectContaining({ code, severity: "error", path: part }));
            expect(strict.valid).toBe(false);
            expect(strict.valid && !wordValid.valid, "strict success must imply word-valid success").toBe(false);
        },
        60_000,
    );

    it("the exact and upper-case controls stay valid under every profile", async () => {
        for (const [label, bytes] of [
            ["the minimal package", await buildDocx(MINIMAL_PARTS)],
            [
                "the main part's Override in upper case (s17c)",
                await specimen(STREAM.replace(`ContentType="${DOCM}"`, `ContentType="${DOCM.toUpperCase()}"`)),
            ],
            [
                "the styles Override in upper case (s17c)",
                await specimen(STREAM.replace(`ContentType="${STYLES}"`, `ContentType="${STYLES.toUpperCase()}"`)),
            ],
        ] as const) {
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(result.valid, `${label} ${profile}`).toBe(true);
                expect(
                    result.issues.filter(
                        (issue) => issue.code === "word-part-content-type-generic" || issue.code === "word-part-content-type-inexact",
                    ),
                    `${label} ${profile}`,
                ).toEqual([]);
            }
        }
    }, 60_000);

    it.each(["footer", "footnotes", "endnotes", "comments"] as const)(
        "keeps a related-but-unused %s part's parameterized Override silent under word-valid while strict still fails it citing clause 11.3 (#516, #570, #589)",
        async (segment) => {
            const contentType = `${S20_PARTS[segment].type}; charset=utf-8`;
            const bytes = await s20Specimen(segment, contentType, false);
            const wordValid = await validateBytes(bytes, "word-valid");
            expect(wordValid.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            expect(wordValid.valid).toBe(true);
            const strict = await validateBytes(bytes, "strict");
            expect(strict.issues).toContainEqual(
                expect.objectContaining({ code: CODE, severity: "error", path: `word/${S20_PARTS[segment].file}` }),
            );
            // s21b opened these specimens clean, so strict's finding cites the
            // spec and claims no Word recovery for a part Word does not read.
            const finding = strict.issues.find((issue) => issue.code === CODE && issue.path === `word/${S20_PARTS[segment].file}`);
            expect(finding?.message, segment).toContain(`clause 11.3 assigns the ${segment} part`);
            expect(finding?.message, segment).not.toContain("offers to recover the package");
            expect(strict.valid).toBe(false);
        },
        60_000,
    );

    it("keeps a related-but-unused header's parameterized Override silent under word-valid while strict cites clause 11.3 with no recovery claim, and partial runs agree (#589, #381)", async () => {
        await withTempDir(async (dir) => {
            const typed = headerInMain(`${HEADER}; charset=utf-8`);
            // The same shape minus the `w:headerReference`: related but unused,
            // s21b's `header-unused-override-charset` specimen.
            const file = path.join(dir, "unused-header.docx");
            writeFileSync(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/header9.xml": typed.extra["word/header9.xml"]!,
                    "word/_rels/document.xml.rels": typed.rels,
                    "[Content_Types].xml": typed.stream,
                }),
            );
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            expect(wordValid.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            const finding = strict.issues.find((issue) => issue.code === CODE && issue.path === "word/header9.xml");
            expect(finding?.message).toContain("clause 11.3 assigns the header part");
            expect(finding?.message).not.toContain("offers to recover the package");
            expect(strict.valid).toBe(false);
            expect(strict.valid && !wordValid.valid, "strict success must imply word-valid success").toBe(false);
            // The unused header is no word-valid finding in partial runs either.
            for (const selection of [{ quick: true }, { checks: [CODE] }, { checks: ["content-types"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = `unused-header ${JSON.stringify(selection)}`;
                expect(partial.valid, tag).toBe(wordValid.valid);
                expect(
                    partial.issues.filter((issue) => issue.code === CODE),
                    tag,
                ).toEqual([]);
            }
        });
    }, 60_000);

    it("keeps a parameterized glossaryDocument Override silent under word-valid — session 20 opened it clean — while strict still fails it (#570)", async () => {
        const bytes = await s20Specimen("glossaryDocument", `${WMLT}.document.glossary+xml; charset=utf-8`);
        const wordValid = await validateBytes(bytes, "word-valid");
        expect(wordValid.issues.filter((issue) => issue.code === CODE)).toEqual([]);
        expect(wordValid.valid).toBe(true);
        const strict = await validateBytes(bytes, "strict");
        expect(strict.issues).toContainEqual(
            expect.objectContaining({ code: CODE, severity: "error", path: "word/glossary/document.xml" }),
        );
        expect(strict.valid).toBe(false);
    }, 60_000);

    // Sessions 21 and 22 (2026-10-24) probed an inexact type on the parts a
    // reached glossary's own relationships part names: its styles part —
    // eagerly parsed, so related is enough (#562, session 19) — and a header
    // the glossary references through `w:headerReference` both get the repair
    // warning, while every exact-type control opens clean (#579). The finding
    // names the glossary's relationship, not the main document's.
    const RELS_PART = (body: string): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    const GLOSSARY_CT = `${WMLT}.document.glossary+xml`;
    /** The s22 builder's `glossaryWithSect` — the reached glossary whose `docPartBody` ends with a `w:sectPr` holding `sect`. */
    const glossaryWithSect = (sect: string): string =>
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p><w:sectPr>${sect}</w:sectPr></w:docPartBody>` +
        `</w:docPart></w:docParts></w:glossaryDocument>`;

    /** The s21/s22 builders' reached-glossary shapes: the minimal package plus a glossary whose own `.rels` relates `stored` at `contentType`. */
    const glossarySpecimen = (segment: string, stored: string, xml: string, contentType: string, glossary: string = GLOSSARY_XML) =>
        specimen(
            STREAM.replace(
                "</Types>",
                `<Override PartName="/word/glossary/document.xml" ContentType="${GLOSSARY_CT}"/>` +
                    `<Override PartName="/${stored}" ContentType="${contentType}"/></Types>`,
            ),
            {
                "word/glossary/document.xml": glossary,
                "word/glossary/_rels/document.xml.rels": RELS_PART(
                    `<Relationship Id="rIdE" Type="${R}/${segment}" Target="${path.posix.basename(stored)}"/>`,
                ),
                [stored]: xml,
            },
            RELS.replace(
                "</Relationships>",
                `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
            ),
        );

    it.each([
        [
            "styles (s21b)",
            "word/glossary/styles.xml",
            glossarySpecimen("styles", "word/glossary/styles.xml", `<w:styles xmlns:w="${W}"></w:styles>`, `${STYLES}; charset=utf-8`),
            "the glossary's styles relationship",
        ],
        [
            "header (s22a)",
            "word/glossary/header9.xml",
            glossarySpecimen(
                "header",
                "word/glossary/header9.xml",
                `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
                `${HEADER}; charset=utf-8`,
                glossaryWithSect(`<w:headerReference w:type="default" r:id="rIdE"/>`),
            ),
            "the glossary's header relationship",
        ],
    ] as const)(
        "word-valid and strict both reject the reached glossary's %s part typed `type; charset=utf-8`, naming the glossary's relationship (#579)",
        async (_label, part, bytes, by) => {
            const wordValid = await validateBytes(await bytes, "word-valid");
            expect(wordValid.issues).toContainEqual(
                expect.objectContaining({ code: CODE, severity: "error", path: part, message: expect.stringContaining(by) }),
            );
            expect(wordValid.valid).toBe(false);
            const strict = await validateBytes(await bytes, "strict");
            expect(strict.issues).toContainEqual(expect.objectContaining({ code: CODE, severity: "error", path: part }));
            expect(strict.valid).toBe(false);
            expect(strict.valid && !wordValid.valid, "strict success must imply word-valid success").toBe(false);
        },
        60_000,
    );

    it("partial runs agree with the full run (#381)", async () => {
        await withTempDir(async (dir) => {
            const code = CODE;
            const shapes: readonly [string, string, Record<string, string>, string][] = [
                ["parameter", STREAM.replace(`ContentType="${DOCM}"`, `ContentType="${DOCM}; charset=utf-8"`), {}, RELS],
                [
                    "header",
                    headerInMain(`${HEADER}; charset=utf-8`).stream,
                    headerInMain(`${HEADER}; charset=utf-8`).extra,
                    headerInMain(`${HEADER}`).rels,
                ],
                // Session 20's widened scope: an eagerly parsed segment and a
                // lazily parsed used one must agree in partial runs too (#381, #570).
                [
                    "numbering",
                    STREAM.replace(
                        "</Types>",
                        `<Override PartName="/word/numbering.xml" ContentType="${WMLT}.numbering+xml; charset=utf-8"/></Types>`,
                    ),
                    await s20Parts("numbering"),
                    RELS.replace(
                        "</Relationships>",
                        `<Relationship Id="rIdX" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
                    ),
                ],
                [
                    "footer",
                    STREAM.replace(
                        "</Types>",
                        `<Override PartName="/word/footer9.xml" ContentType="${WMLT}.footer+xml; charset=utf-8"/></Types>`,
                    ),
                    await s20Parts("footer"),
                    RELS.replace("</Relationships>", `<Relationship Id="rIdX" Type="${R}/footer" Target="footer9.xml"/></Relationships>`),
                ],
                // Session 21's widened scope: a reached glossary's styles part
                // must agree in partial runs too (#381, #579).
                [
                    "glossary-styles",
                    STREAM.replace(
                        "</Types>",
                        `<Override PartName="/word/glossary/document.xml" ContentType="${GLOSSARY_CT}"/>` +
                            `<Override PartName="/word/glossary/styles.xml" ContentType="${STYLES}; charset=utf-8"/></Types>`,
                    ),
                    {
                        "word/glossary/document.xml": GLOSSARY_XML,
                        "word/glossary/_rels/document.xml.rels": RELS_PART(
                            `<Relationship Id="rIdE" Type="${R}/styles" Target="styles.xml"/>`,
                        ),
                        "word/glossary/styles.xml": `<w:styles xmlns:w="${W}"></w:styles>`,
                    },
                    RELS.replace(
                        "</Relationships>",
                        `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
                    ),
                ],
            ];
            for (const [name, stream, extra, rels] of shapes) {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await specimen(stream, extra, rels));
                const word = await validate(file, { profile: "word-valid" });
                expect(word.valid, name).toBe(false);
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid && !word.valid, `${name}: strict success must imply word-valid success`).toBe(false);
                for (const selection of [{ quick: true }, { checks: [code] }, { checks: ["content-types"] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    const tag = `${name} ${JSON.stringify(selection)}`;
                    expect(partial.valid, tag).toBe(word.valid);
                    expect(
                        partial.issues.filter((issue) => issue.code === code),
                        tag,
                    ).toEqual(word.issues.filter((issue) => issue.code === code));
                    expect(
                        [...errorCodes(partial.issues)].every((error) => errorCodes(word.issues).has(error)),
                        tag,
                    ).toBe(true);
                }
            }
        });
    }, 60_000);
});

describe("validation contract: a type with an extra path segment is no Office type (#519)", { timeout: 120000 }, () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const R_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const unclosed = (rootOpen: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootOpen}`;
    const MALFORMED = unclosed(`<w:settings xmlns:w="${W}">`);

    it("reports no word-part-content-type-generic for a generic part named by `…/relationships/custom/styles`", async () => {
        // `custom/styles` is not the styles relationship type — a type has a
        // segment only when it is exactly the namespace plus one segment — so
        // its target is an unobserved Office type's target: no generic-type
        // finding, under word-valid or strict.
        const parts = {
            ...MINIMAL_PARTS,
            "word/extra.xml": `<w:styles xmlns:w="${W}"/>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rId9" Type="${R}/custom/styles" Target="extra.xml"/></Relationships>`,
            ),
        };
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(parts), profile);
            expect(
                result.issues.some((issue) => issue.code === "word-part-content-type-generic"),
                profile,
            ).toBe(false);
            expect(result.valid, profile).toBe(true);
        }
    });

    it.each([
        ["Transitional", R],
        ["Strict", R_STRICT],
    ] as const)(
        "keeps the word/ folder rule under word-valid and strict for a malformed word/ part named by the %s `…/custom/styles`",
        { timeout: 120000 },
        async (_label, namespace) => {
            // An unknown type inside an Office relationship namespace is
            // unprobed (session 5b probed only `urn:x:unknown`, clean-open), so
            // its target keeps the conservative word/ rule — blocking — rather
            // than the styles EAGER path.
            const parts = {
                ...MINIMAL_PARTS,
                "word/xtra.xml": MALFORMED,
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rId9" Type="${namespace}/custom/styles" Target="xtra.xml"/></Relationships>`,
                ),
                "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                    "</Types>",
                    '<Override PartName="/word/xtra.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
                ),
            };
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(await buildDocx(parts), profile);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/xtra.xml")?.severity,
                    profile,
                ).toBe("error");
                expect(result.valid, profile).toBe(false);
            }
        },
    );
});

describe("validation contract: strict judges the other Part 1 part types too (#506)", { timeout: 120000 }, () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    for (const [segment, part, root] of [
        ["header", "word/header1.xml", "hdr"],
        ["footer", "word/footer1.xml", "ftr"],
    ] as const) {
        it(`a ${segment} part typed only by the generic Default fails strict but stays silent under word-valid`, async () => {
            const parts = {
                ...MINIMAL_PARTS,
                [part]: `<w:${root} xmlns:w="${W}"/>`,
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rId9" Type="${R}/${segment}" Target="${part.replace("word/", "")}"/></Relationships>`,
                ),
            };
            const [wordValid, strict] = await Promise.all([
                validateBytes(await buildDocx(parts), "word-valid"),
                validateBytes(await buildDocx(parts), "strict"),
            ]);
            expect(wordValid.issues.some((issue) => issue.code === "word-part-content-type-generic")).toBe(false);
            expect(strict.issues).toContainEqual(
                expect.objectContaining({ code: "word-part-content-type-generic", severity: "error", path: part }),
            );
            expect(strict.valid).toBe(false);
        });
    }
});

describe("validation contract: the main document's own relationships part is read (#514)", { timeout: 120000 }, () => {
    const withoutRootEndTag = (xml: string): string => xml.slice(0, xml.lastIndexOf("</"));
    const SPECIMENS: readonly (readonly [string, string])[] = [
        ["missing its root end tag", withoutRootEndTag(MINIMAL_PARTS["word/_rels/document.xml.rels"]!)],
        ["zero bytes", ""],
    ];

    it.each(SPECIMENS)("word-valid and strict both fail a document.xml.rels that is %s", { timeout: 120000 }, async (_shape, rels) => {
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx({ ...MINIMAL_PARTS, "word/_rels/document.xml.rels": rels }), profile);
            expect(
                result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/_rels/document.xml.rels")?.severity,
                profile,
            ).toBe("error");
            expect(result.valid, profile).toBe(false);
        }
    });

    it("keeps the well-formed control valid under both profiles", { timeout: 120000 }, async () => {
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(await buildDocx(), profile);
            expect(result.valid, profile).toBe(true);
        }
    });
});

describe("validation contract: header, endnotes, theme and glossaryDocument are read (#516)", { timeout: 120000 }, () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const unclosed = (rootOpen: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootOpen}`;
    const MALFORMED: Readonly<Record<string, string>> = {
        header: unclosed(`<w:hdr xmlns:w="${W}"><w:p>`),
        endnotes: unclosed(`<w:endnotes xmlns:w="${W}">`),
        theme: unclosed('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">'),
        glossary: unclosed(`<w:glossaryDocument xmlns:w="${W}">`),
    };
    const documentUsingHeader = (MINIMAL_PARTS["word/document.xml"] as string)
        .replace('mc:Ignorable="w14">', `mc:Ignorable="w14" xmlns:r="${R}">`)
        .replace("<w:sectPr>", `<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>`);
    const documentUsingEndnote = (MINIMAL_PARTS["word/document.xml"] as string).replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:r><w:t>Hello</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r>',
    );

    /** A part stored outside word/, named by a main-document relationship with its own Override. */
    const outsideWord = async (stored: string, type: string, contentType: string, xml: string, doc?: string): Promise<Uint8Array> => {
        const rels = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${type}" Target="../${stored}"/></Relationships>`,
        );
        const stream = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/${stored}" ContentType="${contentType}"/></Types>`,
        );
        return buildDocx({
            ...MINIMAL_PARTS,
            ...(doc ? { "word/document.xml": doc } : {}),
            [stored]: xml,
            "word/_rels/document.xml.rels": rels,
            "[Content_Types].xml": stream,
        });
    };

    /** The four reachability widenings of session 15: two lazy (used), two eager. */
    const CASES: readonly (readonly [string, string, string, string, string, string | undefined])[] = [
        ["used header", "hdr/header1.xml", `${R}/header`, `${WML}.header+xml`, MALFORMED.header!, documentUsingHeader],
        ["used endnotes part", "en/endnotes.xml", `${R}/endnotes`, `${WML}.endnotes+xml`, MALFORMED.endnotes!, documentUsingEndnote],
        ["theme", "th/theme1.xml", `${R}/theme`, "application/vnd.openxmlformats-officedocument.theme+xml", MALFORMED.theme!, undefined],
        ["glossaryDocument", "g/glossary.xml", `${R}/glossaryDocument`, `${WML}.document.glossary+xml`, MALFORMED.glossary!, undefined],
    ];

    it.each(CASES)(
        "word-valid and strict both fail a malformed %s that Word reads",
        { timeout: 120000 },
        async (_label, stored, type, contentType, xml, doc) => {
            const bytes = await outsideWord(stored, type, contentType, xml, doc);
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === stored)?.severity, profile).toBe(
                    "error",
                );
                expect(result.valid, profile).toBe(false);
            }
        },
    );

    it("word-valid passes an unused malformed word/header1.xml and strict still fails it", { timeout: 120000 }, async () => {
        const rels = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${R}/header" Target="header1.xml"/></Relationships>`,
        );
        const stream = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/header1.xml" ContentType="${WML}.header+xml"/></Types>`,
        );
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/header1.xml": MALFORMED.header!,
            "word/_rels/document.xml.rels": rels,
            "[Content_Types].xml": stream,
        });
        const wordValid = await validateBytes(bytes, "word-valid");
        expect(wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/header1.xml")?.severity).toBe(
            "warning",
        );
        expect(wordValid.valid).toBe(true);
        const strict = await validateBytes(bytes, "strict");
        expect(strict.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/header1.xml")?.severity).toBe("error");
        expect(strict.valid).toBe(false);
    });
});

describe("validation contract: the glossary is a second story root (#539, #549)", { timeout: 120000 }, () => {
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const unclosed = (rootOpen: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootOpen}`;
    const HEADER = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
    const MALFORMED_HEADER = unclosed(`<w:hdr xmlns:w="${W}"><w:p>`);
    // The session-16 `ar-header-ref-in-glossary` shape: the glossary's own
    // relationships part names word/glossary/header9.xml, and its
    // w:docPartBody w:sectPr carries the w:headerReference. The main
    // document reaches the glossary through a glossaryDocument relationship.
    const GLOSSARY =
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart>` +
        `<w:docPartPr><w:name w:val="Part"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
        `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p>` +
        `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:docPartBody>` +
        `</w:docPart></w:docParts></w:glossaryDocument>`;
    const withoutRootEndTag = (xml: string): string => xml.slice(0, xml.lastIndexOf("</"));
    const relsPart = (body: string): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    /** The session-17 glossary: `w:docPartBody` carries `run` in its paragraph, `sectPr` in its `w:sectPr`. */
    const glossaryWith = (run: string, sectPr: string): string =>
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart>` +
        `<w:docPartPr><w:name w:val="Part"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
        `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r>${run}</w:p><w:sectPr>${sectPr}</w:sectPr></w:docPartBody>` +
        `</w:docPart></w:docParts></w:glossaryDocument>`;
    const STORY_SEGMENTS = ["footer", "comments", "footnotes", "endnotes"] as const;
    const STORY_FILE: Readonly<Record<(typeof STORY_SEGMENTS)[number], string>> = {
        footer: "footer9.xml",
        comments: "comments.xml",
        footnotes: "footnotes.xml",
        endnotes: "endnotes.xml",
    };
    const STORY_ROOT: Readonly<Record<(typeof STORY_SEGMENTS)[number], string>> = {
        footer: "ftr",
        comments: "comments",
        footnotes: "footnotes",
        endnotes: "endnotes",
    };
    /** The glossary that uses each lazy story segment, as session 17 references it. */
    const STORY_GLOSSARY: Readonly<Record<(typeof STORY_SEGMENTS)[number], string>> = {
        footer: glossaryWith("", `<w:footerReference w:type="default" r:id="rIdH"/>`),
        comments: glossaryWith(`<w:r><w:commentReference w:id="0"/></w:r>`, ""),
        footnotes: glossaryWith(`<w:r><w:footnoteReference w:id="1"/></w:r>`, ""),
        endnotes: glossaryWith(`<w:r><w:endnoteReference w:id="1"/></w:r>`, ""),
    };
    const STORY_WELLFORMED: Readonly<Record<(typeof STORY_SEGMENTS)[number], string>> = {
        footer: `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`,
        comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`,
        footnotes:
            `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:id="1"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`,
        endnotes:
            `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
            `<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>` +
            `<w:endnote w:id="1"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`,
    };
    /**
     * The session-17 `<segment>-ref-in-glossary` shape: `word/glossary/document.xml`
     * reached through a main-document `glossaryDocument` relationship, whose own
     * relationships part names `word/glossary/<file>` as `rIdH`.
     */
    const storyViaGlossary = (
        segment: (typeof STORY_SEGMENTS)[number],
        xml: string,
        glossaryRels?: string,
        glossary?: string,
    ): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/glossary/document.xml": glossary ?? STORY_GLOSSARY[segment],
            "word/glossary/_rels/document.xml.rels":
                glossaryRels ?? relsPart(`<Relationship Id="rIdH" Type="${R}/${segment}" Target="${STORY_FILE[segment]}"/>`),
            [`word/glossary/${STORY_FILE[segment]}`]: xml,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/>` +
                    `<Override PartName="/word/glossary/${STORY_FILE[segment]}" ContentType="${WML}.${segment}+xml"/></Types>`,
            ),
        });

    const headerViaGlossary = (header: string, glossaryRels?: string, glossary?: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/glossary/document.xml": glossary ?? GLOSSARY,
            "word/glossary/_rels/document.xml.rels":
                glossaryRels ??
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                    `<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>` +
                    `</Relationships>`,
            "word/glossary/header9.xml": header,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/>` +
                    `<Override PartName="/word/glossary/header9.xml" ContentType="${WML}.header+xml"/></Types>`,
            ),
        });

    // Session 19 (#562) and session 20 (#569): the eagerly parsed types a
    // reached glossary's relationships part can carry — stored file, content
    // type and root element (`fontTable` is `w:fonts`, a theme is an
    // `a:theme`), as the s19/s20 builders have them.
    const GLOSSARY_EAGER_SEGMENTS = ["styles", "settings", "fontTable", "webSettings", "numbering", "theme"] as const;
    const GLOSSARY_EAGER: Readonly<
        Record<
            (typeof GLOSSARY_EAGER_SEGMENTS)[number],
            { readonly file: string; readonly open: string; readonly close: string; readonly contentType: string }
        >
    > = {
        styles: { file: "styles.xml", open: `<w:styles xmlns:w="${W}">`, close: "</w:styles>", contentType: `${WML}.styles+xml` },
        settings: { file: "settings.xml", open: `<w:settings xmlns:w="${W}">`, close: "</w:settings>", contentType: `${WML}.settings+xml` },
        fontTable: { file: "fontTable.xml", open: `<w:fonts xmlns:w="${W}">`, close: "</w:fonts>", contentType: `${WML}.fontTable+xml` },
        webSettings: {
            file: "webSettings.xml",
            open: `<w:webSettings xmlns:w="${W}">`,
            close: "</w:webSettings>",
            contentType: `${WML}.webSettings+xml`,
        },
        numbering: {
            file: "numbering.xml",
            open: `<w:numbering xmlns:w="${W}">`,
            close: "</w:numbering>",
            contentType: `${WML}.numbering+xml`,
        },
        theme: {
            file: "theme/theme1.xml",
            open: `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">`,
            close: "</a:theme>",
            contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
        },
    };
    /**
     * The session-19/20 `glossary-<segment>` shape: a glossary that references no
     * story, reached through a main-document `glossaryDocument` relationship,
     * whose own relationships part relates `word/glossary/<file>` as `rIdE`
     * with a matching Override.
     */
    const eagerViaGlossary = (segment: keyof typeof GLOSSARY_EAGER, xml: string): Promise<Buffer> => {
        const { file, contentType } = GLOSSARY_EAGER[segment];
        return buildDocx({
            ...MINIMAL_PARTS,
            "word/glossary/document.xml": glossaryWith("", ""),
            "word/glossary/_rels/document.xml.rels": relsPart(`<Relationship Id="rIdE" Type="${R}/${segment}" Target="${file}"/>`),
            [`word/glossary/${file}`]: xml,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/>` +
                    `<Override PartName="/word/glossary/${file}" ContentType="${contentType}"/></Types>`,
            ),
        });
    };

    it("word-valid and strict both fail a malformed header the glossary references", { timeout: 120000 }, async () => {
        // Word 16.115 offers to repair `ar-header-ref-in-glossary-malformed`
        // (a9d3c660c1ba09c1, session 16): the glossary's headerReference makes
        // the header Word-read, so its xml-syntax error blocks.
        const bytes = await headerViaGlossary(MALFORMED_HEADER);
        for (const profile of ["word-valid", "strict"] as const) {
            const result = await validateBytes(bytes, profile);
            expect(
                result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/header9.xml")?.severity,
                profile,
            ).toBe("error");
            expect(result.valid, profile).toBe(false);
        }
    });

    it("word-valid and a named xml-syntax run agree on the glossary header", { timeout: 120000 }, async () => {
        // #381: a partial run applies the profile the full run applies — the
        // reached set must name the glossary header in both.
        const bytes = await headerViaGlossary(MALFORMED_HEADER);
        const named = await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, bytes);
            return validate(file, { profile: "word-valid", checks: ["xml-syntax"] });
        });
        expect(named.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/header9.xml")?.severity).toBe(
            "error",
        );
        expect(named.valid).toBe(false);
    });

    it("keeps the well-formed glossary-referenced header word-valid (control)", { timeout: 120000 }, async () => {
        // `ar-header-ref-in-glossary-wellformed-control` (acbbb90b2033c15f)
        // opens clean in Word.
        const result = await validateBytes(await headerViaGlossary(HEADER), "word-valid");
        expect(result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/header9.xml")).toBeUndefined();
        expect(result.valid).toBe(true);
    });

    it("keeps endnotes referenced only from a footnote unreached (session 16)", { timeout: 120000 }, async () => {
        // `ar-endnote-ref-in-footnote-malformed` (40c5524a7ee87ebe) opened
        // clean in Word while malformed: Word does not follow every story's
        // references, so the endnotes part stays unreached and warns.
        const footnotes =
            `<w:footnotes xmlns:w="${W}">` +
            `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:id="1"><w:p><w:r><w:t>Footnote</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:footnote>` +
            `</w:footnotes>`;
        const document = MINIMAL_PARTS["word/document.xml"]!.replace(
            "<w:r><w:t>Hello</w:t></w:r>",
            '<w:r><w:footnoteReference w:id="1"/></w:r>',
        );
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": document,
            "word/footnotes.xml": footnotes,
            "word/endnotes.xml": unclosed(`<w:endnotes xmlns:w="${W}">`),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/>` +
                    `<Relationship Id="rIdE" Type="${R}/endnotes" Target="endnotes.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/footnotes.xml" ContentType="${WML}.footnotes+xml"/>` +
                    `<Override PartName="/word/endnotes.xml" ContentType="${WML}.endnotes+xml"/></Types>`,
            ),
        });
        const wordValid = await validateBytes(bytes, "word-valid");
        expect(wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/endnotes.xml")?.severity).toBe(
            "warning",
        );
        expect(wordValid.valid).toBe(true);
        const strict = await validateBytes(bytes, "strict");
        expect(strict.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/endnotes.xml")?.severity).toBe("error");
        expect(strict.valid).toBe(false);
    });

    it.each(STORY_SEGMENTS)(
        "word-valid and strict both fail a malformed %s part the glossary references (#549)",
        { timeout: 120000 },
        async (segment) => {
            // Session 17: `*-ref-in-glossary-malformed` all get repair-warning
            // — the footer, comments, footnotes and endnotes the glossary
            // references are Word-read.
            const bytes = await storyViaGlossary(segment, unclosed(`<w:${STORY_ROOT[segment]} xmlns:w="${W}">`));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${STORY_FILE[segment]}`)
                        ?.severity,
                    profile,
                ).toBe("error");
                expect(result.valid, profile).toBe(false);
            }
        },
    );

    it.each(STORY_SEGMENTS)(
        "keeps a well-formed %s part the glossary references word-valid (#549 control)",
        { timeout: 120000 },
        async (segment) => {
            // Session 17: every `*-ref-in-glossary-wellformed-control` opens
            // clean — the malformed part is what makes Word repair.
            const result = await validateBytes(await storyViaGlossary(segment, STORY_WELLFORMED[segment]), "word-valid");
            expect(
                result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${STORY_FILE[segment]}`),
            ).toBeUndefined();
            expect(result.valid).toBe(true);
        },
    );

    it.each([
        [
            "missing its root end tag",
            withoutRootEndTag(relsPart(`<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>`)),
            undefined,
        ],
        ["zero bytes", "", undefined],
        [
            "missing its root end tag with nothing referenced",
            withoutRootEndTag(relsPart(`<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>`)),
            glossaryWith("", ""),
        ],
    ] as const)(
        "word-valid and strict both fail a glossary relationships part that is %s (#549)",
        { timeout: 120000 },
        async (_shape, glossaryRels, glossary) => {
            // Session 17: `glossary-rels-unclosed`, `glossary-rels-zero-byte` and
            // `glossary-rels-unclosed-unreferenced` all get repair-warning — Word
            // reads a reached glossary's relationships part as eagerly as the
            // main document's (#514).
            const bytes = await headerViaGlossary(HEADER, glossaryRels, glossary);
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/_rels/document.xml.rels")
                        ?.severity,
                    profile,
                ).toBe("error");
                expect(result.valid, profile).toBe(false);
            }
        },
    );

    it("keeps a malformed header the glossary relates but does not reference word-valid (#549)", { timeout: 120000 }, async () => {
        // Session 17: `header-related-from-glossary-unused-malformed` opens
        // clean — related-but-unused stays unreached on a glossary root too.
        const bytes = await headerViaGlossary(MALFORMED_HEADER, undefined, glossaryWith("", ""));
        const wordValid = await validateBytes(bytes, "word-valid");
        expect(wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/header9.xml")?.severity).toBe(
            "warning",
        );
        expect(wordValid.valid).toBe(true);
        const strict = await validateBytes(bytes, "strict");
        expect(strict.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/glossary/header9.xml")?.severity).toBe(
            "error",
        );
        expect(strict.valid).toBe(false);
    });

    it(
        "word-valid and named xml-syntax runs agree on the glossary's relationships part and stories (#549, #381)",
        { timeout: 120000 },
        async () => {
            const relsBytes = await headerViaGlossary(
                HEADER,
                withoutRootEndTag(relsPart(`<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>`)),
            );
            const commentsBytes = await storyViaGlossary("comments", unclosed(`<w:comments xmlns:w="${W}">`));
            for (const [bytes, part] of [
                [relsBytes, "word/glossary/_rels/document.xml.rels"],
                [commentsBytes, "word/glossary/comments.xml"],
            ] as const) {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, "specimen.docx");
                    writeFileSync(file, bytes);
                    for (const selection of [{}, { checks: ["xml-syntax"] }] as const) {
                        const result = await validate(file, { profile: "word-valid", ...selection });
                        const tag = JSON.stringify(selection);
                        expect(result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === part)?.severity, tag).toBe(
                            "error",
                        );
                        expect(result.valid, tag).toBe(false);
                    }
                });
            }
        },
    );

    it.each(GLOSSARY_EAGER_SEGMENTS)(
        "word-valid and strict both fail a malformed %s part a reached glossary relates (#562, #569)",
        { timeout: 120000 },
        async (segment) => {
            // Session 19 (`s19builder.mts`, `glossary-<segment>-malformed`):
            // Word 16.115 offers repair for a reached glossary's malformed
            // styles, settings, fontTable, webSettings or numbering part even
            // when the glossary references no story — its eager relationships
            // are parsed at open like the main document's. Session 20
            // (`s20builder.mts`, `glossary-theme-malformed`): so is its theme.
            const { file, open } = GLOSSARY_EAGER[segment];
            const bytes = await eagerViaGlossary(segment, unclosed(open));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${file}`)?.severity,
                    profile,
                ).toBe("error");
                expect(result.valid, profile).toBe(false);
            }
        },
    );

    it.each(GLOSSARY_EAGER_SEGMENTS)(
        "keeps a well-formed %s part a reached glossary relates word-valid (#562/#569 control)",
        { timeout: 120000 },
        async (segment) => {
            // Session 19: every `glossary-<segment>-wellformed-control` opens
            // clean in Word. Session 20: so does
            // `glossary-theme-wellformed-control`.
            const { file, open, close } = GLOSSARY_EAGER[segment];
            const result = await validateBytes(await eagerViaGlossary(segment, `${open}${close}`), "word-valid");
            expect(result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${file}`)).toBeUndefined();
            expect(result.valid).toBe(true);
        },
    );

    it("word-valid full, quick and named runs agree on a glossary's eager part (#562, #381)", { timeout: 120000 }, async () => {
        const { file, open } = GLOSSARY_EAGER.styles;
        const bytes = await eagerViaGlossary("styles", unclosed(open));
        await withTempDir(async (dir) => {
            const target = path.join(dir, "specimen.docx");
            writeFileSync(target, bytes);
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(target, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${file}`)?.severity,
                    tag,
                ).toBe("error");
                expect(result.valid, tag).toBe(false);
            }
        });
    });

    it("word-valid full, quick and named runs agree on a glossary's theme part (#569, #381)", { timeout: 120000 }, async () => {
        // #381: the theme reach session 20 widened must reach partial runs
        // too — a partial word-valid run applies the profile the full run
        // applies.
        const { file, open } = GLOSSARY_EAGER.theme;
        const bytes = await eagerViaGlossary("theme", unclosed(open));
        await withTempDir(async (dir) => {
            const target = path.join(dir, "specimen.docx");
            writeFileSync(target, bytes);
            for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                const result = await validate(target, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(
                    result.issues.find((issue) => issue.code === "xml-syntax" && issue.path === `word/glossary/${file}`)?.severity,
                    tag,
                ).toBe("error");
                expect(result.valid, tag).toBe(false);
            }
        });
    });

    it.each([
        [
            "customXml item",
            "customXml/item1.xml",
            "customXml",
            "../../customXml/item1.xml",
            "customXml/item1.xml",
            unclosed(`<datastore xmlns="urn:x-custom">`),
            undefined,
        ],
        [
            "nested glossaryDocument",
            "word/glossary/glossary2.xml",
            "glossaryDocument",
            "glossary2.xml",
            "word/glossary/glossary2.xml",
            unclosed(`<w:glossaryDocument xmlns:w="${W}">`),
            `${WML}.document.glossary+xml`,
        ],
    ] as const)(
        "keeps a malformed %s a reached glossary relates word-valid — observed, not read (session 20, #569)",
        { timeout: 120000 },
        async (_label, stored, segment, target, storedName, xml, contentType) => {
            // Session 20: `glossary-customxml-malformed` and
            // `glossary-nested-glossary-malformed` open clean in Word 16.115
            // while malformed — observed, not read — so both stay unreached:
            // word-valid warns instead of erroring, strict still fails.
            const bytes = await buildDocx({
                ...MINIMAL_PARTS,
                "word/glossary/document.xml": glossaryWith("", ""),
                "word/glossary/_rels/document.xml.rels": relsPart(`<Relationship Id="rIdX" Type="${R}/${segment}" Target="${target}"/>`),
                [storedName]: xml,
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
                ),
                "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                    "</Types>",
                    `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/>` +
                        (contentType === undefined ? "" : `<Override PartName="/${stored}" ContentType="${contentType}"/>`) +
                        "</Types>",
                ),
            });
            const wordValid = await validateBytes(bytes, "word-valid");
            expect(wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === stored)?.severity).toBe("warning");
            expect(wordValid.valid).toBe(true);
            const strict = await validateBytes(bytes, "strict");
            expect(strict.issues.find((issue) => issue.code === "xml-syntax" && issue.path === stored)?.severity).toBe("error");
            expect(strict.valid).toBe(false);
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
    const mkfifo = (file: string) =>
        new Promise<void>((resolve, reject) => execFile("mkfifo", [file], (err) => (err ? reject(err) : resolve())));
    // A directory input whose main part's stored file cannot be read at all
    // — a broken symbolic link, or a FIFO (#426).
    const unreadableInput = async (dir: string, kind: "broken-symlink" | "fifo"): Promise<string> => {
        const parts: Record<string, string | Uint8Array> = { ...MINIMAL_PARTS };
        delete parts["word/document.xml"];
        const input = await writeDirectory(path.join(dir, `input-${kind}.docx`), parts, true);
        const stored = path.join(input, storedName("word/document.xml"));
        if (kind === "broken-symlink") await fs.symlink(path.join(dir, "missing-target.xml"), stored);
        else await mkfifo(stored);
        return input;
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

    it("refuses a stored symlink out of the input root before any repair (#425, #576)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
                "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            };
            const input = path.join(dir, "input.docx");
            await fs.mkdir(input, { recursive: true });
            for (const [part, bytes] of Object.entries(parts)) {
                if (part === "word/document.xml") continue;
                await fs.writeFile(path.join(input, storedName(part)), bytes);
            }
            // The main part's stored file is a symlink to a regular file
            // OUTSIDE the input directory: the walk does not follow a link
            // out of the root (#576), so the run refuses it before any part
            // is read and the outside target is never touched.
            const outside = path.join(dir, "outside");
            await fs.mkdir(outside);
            const target = path.join(outside, "document-target.xml");
            await fs.writeFile(target, parts["word/document.xml"]);
            const targetBefore = await fs.readFile(target);
            await fs.symlink(target, path.join(input, storedName("word/document.xml")));

            const actual = await validateResolves(input, { strict: true, autoRepair: true });
            expect(actual.valid).toBe(false);
            expect(actual.issues).toEqual([
                expect.objectContaining({
                    severity: "error",
                    code: "package-extract-failed",
                    message: expect.stringContaining(
                        `Cannot read ${storedName("word/document.xml")}: a symbolic link outside the input directory`,
                    ),
                }),
            ]);
            // The link itself still stands and its outside target is byte-identical.
            expect((await fs.lstat(path.join(input, storedName("word/document.xml")))).isSymbolicLink()).toBe(true);
            expect(await fs.readFile(target)).toEqual(targetBefore);
            expect(targetBefore.toString()).not.toContain('xml:space="preserve"');
            expect(actual.repairs).toBe(0);
            // The other repair never ran either: the run failed before the
            // package view was staged.
            expect(existsSync(path.join(input, storedName("word/_rels/unused.xml.rels")))).toBe(true);
        });
    });

    it("never writes an auto-repair through a stored symlink, reports the refused write-back, and still applies the other repairs (#425)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
                "word/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            };
            const input = path.join(dir, "input.docx");
            await fs.mkdir(input, { recursive: true });
            for (const [part, bytes] of Object.entries(parts)) {
                if (part === "word/document.xml") continue;
                await fs.writeFile(path.join(input, storedName(part)), bytes);
            }
            // The main part's stored file is a symlink to a regular file
            // inside the input directory: reading follows it, so the package
            // view holds the part, but the repair must never reach the target.
            const target = path.join(input, "document-target.xml");
            await fs.writeFile(target, parts["word/document.xml"]);
            const targetBefore = await fs.readFile(target);
            await fs.symlink(target, path.join(input, storedName("word/document.xml")));

            const control = await validate(
                await writeDirectory(
                    path.join(dir, "control.docx"),
                    // The control holds the link's target as a stored file of
                    // its own, so both runs repair the same number of parts.
                    { ...parts, "document-target.xml": parts["word/document.xml"] },
                    true,
                ),
                {
                    strict: true,
                    autoRepair: true,
                },
            );
            const actual = await validateResolves(input, { strict: true, autoRepair: true });
            // The link itself still stands and its target is byte-identical.
            expect((await fs.lstat(path.join(input, storedName("word/document.xml")))).isSymbolicLink()).toBe(true);
            expect(await fs.readFile(target)).toEqual(targetBefore);
            expect(targetBefore.toString()).not.toContain('xml:space="preserve"');
            // The repair applied to the package view either way; only its
            // write-back into the input is refused.
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
            // The other repair — deleting the empty .rels — still lands on
            // its own regular stored file.
            expect(existsSync(path.join(input, storedName("word/_rels/unused.xml.rels")))).toBe(false);
        });
    });

    it("refuses a symlinked directory segment out of the input root before any repair (#425, #576)", async () => {
        await withTempDir(async (dir) => {
            const parts: Parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
            };
            // The whole word/ tree lives outside the input; a symlink inside
            // the input carries the segment, and the walk does not follow a
            // link out of the root (#576).
            const wordDir = path.join(dir, "outside-word");
            await fs.mkdir(path.join(wordDir, "_rels"), { recursive: true });
            await fs.writeFile(path.join(wordDir, "document.xml"), parts["word/document.xml"]);
            await fs.writeFile(path.join(wordDir, "_rels", "document.xml.rels"), parts["word/_rels/document.xml.rels"]);
            await fs.writeFile(path.join(wordDir, "styles.xml"), parts["word/styles.xml"]);
            const outsideBefore = await snapshot(wordDir);

            const input = path.join(dir, "input.docx");
            await fs.mkdir(path.join(input, "_rels"), { recursive: true });
            await fs.writeFile(path.join(input, "[Content_Types].xml"), parts["[Content_Types].xml"]);
            await fs.writeFile(path.join(input, "_rels", ".rels"), parts["_rels/.rels"]);
            await fs.symlink(wordDir, path.join(input, "word"));

            const actual = await validateResolves(input, { strict: true, autoRepair: true });
            // Nothing outside the input directory changed...
            expect(await snapshot(wordDir)).toEqual(outsideBefore);
            // ...and the run refuses the segment before any part is read.
            expect(actual.valid).toBe(false);
            expect(actual.issues).toEqual([
                expect.objectContaining({
                    severity: "error",
                    code: "package-extract-failed",
                    message: expect.stringContaining("Cannot read word: a symbolic link outside the input directory"),
                }),
            ]);
            expect(actual.repairs).toBe(0);
        });
    });

    it.each(["strict", "word-valid", "lenient"] as const)(
        "fails with package-extract-failed naming a stored part that is a broken symlink (%s, #426)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const input = await unreadableInput(dir, "broken-symlink");
                const before = await snapshot(input);
                const result = await validateResolves(input, { profile });
                // The part silently vanished before (#426): the run judged the
                // remaining parts and never said the stored file cannot be read.
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        message: expect.stringContaining(storedName("word/document.xml")),
                    }),
                ]);
                expect(result.xsdWalk.started).toBe(false);
                expect(result.metadata?.schemaCompleted).toBe(false);
                expect(result.repairs).toBe(0);
                expect(await snapshot(input)).toEqual(before);
            });
        },
    );

    it.each(["strict", "lenient"] as const)(
        "fails with package-extract-failed naming a stored FIFO part, without opening it (%s, #426)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const input = await unreadableInput(dir, "fifo");
                // Reading a FIFO would block: completing this test at all is
                // the no-hang assertion.
                const result = await validateResolves(input, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        message: expect.stringContaining(storedName("word/document.xml")),
                    }),
                    // The reason names the kind of entry.
                ]);
                expect(result.issues[0]?.message).toContain("not a regular file");
                expect(result.repairs).toBe(0);
                expect(result.xsdWalk.started).toBe(false);
            });
        },
    );

    it("fails the direct path too when a forward-slash stored part is a broken symlink (#426)", async () => {
        await withTempDir(async (dir) => {
            const parts: Record<string, string | Uint8Array> = { ...MINIMAL_PARTS };
            delete parts["word/document.xml"];
            const input = await writeDirectory(path.join(dir, "input.docx"), parts, false);
            await fs.symlink(path.join(dir, "missing-target.xml"), path.join(input, "word", "document.xml"));
            const result = await validateResolves(input, { strict: true });
            expect(result.valid).toBe(false);
            expect(result.issues).toEqual([
                expect.objectContaining({
                    severity: "error",
                    code: "package-extract-failed",
                    message: expect.stringContaining("word/document.xml"),
                }),
            ]);
        });
    });

    it.each(["broken-symlink", "fifo"] as const)(
        "quick and named partial runs report the unreadable stored file, never more than the full run (%s, #426, #381)",
        async (kind) => {
            await withTempDir(async (dir) => {
                const input = await unreadableInput(dir, kind);
                const full = await validateResolves(input, { profile: "word-valid" });
                expect(full.valid).toBe(false);
                const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                    new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`));
                for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }]) {
                    const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                    expect(partial.valid).toBe(false);
                    expect(partial.issues).toEqual(
                        expect.arrayContaining([
                            expect.objectContaining({
                                severity: "error",
                                code: "package-extract-failed",
                                message: expect.stringContaining(storedName("word/document.xml")),
                            }),
                        ]),
                    );
                    // #381: a partial run never reports an error the full run does not.
                    for (const error of errors(partial)) expect(errors(full)).toContain(error);
                }
            });
        },
    );

    it.each(["strict", "word-valid", "lenient"] as const)(
        "fails with package-extract-failed naming a directory symlink loop as stored, without descending into it (%s, #546)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, `loop-${profile}.docx`), MINIMAL_PARTS, false);
                await fs.symlink(".", path.join(input, "word", "self"));
                const result = await validateResolves(input, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        // The link is named as stored, at one level — not the
                        // 35-level synthetic path the kernel's symlink budget left.
                        message: expect.stringContaining("Cannot read word/self: a symbolic link loop"),
                    }),
                ]);
                expect(result.issues[0]?.message).not.toContain("self/self");
                expect(result.xsdWalk.started).toBe(false);
                expect(result.metadata?.schemaCompleted).toBe(false);
                expect(result.repairs).toBe(0);
            });
        },
        60_000,
    );

    it("names the first stored link that re-enters an ancestor in a mutual directory loop (#546)", async () => {
        await withTempDir(async (dir) => {
            const input = await writeDirectory(path.join(dir, "mutual.docx"), MINIMAL_PARTS, false);
            await fs.mkdir(path.join(input, "a"), { recursive: true });
            await fs.mkdir(path.join(input, "b"), { recursive: true });
            await fs.symlink("../b", path.join(input, "a", "to-b"));
            await fs.symlink("../a", path.join(input, "b", "to-a"));
            const result = await validateResolves(input, { strict: true });
            expect(result.valid).toBe(false);
            const failed = result.issues.filter((issue) => issue.code === "package-extract-failed");
            expect(failed).toHaveLength(1);
            const message = failed[0]?.message ?? "";
            // Each named link is three segments deep, not a budget-length chain.
            expect(message).toContain("a symbolic link loop");
            expect(message).toContain("a/to-b/to-a");
            expect(message).toContain("b/to-a/to-b");
            expect(message).not.toContain("to-b/to-a/to-b");
        });
    }, 60_000);

    it("reports a symlink chain loop as a loop, never as a broken symbolic link (#546)", async () => {
        await withTempDir(async (dir) => {
            const input = await writeDirectory(path.join(dir, "chain.docx"), MINIMAL_PARTS, false);
            await fs.symlink("two", path.join(input, "one"));
            await fs.symlink("one", path.join(input, "two"));
            const result = await validateResolves(input, { strict: true });
            expect(result.valid).toBe(false);
            expect(result.issues[0]?.message).toContain("a symbolic link loop");
            expect(result.issues[0]?.message).not.toContain("a broken symbolic link");
        });
    }, 60_000);

    it("still reports a genuinely broken symbolic link as broken (#546)", async () => {
        await withTempDir(async (dir) => {
            const input = await writeDirectory(path.join(dir, "broken.docx"), MINIMAL_PARTS, false);
            await fs.symlink("../missing.xml", path.join(input, "word", "gone.xml"));
            const result = await validateResolves(input, { strict: true });
            expect(result.valid).toBe(false);
            expect(result.issues).toEqual([
                expect.objectContaining({
                    severity: "error",
                    code: "package-extract-failed",
                    message: expect.stringContaining("Cannot read word/gone.xml: a broken symbolic link"),
                }),
            ]);
        });
    }, 60_000);

    it.each(["quick", "checks"] as const)(
        "partial runs report the symlink loop, never more than the full run (%s, #546, #381)",
        async (mode) => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "loop.docx"), MINIMAL_PARTS, false);
                await fs.symlink(".", path.join(input, "word", "self"));
                const full = await validateResolves(input, { profile: "word-valid" });
                expect(full.valid).toBe(false);
                const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                    new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`));
                const selection = mode === "quick" ? { quick: true } : { checks: ["xml-syntax"] };
                const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                expect(partial.valid).toBe(false);
                expect(partial.issues).toEqual(
                    expect.arrayContaining([
                        expect.objectContaining({
                            severity: "error",
                            code: "package-extract-failed",
                            message: expect.stringContaining("Cannot read word/self: a symbolic link loop"),
                        }),
                    ]),
                );
                // #381: a partial run never reports an error the full run does not.
                for (const error of errors(partial)) expect(errors(full)).toContain(error);
            });
        },
        60_000,
    );

    it("refuses a symlinked directory segment and its target without double-writing one physical file (#544)", async () => {
        const runCase = async (profile: "word-valid" | "strict") => {
            return withTempDir(async (dir) => {
                const input = path.join(dir, "input.docx");
                const real = path.join(input, "real");
                const custom = path.join(input, "custom");
                await fs.mkdir(path.join(input, "_rels"), { recursive: true });
                await fs.mkdir(path.join(real, "_rels"), { recursive: true });
                await fs.mkdir(custom, { recursive: true });
                await fs.writeFile(path.join(input, "[Content_Types].xml"), MINIMAL_PARTS["[Content_Types].xml"]!);
                await fs.writeFile(path.join(input, "_rels", ".rels"), MINIMAL_PARTS["_rels/.rels"]!);
                await fs.writeFile(
                    path.join(real, "document.xml"),
                    MINIMAL_PARTS["word/document.xml"]!.replace("<w:t>Hello</w:t>", "<w:t> Hello</w:t>"),
                );
                await fs.writeFile(path.join(real, "_rels", "document.xml.rels"), MINIMAL_PARTS["word/_rels/document.xml.rels"]!);
                await fs.writeFile(path.join(real, "styles.xml"), MINIMAL_PARTS["word/styles.xml"]!);
                await fs.writeFile(
                    path.join(custom, "item.xml"),
                    '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t> Custom</w:t></root>',
                );
                // Both stored paths name one physical file: word/ is a link to
                // real/, so the walk sees real/document.xml and
                // word/document.xml as separate package entries.
                await fs.symlink("real", path.join(input, "word"), "dir");

                const target = path.join(real, "document.xml");
                const regular = path.join(custom, "item.xml");
                const targetBefore = await fs.readFile(target);
                const regularBefore = await fs.readFile(regular);
                const partial = profile === "word-valid" ? await validateResolves(input, { profile, quick: true }) : undefined;
                const full = await validateResolves(input, { profile });
                const actual = await validateResolves(input, { profile, autoRepair: true });

                // The link remains a link and its physical target is neither
                // rewritten through the link nor rewritten through its direct
                // alias. One physical file therefore receives no duplicate
                // write-back.
                expect((await fs.lstat(path.join(input, "word"))).isSymbolicLink()).toBe(true);
                expect(await fs.readFile(target)).toEqual(targetBefore);
                expect(targetBefore.toString()).not.toContain('xml:space="preserve"');

                // Every stored path for that physical file is refused with
                // the operational error, while a separate repair still lands.
                const refusals = actual.issues.filter((issue) => issue.code === "repair-write-back-failed");
                expect(refusals).toEqual(
                    expect.arrayContaining([
                        expect.objectContaining({
                            severity: "error",
                            code: "repair-write-back-failed",
                            path: "word/document.xml",
                            message: expect.stringContaining("word/document.xml"),
                        }),
                        expect.objectContaining({
                            severity: "error",
                            code: "repair-write-back-failed",
                            path: "real/document.xml",
                            message: expect.stringContaining("real/document.xml"),
                        }),
                    ]),
                );
                expect(actual.valid).toBe(false);
                expect(await fs.readFile(regular)).not.toEqual(regularBefore);
                expect((await fs.readFile(regular)).toString()).toContain('xml:space="preserve"');

                // Strict remains a superset of Word-valid, and a partial run
                // reports no error the corresponding full run omits (#381).
                if (profile === "word-valid") {
                    for (const code of errorCodes(full.issues)) expect(errorCodes(actual.issues).has(code)).toBe(true);
                    for (const code of errorCodes(partial?.issues ?? [])) expect(errorCodes(full.issues).has(code)).toBe(true);
                }
                return errorCodes(actual.issues);
            });
        };

        const wordValidErrors = await runCase("word-valid");
        const strictErrors = await runCase("strict");
        for (const code of wordValidErrors) expect(strictErrors.has(code)).toBe(true);
    }, 60_000);

    it("refuses a hard-linked write-back destination and leaves its outside target untouched (#545)", async () => {
        const runCase = async (profile: "word-valid" | "strict") => {
            return withTempDir(async (dir) => {
                const parts: Parts = {
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:t>Hello</w:t>", "<w:t> Hello</w:t>"),
                    "custom/item.xml":
                        '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t> Custom</w:t></root>',
                };
                const input = await writeDirectory(path.join(dir, "input.docx"), parts, false);
                const outside = path.join(dir, "outside");
                const outsideTarget = path.join(outside, "target.xml");
                await fs.mkdir(outside);
                await fs.writeFile(outsideTarget, parts["word/document.xml"]!);
                await fs.rm(path.join(input, "word", "document.xml"));
                await fs.link(outsideTarget, path.join(input, "word", "document.xml"));

                const stored = path.join(input, "word", "document.xml");
                const regular = path.join(input, "custom", "item.xml");
                const outsideBefore = await fs.readFile(outsideTarget);
                const storedBefore = await fs.readFile(stored);
                const regularBefore = await fs.readFile(regular);
                const partial = profile === "word-valid" ? await validateResolves(input, { profile, quick: true }) : undefined;
                const full = await validateResolves(input, { profile });
                const actual = await validateResolves(input, { profile, autoRepair: true });

                // The hard link's other name lives outside the input tree. It
                // and the stored file keep their original bytes.
                expect(await fs.readFile(outsideTarget)).toEqual(outsideBefore);
                expect(await fs.readFile(stored)).toEqual(storedBefore);
                expect(outsideBefore.toString()).not.toContain('xml:space="preserve"');

                // The refusal names the stored path and says why a
                // multiply-linked inode is not a safe destination.
                const refusals = actual.issues.filter((issue) => issue.code === "repair-write-back-failed");
                expect(refusals).toEqual(
                    expect.arrayContaining([
                        expect.objectContaining({
                            severity: "error",
                            code: "repair-write-back-failed",
                            path: "word/document.xml",
                            message: expect.stringMatching(/word\/document\.xml.*(?:link count|hard link)/i),
                        }),
                    ]),
                );
                expect(actual.valid).toBe(false);

                // A separate, singly-linked repair still lands.
                expect(await fs.readFile(regular)).not.toEqual(regularBefore);
                expect((await fs.readFile(regular)).toString()).toContain('xml:space="preserve"');

                // Strict remains a superset of Word-valid, and a partial run
                // reports no error the corresponding full run omits (#381).
                if (profile === "word-valid") {
                    for (const code of errorCodes(full.issues)) expect(errorCodes(actual.issues).has(code)).toBe(true);
                    for (const code of errorCodes(partial?.issues ?? [])) expect(errorCodes(full.issues).has(code)).toBe(true);
                }
                return errorCodes(actual.issues);
            });
        };

        const wordValidErrors = await runCase("word-valid");
        const strictErrors = await runCase("strict");
        for (const code of wordValidErrors) expect(strictErrors.has(code)).toBe(true);
    }, 60_000);

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

    it.skipIf(runsAsRoot).each(["strict", "word-valid", "lenient"] as const)(
        "fails with package-extract-failed naming a subdirectory it cannot list, not the parts under it (%s, #543)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, `input-${profile}.docx`), MINIMAL_PARTS, false);
                // Every stored file under word/ vanishes when the walk cannot
                // list the directory (#543): the run must say so instead of
                // judging the remaining parts with misleading rels-* findings.
                const wordDir = path.join(input, "word");
                await fs.chmod(wordDir, 0o000);
                let result: Awaited<ReturnType<typeof validate>>;
                try {
                    result = await validateResolves(input, { profile });
                } finally {
                    await fs.chmod(wordDir, 0o755);
                }
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        // The directory itself is named — word/document.xml and
                        // its siblings are only invisible because of it.
                        message: expect.stringContaining("Cannot read word: an unreadable directory"),
                    }),
                ]);
                expect(result.xsdWalk.started).toBe(false);
                expect(result.metadata?.schemaCompleted).toBe(false);
                expect(result.repairs).toBe(0);
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot).each([false, true])(
        "does not let an unreadable subtree vanish when it holds only unreferenced parts (staged=%s, #543)",
        async (backslashes) => {
            await withTempDir(async (dir) => {
                // The staged variant stores the parts under backslash names,
                // so the unreadable subtree is a plain forward-slash directory
                // beside them: `writeDirectory`'s backslash spellings are
                // single file names, not directories.
                const parts: Parts = backslashes ? MINIMAL_PARTS : { ...MINIMAL_PARTS, "hidden/settings.xml": "<w:settings/>" };
                const input = await writeDirectory(path.join(dir, "input.docx"), parts, backslashes);
                if (backslashes) {
                    await fs.mkdir(path.join(input, "hidden"), { recursive: true });
                    await fs.writeFile(path.join(input, "hidden", "settings.xml"), "<w:settings/>");
                }
                const hidden = path.join(input, "hidden");
                await fs.chmod(hidden, 0o000);
                let result: Awaited<ReturnType<typeof validate>>;
                try {
                    // Nothing references hidden/settings.xml, so the vanished
                    // subtree cannot produce findings of its own: without the
                    // failure the run passes as valid (#543).
                    result = await validateResolves(input, { strict: true });
                } finally {
                    await fs.chmod(hidden, 0o755);
                }
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        message: expect.stringContaining("Cannot read hidden: an unreadable directory"),
                    }),
                ]);
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot)(
        "names the input directory itself when it cannot be listed at all (#543)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                await fs.chmod(input, 0o000);
                let result: Awaited<ReturnType<typeof validate>>;
                try {
                    result = await validateResolves(input, { strict: true });
                } finally {
                    await fs.chmod(input, 0o755);
                }
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        message: expect.stringContaining("Cannot read .: an unreadable directory"),
                    }),
                ]);
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot)(
        "quick and named partial runs report the unlistable subdirectory, never more than the full run (#543, #381)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                const wordDir = path.join(input, "word");
                await fs.chmod(wordDir, 0o000);
                try {
                    const full = await validateResolves(input, { profile: "word-valid" });
                    expect(full.valid).toBe(false);
                    const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                        new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`));
                    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }]) {
                        const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                        expect(partial.valid).toBe(false);
                        expect(partial.issues).toEqual(
                            expect.arrayContaining([
                                expect.objectContaining({
                                    severity: "error",
                                    code: "package-extract-failed",
                                    message: expect.stringContaining("Cannot read word: an unreadable directory"),
                                }),
                            ]),
                        );
                        // #381: a partial run never reports an error the full run does not.
                        for (const error of errors(partial)) expect(errors(full)).toContain(error);
                    }
                } finally {
                    await fs.chmod(wordDir, 0o755);
                }
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot).each(["strict", "word-valid", "lenient"] as const)(
        "does not judge an unreadable directory as a stored part name (%s, #573)",
        async (profile) => {
            await withTempDir(async (dir) => {
                for (const name of ["docu%6Dent", "café"]) {
                    const input = await writeDirectory(path.join(dir, `input-${name}.docx`), MINIMAL_PARTS, false);
                    const unreadable = path.join(input, "word", name);
                    await fs.mkdir(unreadable, { recursive: true });
                    await fs.writeFile(path.join(unreadable, "x.xml"), "<x/>");
                    await fs.chmod(unreadable, 0o000);
                    let result: Awaited<ReturnType<typeof validate>>;
                    try {
                        result = await validateResolves(input, { profile });
                    } finally {
                        await fs.chmod(unreadable, 0o755);
                    }
                    expect(result.valid).toBe(false);
                    // The one refusal, and no stored-name finding about a name
                    // no package holds.
                    expect(result.issues).toEqual([
                        expect.objectContaining({
                            severity: "error",
                            code: "package-extract-failed",
                            message: expect.stringContaining(`Cannot read word/${name}: an unreadable directory`),
                        }),
                    ]);
                    expect(result.issues.map((issue) => issue.code)).not.toContain("opc-part-name-escaped-unreserved");
                    expect(result.issues.map((issue) => issue.code)).not.toContain("opc-part-name-word-refused");
                    expect(result.xsdWalk.started).toBe(false);
                    expect(result.metadata?.schemaCompleted).toBe(false);
                    expect(result.repairs).toBe(0);
                }
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot)(
        "still judges the files under a readable directory with the same name (#573)",
        async () => {
            await withTempDir(async (dir) => {
                // The control: the same stored spelling, readable, is a
                // stored file-entry name and is judged as one.
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                const readable = path.join(input, "word", "docu%6Dent");
                await fs.mkdir(readable, { recursive: true });
                await fs.writeFile(path.join(readable, "x.xml"), "<x/>");
                const result = await validateResolves(input, { profile: "word-valid" });
                expect(result.issues).toEqual(
                    expect.arrayContaining([
                        expect.objectContaining({
                            severity: "error",
                            code: "opc-part-name-escaped-unreserved",
                            path: "word/docu%6Dent/x.xml",
                        }),
                    ]),
                );
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot)(
        "quick and named partial runs report the unreadable directory, never more than the full run (#573, #381)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                const unreadable = path.join(input, "word", "café");
                await fs.mkdir(unreadable, { recursive: true });
                await fs.writeFile(path.join(unreadable, "x.xml"), "<x/>");
                await fs.chmod(unreadable, 0o000);
                try {
                    const full = await validateResolves(input, { profile: "word-valid" });
                    expect(full.valid).toBe(false);
                    const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                        new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`));
                    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }]) {
                        const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                        expect(partial.valid).toBe(false);
                        // The one refusal: a partial run judges no stored name
                        // the unreadable directory contributed (#573).
                        expect(partial.issues).toEqual([
                            expect.objectContaining({
                                severity: "error",
                                code: "package-extract-failed",
                                message: expect.stringContaining("Cannot read word/café: an unreadable directory"),
                            }),
                        ]);
                        // #381: a partial run never reports an error the full run does not.
                        for (const error of errors(partial)) expect(errors(full)).toContain(error);
                    }
                } finally {
                    await fs.chmod(unreadable, 0o755);
                }
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot)(
        "orders the unreadable entries by code unit, so the message is the same on every host (#575)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                for (const name of ["ä", "z"]) {
                    const unreadable = path.join(input, "word", name);
                    await fs.mkdir(unreadable, { recursive: true });
                    await fs.writeFile(path.join(unreadable, "x.xml"), "<x/>");
                    await fs.chmod(unreadable, 0o000);
                }
                let result: Awaited<ReturnType<typeof validate>>;
                try {
                    result = await validateResolves(input, { profile: "strict" });
                } finally {
                    for (const name of ["ä", "z"]) await fs.chmod(path.join(input, "word", name), 0o755);
                }
                expect(result.valid).toBe(false);
                const message = result.issues.find((issue) => issue.code === "package-extract-failed")?.message ?? "";
                // Code-unit order: "z" (U+007A) precedes "ä" (U+00E4). The
                // host locale never decides the order of the message.
                expect(message).toContain("Cannot read word/z: an unreadable directory (EACCES); Cannot read word/ä");
                expect(message.indexOf("word/z")).toBeLessThan(message.indexOf("word/ä"));
            });
        },
        60_000,
    );

    it.skipIf(runsAsRoot || process.platform === "win32")(
        "keeps strict a superset of word-valid for a walk refusal, in full and partial runs (#573, #576, #381)",
        async () => {
            const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                new Set(result.issues.filter((i) => i.severity === "error").map((i) => i.code ?? ""));
            await withTempDir(async (dir) => {
                // The unreadable directory (#573) and the link out of the
                // input root (#576): each is one refusal under every profile.
                const unreadableInput = await writeDirectory(path.join(dir, "unreadable.docx"), MINIMAL_PARTS, false);
                const unreadable = path.join(unreadableInput, "word", "café");
                await fs.mkdir(unreadable, { recursive: true });
                await fs.chmod(unreadable, 0o000);
                const linkedInput = await writeDirectory(path.join(dir, "linked.docx"), MINIMAL_PARTS, false);
                const outside = path.join(dir, "outside");
                await fs.mkdir(outside, { recursive: true });
                await fs.symlink(outside, path.join(linkedInput, "word", "ext"));
                try {
                    for (const input of [unreadableInput, linkedInput]) {
                        const word = await validateResolves(input, { profile: "word-valid" });
                        const strict = await validateResolves(input, { profile: "strict" });
                        expect(word.valid).toBe(false);
                        expect(strict.valid).toBe(false);
                        // Strict is a superset of word-valid: every error the
                        // Word-valid profile reports, strict reports too.
                        for (const code of errors(word)) expect(errors(strict)).toContain(code);
                        // #381: a partial run reports no error the full run does not.
                        for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }]) {
                            const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                            expect(partial.valid).toBe(false);
                            for (const code of errors(partial)) expect(errors(word)).toContain(code);
                        }
                    }
                } finally {
                    await fs.chmod(unreadable, 0o755);
                }
            });
        },
        60_000,
    );

    it.skipIf(process.platform === "win32").each(["strict", "word-valid", "lenient"] as const)(
        "fails with package-extract-failed naming a link out of the input root (%s, #576)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, `input-${profile}.docx`), MINIMAL_PARTS, false);
                // The link leaves the input root: its target is not part of
                // this package, so the walk must not follow it (#576).
                const outside = path.join(dir, "outside");
                await fs.mkdir(outside, { recursive: true });
                await fs.writeFile(path.join(outside, "note.txt"), "outside");
                await fs.symlink(outside, path.join(input, "word", "ext"));
                const result = await validateResolves(input, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        message: expect.stringContaining("Cannot read word/ext: a symbolic link outside the input directory"),
                    }),
                ]);
                // The foreign file is never judged as a part of this package.
                expect(result.issues.map((issue) => issue.code)).not.toContain("ct-uncovered-part");
                expect(result.issues.map((issue) => issue.code)).not.toContain("rels-unreferenced");
                expect(result.xsdWalk.started).toBe(false);
                expect(result.metadata?.schemaCompleted).toBe(false);
                expect(result.repairs).toBe(0);
            });
        },
        60_000,
    );

    it.skipIf(process.platform === "win32")(
        "reports a link to an ancestor of the input root at the link's stored name (#576)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                // A sibling specimen beside the input: the walk must not
                // descend into it and must not judge its contents.
                const sibling = path.join(dir, "sibling.docx");
                await writeDirectory(sibling, MINIMAL_PARTS, false);
                await fs.symlink("..", path.join(input, "up"));
                const result = await validateResolves(input, { profile: "word-valid" });
                expect(result.valid).toBe(false);
                expect(result.issues).toEqual([
                    expect.objectContaining({
                        severity: "error",
                        code: "package-extract-failed",
                        // The link's own stored name, not a path under it.
                        message: expect.stringContaining("Cannot read up: a symbolic link outside the input directory"),
                    }),
                ]);
                expect(result.issues[0]?.message).not.toContain("up/sibling.docx");
            });
        },
        60_000,
    );

    it.skipIf(process.platform === "win32")(
        "quick and named partial runs report the link out of the root, never more than the full run (#576, #381)",
        async () => {
            await withTempDir(async (dir) => {
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                const outside = path.join(dir, "outside");
                await fs.mkdir(outside, { recursive: true });
                await fs.writeFile(path.join(outside, "note.txt"), "outside");
                await fs.symlink(outside, path.join(input, "word", "ext"));
                const full = await validateResolves(input, { profile: "word-valid" });
                expect(full.valid).toBe(false);
                const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                    new Set(result.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`));
                for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }]) {
                    const partial = await validateResolves(input, { profile: "word-valid", ...selection });
                    expect(partial.valid).toBe(false);
                    expect(partial.issues).toEqual(
                        expect.arrayContaining([
                            expect.objectContaining({
                                severity: "error",
                                code: "package-extract-failed",
                                message: expect.stringContaining("Cannot read word/ext: a symbolic link outside the input directory"),
                            }),
                        ]),
                    );
                    // #381: a partial run never reports an error the full run does not.
                    for (const error of errors(partial)) expect(errors(full)).toContain(error);
                }
            });
        },
        60_000,
    );

    it.skipIf(process.platform === "win32")(
        "keeps following a link that resolves inside the input root (#576)",
        async () => {
            await withTempDir(async (dir) => {
                // The control: the link resolves inside the root, so the walk
                // follows it and the tree stays one package (#546, #544).
                const input = await writeDirectory(path.join(dir, "input.docx"), MINIMAL_PARTS, false);
                await fs.mkdir(path.join(input, "real"), { recursive: true });
                await fs.writeFile(path.join(input, "real", "note.txt"), "inside");
                await fs.symlink("../real", path.join(input, "word", "inside"));
                const result = await validateResolves(input, { profile: "word-valid" });
                expect(result.issues.map((issue) => issue.code)).not.toContain("package-extract-failed");
            });
        },
        60_000,
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

    // chmod has no effect for root; a read-only source file is the trigger.
    it.skipIf(runsAsRoot)(
        "keeps repair-write-back-failed an error under every profile — a repair the caller asked for and did not get is an operational failure, not a document property (#428)",
        async () => {
            await withTempDir(async (dir) => {
                for (const profile of ["strict", "lenient", "word-valid"] as const) {
                    const parts = {
                        ...MINIMAL_PARTS,
                        "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(">Hello<", "> Hello <"),
                    };
                    const input = await writeDirectory(path.join(dir, `input-${profile}.docx`), parts, true);
                    const readOnly = path.join(input, storedName("word/document.xml"));
                    await fs.chmod(readOnly, 0o444);
                    let actual;
                    try {
                        actual = await validateResolves(input, { profile, autoRepair: true });
                    } finally {
                        await fs.chmod(readOnly, 0o644);
                    }
                    expect(actual.valid, profile).toBe(false);
                    expect(actual.issues, profile).toEqual(
                        expect.arrayContaining([
                            expect.objectContaining({ severity: "error", code: "repair-write-back-failed", path: "word/document.xml" }),
                        ]),
                    );
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

    it.skipIf(runsAsRoot)(
        "strict and word-valid agree a packed input that cannot be read fails as package-open-failed (#422)",
        async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "packed-ro.docx");
                await fs.writeFile(file, await buildDocx());
                await fs.chmod(file, 0o000);
                let wordValid;
                let strict;
                try {
                    wordValid = await validateResolves(file);
                    strict = await validateResolves(file, { strict: true });
                } finally {
                    await fs.chmod(file, 0o644);
                }
                for (const [profile, result] of [
                    ["word-valid", wordValid],
                    ["strict", strict],
                ] as const) {
                    expect(result!.valid, profile).toBe(false);
                    expect(result!.issues, profile).toEqual([expect.objectContaining({ severity: "error", code: "package-open-failed" })]);
                    expect(result!.metadata?.executedChecks, profile).toEqual([]);
                }
            });
        },
    );
});

describe("validation contract: refused host names stay strict failures (#423)", () => {
    it.each([["word/header1.xml", "word/HEADER1.xml"], ["word/caf\u00e9.xml", "word/cafe\u0301.xml"], ["word/x\0y.xml"]])(
        "strict includes Word-valid errors for %j",
        async (...names) => {
            const bytes = buildRawZip([
                ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                ...names.map((name) => ({ name, content: "<w:hdr", utf8: true })),
            ]);
            const word = await validateBytes(bytes, "word-valid");
            const strict = await validateBytes(bytes, "strict");
            expect(word.valid).toBe(false);
            expect(strict.valid).toBe(false);
            expect([...errorCodes(strict.issues)]).toEqual(expect.arrayContaining([...errorCodes(word.issues)]));
        },
        60_000,
    );
});

describe.skipIf(process.platform === "win32")("validation contract: directory refusal includes strict (#424)", () => {
    it.each(["word\\.\\x.xml", "word\\\\x.xml", "word\\a\\..\\x.xml"])(
        "refuses %j under both profiles",
        async (stored) => {
            await withTempDir(async (dir) => {
                const root = path.join(dir, "directory.docx");
                mkdirSync(root);
                for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
                    writeFileSync(path.join(root, name.replaceAll("/", "\\")), content);
                }
                writeFileSync(path.join(root, stored), "<w:document");
                for (const profile of ["word-valid", "strict"] as const) {
                    const result = await validate(root, { profile });
                    expect(result.valid).toBe(false);
                    expect(errorCodes(result.issues)).toContain("package-extract-failed");
                    expect(result.metadata?.executedChecks).not.toContain("xml-syntax");
                }
            });
        },
        60_000,
    );
});

const REFUSED_DIRECTORY_NAME_PARTS = {
    "word/docu%6Dent.xml": "<x/>",
    "word/escaped%2Fslash.xml": "<x/>",
    "word/escaped%21mark.xml": "<x/>",
};
const REFUSED_DIRECTORY_CODES = [
    "opc-equivalent-part-names",
    "opc-part-name-escaped-unreserved",
    "opc-part-name-escaped-slash",
    "opc-part-name-word-refused",
    "package-extract-failed",
];

describe.skipIf(process.platform === "win32")("validation contract: direct directory refusals (#548, #572, #381)", () => {
    it.each([
        ["docx", DOCXSchemaValidator],
        ["pptx", PPTXSchemaValidator],
    ] as const)(
        "%s agrees across direct, full, strict and partial runs",
        async (suffix, Validator) => {
            await withTempDir(async (dir) => {
                const root = path.join(dir, `input.${suffix}`);
                mkdirSync(root);
                for (const [name, content] of Object.entries({ ...MINIMAL_PARTS, ...REFUSED_DIRECTORY_NAME_PARTS })) {
                    writeFileSync(path.join(root, name.replaceAll("/", "\\")), content);
                }
                writeFileSync(path.join(root, "word\\a\\..\\x.xml"), "<x/>");
                const word = await validate(root, { profile: "word-valid" });
                expect(word.valid).toBe(false);
                expect(word.issues.map((issue) => issue.code)).toEqual(REFUSED_DIRECTORY_CODES);
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const full = await validate(root, { profile });
                    expect(full.valid).toBe(false);
                    expect(full.issues.map((issue) => issue.code)).toEqual(REFUSED_DIRECTORY_CODES);
                    if (profile === "strict") {
                        expect([...errorCodes(word.issues)].every((code) => errorCodes(full.issues).has(code))).toBe(true);
                    }
                    await expect(new Validator({ unpackedDir: root, profile }).validate()).resolves.toEqual({
                        valid: false,
                        issues: full.issues,
                    });
                    if (profile === "strict") continue;
                    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }]) {
                        const partial = await validate(root, { profile, ...selection });
                        expect(partial.valid).toBe(false);
                        expect(partial.issues).toEqual(
                            selection.checks?.includes("xml-syntax")
                                ? full.issues.filter((issue) => issue.code === "package-extract-failed")
                                : full.issues,
                        );
                    }
                }
            });
        },
        60_000,
    );
});

describe.skipIf(process.platform === "win32")("validation contract: direct walk refusals (#557, #572, #381)", () => {
    describe.each([
        ["docx", DOCXSchemaValidator],
        ["pptx", PPTXSchemaValidator],
    ] as const)("%s", (suffix, Validator) => {
        for (const defect of ["dangling symlink", "unreadable subdirectory", "symlink loop"]) {
            it.skipIf(defect === "unreadable subdirectory" && process.getuid?.() === 0)(
                `${defect} agrees across direct, full, strict and partial runs`,
                async () => {
                    await withTempDir(async (dir) => {
                        const root = path.join(dir, `input.${suffix}`);
                        for (const [name, content] of Object.entries({ ...MINIMAL_PARTS, ...REFUSED_DIRECTORY_NAME_PARTS })) {
                            mkdirSync(path.dirname(path.join(root, name)), { recursive: true });
                            writeFileSync(path.join(root, name), content);
                        }
                        const customXml = path.join(root, "customXml");
                        if (defect === "dangling symlink") symlinkSync("nowhere.xml", path.join(root, "word/dangling.xml"));
                        else if (defect === "symlink loop") symlinkSync(".", path.join(root, "word/self"));
                        else {
                            mkdirSync(customXml);
                            writeFileSync(path.join(customXml, "item.xml"), "<x/>");
                            chmodSync(customXml, 0o000);
                        }
                        try {
                            const word = await validate(root, { profile: "word-valid" });
                            expect(word.valid).toBe(false);
                            expect(word.issues.map((issue) => issue.code)).toEqual(REFUSED_DIRECTORY_CODES);
                            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                                const full = await validate(root, { profile });
                                expect(full.valid, profile).toBe(false);
                                expect(
                                    full.issues.map((issue) => issue.code),
                                    profile,
                                ).toEqual(REFUSED_DIRECTORY_CODES);
                                if (profile === "strict") {
                                    expect([...errorCodes(word.issues)].every((code) => errorCodes(full.issues).has(code))).toBe(true);
                                }
                                await expect(new Validator({ unpackedDir: root, profile }).validate()).resolves.toEqual({
                                    valid: false,
                                    issues: full.issues,
                                });
                                if (profile === "strict") continue; // Strict rejects partial selectors.
                                for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }]) {
                                    const partial = await validate(root, { profile, ...selection });
                                    expect(partial.valid).toBe(false);
                                    expect(partial.issues).toEqual(
                                        selection.checks?.includes("xml-syntax")
                                            ? full.issues.filter((issue) => issue.code === "package-extract-failed")
                                            : full.issues,
                                    );
                                }
                            }
                        } finally {
                            if (defect === "unreadable subdirectory") chmodSync(customXml, 0o755);
                        }
                    });
                },
                60_000,
            );
        }
    });
});

describe("validation contract: portable collisions include strict (#441)", () => {
    it.each([
        ["docProps/core.xml", "DOCPROPS/CORE.XML"],
        ["word/document.xml", "WORD/DOCUMENT.XML"],
        ["word/a.xml", "WORD/b.xml"],
        ["caf\u00e9/a.xml", "cafe\u0301/b.xml"],
    ])(
        "refuses %j and %j before reading malformed parts",
        async (first, second) => {
            const bytes = buildRawZip([
                ...Object.entries(MINIMAL_PARTS)
                    .filter(([name]) => name !== first)
                    .map(([name, content]) => ({ name, content })),
                { name: first, content: "<x", utf8: true },
                { name: second, content: "<x", utf8: true },
            ]);
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validateBytes(bytes, profile);
                expect(result.valid).toBe(false);
                expect(errorCodes(result.issues)).toContain("package-extract-failed");
                expect(result.metadata?.executedChecks).not.toContain("xml-syntax");
                expect(result.xsdWalk.started).toBe(false);
            }
        },
        60_000,
    );
});

describe("validation contract: generic renamed main parts (#535, #381)", () => {
    it.each(GENERIC_RENAMED_MAIN_PARTS)(
        "rejects $part typed by $contentType in full and selected runs",
        async ({ part, parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const code = "word-part-content-type-generic";
                const word = await validate(file, { profile: "word-valid" });
                const strict = await validate(file, { profile: "strict" });
                expect(word.valid).toBe(false);
                expect(strict.valid && !word.valid, "strict success must imply word-valid success").toBe(false);
                expect(strict.issues).toContainEqual(expect.objectContaining({ code, path: part, severity: "error" }));
                for (const selection of [{ quick: true }, { checks: [code] }, { checks: ["content-types"] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect(partial.valid).toBe(word.valid);
                    expect(partial.issues.filter((issue) => issue.code === code)).toEqual(
                        word.issues.filter((issue) => issue.code === code),
                    );
                    expect([...errorCodes(partial.issues)].every((error) => errorCodes(word.issues).has(error))).toBe(true);
                }
            });
        },
        60_000,
    );
});

describe.each([false, true])("validation contract: byte-rejected main usage, staging fallback=%s (#536, #381)", (fallback) => {
    it.each(BYTE_REJECTED_MAIN_PARTS.slice(0, 2))(
        "keeps $segment errors in full and selected runs",
        async ({ part, parts }) => {
            const entries = Object.entries(parts).map(([name, content]) => ({ name, content }));
            // A filename beyond the filesystem component limit forces archiveFindings
            // after extraction fails, while the in-memory package stays readable.
            if (fallback) entries.push({ name: `word/${"a".repeat(300)}.xml`, content: "<x/>" });
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, buildRawZip(entries));
                const word = await validate(file, { profile: "word-valid" });
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid && !word.valid).toBe(false);
                for (const full of [word, strict]) {
                    expect(full.valid).toBe(false);
                    expect(full.issues).toContainEqual(
                        expect.objectContaining({ code: "xml-encoding", path: "word/document.xml", severity: "error" }),
                    );
                    expect(full.issues).toContainEqual(expect.objectContaining({ code: "xml-syntax", path: part, severity: "error" }));
                }
                expect(errorCodes(word.issues).has("package-extract-failed")).toBe(fallback);
                for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["xml"] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect(partial.valid).toBe(word.valid);
                    expect(partial.issues.filter((issue) => issue.code === "xml-syntax" && issue.path === part)).toEqual(
                        word.issues.filter((issue) => issue.code === "xml-syntax" && issue.path === part),
                    );
                    expect([...errorCodes(partial.issues)].every((error) => errorCodes(word.issues).has(error))).toBe(true);
                }
            });
        },
        60_000,
    );
});

describe("validation contract: numbering/theme generic Default (#571, #432, #535, #381)", () => {
    it.each(GENERIC_NUMBERING_THEME_SPECIMENS)(
        "rejects $segment in word-valid and strict with partial/full agreement",
        async (specimen) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "generic.docx");
                const control = path.join(dir, "exact.docx");
                writeFileSync(file, await buildDocx(genericNumberingThemeParts(specimen)));
                writeFileSync(control, await buildDocx(genericNumberingThemeParts(specimen, true)));
                const code = "word-part-content-type-generic";
                const word = await validate(file, { profile: "word-valid" });
                expect(word.valid).toBe(false);
                expect(errorCodes(word.issues)).toEqual(new Set([code]));
                expect(word.issues).toContainEqual(expect.objectContaining({ code, path: specimen.part, severity: "error" }));
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid).toBe(false);
                // Theme's declarable root already fails ct-undeclared-part;
                // suppress only that same-path duplicate, as for main (#535).
                const strictCode = specimen.segment === "theme" ? "ct-undeclared-part" : code;
                expect(errorCodes(strict.issues)).toEqual(new Set([strictCode]));
                expect(strict.issues).toContainEqual(expect.objectContaining({ code: strictCode, path: specimen.part, severity: "error" }));
                for (const selection of [{ quick: true }, { checks: [code] }, { checks: ["content-types"] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect(partial.valid).toBe(word.valid);
                    expect(errorCodes(partial.issues)).toEqual(errorCodes(word.issues));
                    expect(partial.issues.filter((issue) => issue.code === code)).toEqual(
                        word.issues.filter((issue) => issue.code === code),
                    );
                    expect((await validate(control, { profile: "word-valid", ...selection })).valid).toBe(true);
                }
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, path: specimen.part, severity: "warning" }),
                ]);
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    expect((await validate(control, { profile })).valid, `${specimen.segment} ${profile} control`).toBe(true);
                }
            });
        },
        60_000,
    );
});

// session 22 in docs/plans/evidence/word-probes-2026-10-24.md (#581).
describe("validation contract: empty main-document content type (#581, #381)", () => {
    const mainType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const stream = MINIMAL_PARTS["[Content_Types].xml"]!;
    const findings = (result: Awaited<ReturnType<typeof validate>>) =>
        result.issues.filter((issue) => issue.code === "word-content-type-invalid");

    it("blocks word-valid and strict, leaves lenient to its XSD pass, and agrees across full/quick/named Word-valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "empty-main.docx");
            writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "[Content_Types].xml": stream.replace(mainType, "") }));
            const full = await validate(file);
            expect(full.valid).toBe(false);
            expect(findings(full)).toEqual([
                expect.objectContaining({
                    severity: "error",
                    path: "[Content_Types].xml",
                    message: expect.stringContaining("empty content type '' for word/document.xml"),
                }),
            ]);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(findings(strict)).toEqual(findings(full));
            expect(errorCodes(strict.issues)).toContain("xsd-error");
            // Lenient runs no Word rule, fully or partially (#381); its XSD
            // pass already rejects the empty value.
            const lenient = await validate(file, { profile: "lenient" });
            expect(findings(lenient)).toEqual([]);
            expect(lenient.valid).toBe(false);
            expect(errorCodes(lenient.issues)).toContain("xsd-error");
            expect(findings(await validate(file, { profile: "lenient", quick: true }))).toEqual([]);
            for (const selection of [{ quick: true }, { checks: ["word-content-type-invalid"] }, { checks: ["content-types"] }]) {
                const partial = await validate(file, selection);
                expect(partial.valid).toBe(false);
                expect(findings(partial)).toEqual(findings(full));
            }
        });
    }, 60_000);

    it.each([
        ["proper Override control", stream, true],
        ["empty styles Override", stream.replace("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", ""), false],
        ["empty xml Default", stream.replace('ContentType="application/xml"', 'ContentType=""'), false],
        [
            "main typed by empty Default",
            stream
                .replace(`<Override PartName="/word/document.xml" ContentType="${mainType}"/>`, "")
                .replace('ContentType="application/xml"', 'ContentType=""'),
            false,
        ],
    ])(
        "keeps %s outside the new finding",
        async (_label, contentTypes, strictValid) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.docx");
                writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "[Content_Types].xml": contentTypes }));
                for (const selection of [{}, { quick: true }, { checks: ["word-content-type-invalid"] }]) {
                    const result = await validate(file, selection);
                    expect(findings(result)).toEqual([]);
                    expect(result.valid).toBe(true);
                }
                const strict = await validate(file, { profile: "strict" });
                expect(findings(strict)).toEqual([]);
                expect(strict.valid).toBe(strictValid);
                if (!strictValid) expect(errorCodes(strict.issues)).toContain("xsd-error");
            });
        },
        60_000,
    );
});
