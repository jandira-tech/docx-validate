/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word 16.115 reads each ZIP entry's central-directory name and ignores the
 * local file header's name and the Info-ZIP unicode-path extra field
 * (0x7075). The package view and extraction are keyed by the same names
 * (#445), in full, quick and named runs. A `..` that lives only in the local
 * header or only in 0x7075 is no refusal; a central `..` still is. Two
 * central entries under the identical spelling keep "the last stored entry
 * is the one read" true (#446). Evidence: review Y's specimens in
 * docs/plans/evidence/word-probes-2026-10-24.md, session 4d.
 */

import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { zipCentralDirectoryEntries } from "../src/lib/zip-central-directory";
import { extractPackageParts } from "../src/lib/zip-path";
import { centralKeyedArchive, loadCentralNamedArchive } from "../src/lib/zip-central-names";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip, type RawZipEntry } from "./support/raw-zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const SECOND_DOCUMENT =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;

const minimalEntries = (): RawZipEntry[] => Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content }));

// review Y `ry-local-dotdot`: a `..` that exists only in the local header.
const LOCAL_DOTDOT = buildRawZip([...minimalEntries(), { name: "word/extra.xml", localName: "../evil.xml", content: "<x/>" }]);

// review Y `ry-upath-injects-dotdot`: a `..` that exists only in 0x7075.
const UPATH_INJECTS = buildRawZip([...minimalEntries(), { name: "x.xml", unicodePath: "../evil.xml", content: "<x/>" }]);

// review Y `ry-central-dotdot-local-clean`: the `..` is in the central name.
const CENTRAL_DOTDOT = buildRawZip([...minimalEntries(), { name: "evil/../x.xml", localName: "x.xml", content: "<x/>" }]);

// review Y `ry-upath-hides-dotdot`: 0x7075 hides the central `..`.
const UPATH_HIDES = buildRawZip([...minimalEntries(), { name: "a/../x.xml", unicodePath: "x.xml", content: "<x/>" }]);

// review Y `ry-dup-central-distinct-local`: two central entries under one
// spelling, with distinct local names and distinct content.
const DUP_DISTINCT_LOCAL_ENTRIES: readonly RawZipEntry[] = [
    ...minimalEntries().map((entry) => (entry.name === "word/document.xml" ? { ...entry, localName: "AAAA-first.xml" } : entry)),
    { name: "word/document.xml", localName: "BBBB-second.xml", content: SECOND_DOCUMENT },
];
const DUP_DISTINCT_LOCAL = buildRawZip(DUP_DISTINCT_LOCAL_ENTRIES);

const specimen = (dir: string, name: string, bytes: Uint8Array): string => {
    const file = path.join(dir, name);
    writeFileSync(file, bytes);
    return file;
};

const codes = (result: { issues: readonly ValidationIssue[] }, severity?: string): string[] =>
    result.issues.filter((issue) => severity === undefined || issue.severity === severity).map((issue) => issue.code ?? "");

/** Error findings as `code@path`, for the #381 subset invariant. */
const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    [...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code ?? ""}@${issue.path ?? ""}`))].sort();

describe("central and package name decoding (#480)", () => {
    const cp437 = Uint8Array.from([0x63, 0x61, 0x66, 0x82, 0x2e, 0x78, 0x6d, 0x6c]);

    it.each([
        { label: "as-given", localName: undefined, unicodePath: undefined },
        { label: "re-keyed local name", localName: "other.xml", unicodePath: undefined },
        { label: "re-keyed Unicode Path", localName: undefined, unicodePath: "unicode.xml" },
    ])("uses the stored CP437 spelling in the $label package view", async ({ localName, unicodePath }) => {
        const bytes = buildRawZip([{ name: "café.xml", nameBytes: cp437, localName, unicodePath, content: "<x/>" }]);
        const stored = zipCentralDirectoryEntries(bytes)!.map((entry) => entry.name);
        expect(stored).toEqual(["café.xml"]);
        const zip = await loadCentralNamedArchive(bytes);
        expect(Object.keys(zip.files)).toEqual(stored);
        expect(await zip.file("café.xml")!.async("string")).toBe("<x/>");
        const keyed = centralKeyedArchive(bytes);
        expect(zipCentralDirectoryEntries(keyed)!.map((entry) => entry.name)).toEqual(stored);
    });

    it("uses one CP437 spelling for every finding on the part", async () => {
        const bytes = buildRawZip([...minimalEntries(), { name: "café.xml", nameBytes: cp437, content: "<x/>" }]);
        await withTempDir(async (dir) => {
            const result = await validate(specimen(dir, "cp437.docx", bytes), { profile: "word-valid" });
            const findings = result.issues.filter((issue) => issue.path?.startsWith("caf"));
            expect(findings.map((issue) => issue.path)).toEqual(["café.xml", "café.xml"]);
            expect(findings).toEqual(
                expect.arrayContaining([
                    expect.objectContaining({ code: "rels-unreferenced", severity: "warning", path: "café.xml" }),
                    expect.objectContaining({ code: "opc-part-name-word-refused", severity: "error", path: "café.xml" }),
                ]),
            );
            expect(JSON.stringify(result.issues)).not.toContain("caf�.xml");
        });
    });

    it.each([undefined, "other.xml"])("keeps bit-11 names UTF-8 with local name %s", async (localName) => {
        const bytes = buildRawZip([{ name: "café.xml", utf8: true, localName, content: "<x/>" }]);
        expect(Object.keys((await loadCentralNamedArchive(bytes)).files)).toEqual(["café.xml"]);
    });

    it("keeps JSZip's valid central Unicode Path preference when re-keying is unavailable", async () => {
        // An entry whose local-header offset lives in a ZIP64 extra field
        // stays as given: the decoder must not override JSZip's valid 0x7075
        // field. Re-keyable archives deliberately strip it (#445), ZIP64
        // included (#479).
        const bytes = buildRawZip([{ name: "café.xml", nameBytes: cp437, unicodePath: "unicode-é.xml", zip64LocalOffset: true }], {
            zip64: true,
        });
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(Object.keys((await loadCentralNamedArchive(bytes)).files)).toEqual(["unicode-é.xml"]);
    });
});

describe("the package view reads central-directory names (#445, #446)", { timeout: 120_000 }, () => {
    describe.each([
        { name: "local-dotdot", bytes: LOCAL_DOTDOT, unreferenced: "word/extra.xml" },
        { name: "upath-injects-dotdot", bytes: UPATH_INJECTS, unreferenced: "x.xml" },
    ])("$name: the extra name lives outside the central directory", ({ name, bytes, unreferenced }) => {
        it.each(["docx", "docm"] as const)("is valid under word-valid .%s, in a full run", async (suffix) => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.${suffix}`, bytes);
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid, JSON.stringify(result.issues)).toBe(true);
                // The extra part is unreferenced: tolerated, a warning only.
                expect(result.issues).toContainEqual(
                    expect.objectContaining({ code: "rels-unreferenced", severity: "warning", path: unreferenced }),
                );
                expect(codes(result)).not.toContain("package-extract-failed");
                expect(codes(result)).not.toContain("package-open-failed");
            });
        });

        it("is valid in word-valid quick and package-group runs (#381)", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                const full = await validate(file, { profile: "word-valid" });
                for (const selection of [{ quick: true }, { checks: ["package"] }] as const) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    const tag = JSON.stringify(selection);
                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            });
        });

        it("judges the clean central name under strict: no package refusal", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                const result = await validate(file, { profile: "strict" });
                // The extra part is unreferenced: an error under strict — but
                // the refusal must not come from the ignored local name.
                expect(codes(result)).not.toContain("package-extract-failed");
                expect(codes(result)).not.toContain("package-open-failed");
                expect(result.issues).toContainEqual(
                    expect.objectContaining({ code: "rels-unreferenced", severity: "error", path: unreferenced }),
                );
            });
        });

        it("keeps strict refusing partial runs", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                for (const selection of [{ quick: true }, { checks: ["package"] }] as const) {
                    await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toBeInstanceOf(
                        ModeConflictError,
                    );
                }
            });
        });

        it("keeps the lenient package view, without the refusal", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                const full = await validate(file, { profile: "lenient" });
                expect(codes(full)).not.toContain("package-extract-failed");
                expect(full.issues).toContainEqual(
                    expect.objectContaining({ code: "rels-unreferenced", severity: "error", path: unreferenced }),
                );
                const quick = await validate(file, { profile: "lenient", quick: true });
                expect(quick.valid, JSON.stringify(quick.issues)).toBe(true);
                const named = await validate(file, { profile: "lenient", checks: ["package"] });
                expect(named.valid, JSON.stringify(named.issues)).toBe(true);
                expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(named.issues)));
            });
        });
    });

    describe.each([
        { name: "central-dotdot-local-clean", bytes: CENTRAL_DOTDOT },
        { name: "upath-hides-dotdot", bytes: UPATH_HIDES },
    ])("$name: the central name holds the '..'", ({ name, bytes }) => {
        it.each(["word-valid", "strict", "lenient"] as const)("stays refused under %s, in a full run", async (profile) => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(codes(result)).toContain("package-extract-failed");
            });
        });

        it.each(["docx", "docm"] as const)("stays refused under word-valid .%s", async (suffix) => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.${suffix}`, bytes);
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(false);
                expect(codes(result)).toContain("package-extract-failed");
            });
        });

        it("stays refused in word-valid and lenient quick and package-group runs", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, `${name}.docx`, bytes);
                for (const profile of ["word-valid", "lenient"] as const) {
                    for (const selection of [{ quick: true }, { checks: ["package"] }] as const) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(partial.valid, `${profile} ${JSON.stringify(selection)}`).toBe(false);
                        expect(codes(partial)).toContain("package-extract-failed");
                    }
                }
            });
        });
    });

    describe("dup-central-distinct-local: two central entries under one spelling", () => {
        it.each(["docx", "docm"] as const)(
            "stays invalid under word-valid .%s on opc-equivalent-part-names, and the message holds (#446)",
            async (suffix) => {
                await withTempDir(async (dir) => {
                    const file = specimen(dir, `dup-central-distinct-local.${suffix}`, DUP_DISTINCT_LOCAL);
                    const result = await validate(file, { profile: "word-valid" });
                    expect(result.valid).toBe(false);
                    const equivalent = result.issues.filter((issue) => issue.code === "opc-equivalent-part-names");
                    expect(equivalent).toHaveLength(1);
                    expect(equivalent[0]).toMatchObject({ severity: "error", path: "word/document.xml" });
                    expect(equivalent[0]?.message).toContain("The last stored entry is the one read.");
                    // The package view really does hold word/document.xml now:
                    // no main-part finding, no broken relationship.
                    expect(codes(result)).not.toContain("package-extract-failed");
                    expect(codes(result)).not.toContain("word-main-document-part");
                    expect(codes(result)).not.toContain("rels-broken");
                });
            },
        );

        it("reports the same error under strict and a warning under lenient", async () => {
            await withTempDir(async (dir) => {
                const file = specimen(dir, "dup-central-distinct-local.docx", DUP_DISTINCT_LOCAL);
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid).toBe(false);
                expect(strict.issues).toContainEqual(expect.objectContaining({ code: "opc-equivalent-part-names", severity: "error" }));
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues).toContainEqual(expect.objectContaining({ code: "opc-equivalent-part-names", severity: "warning" }));
                expect(codes(lenient)).not.toContain("package-extract-failed");
            });
        });

        it("holds the last central entry's bytes under the shared name, in the view and on extraction", async () => {
            await withTempDir(async (dir) => {
                const context = await ValidationContext.fromArchive(DUP_DISTINCT_LOCAL);
                expect(context.storedNames.filter((name) => name === "word/document.xml")).toHaveLength(2);
                const held = context.bytes("word/document.xml");
                expect(held).toBeDefined();
                expect(new TextDecoder().decode(held)).toBe(SECOND_DOCUMENT);
                expect(context.bytes("AAAA-first.xml")).toBeUndefined();
                expect(context.bytes("BBBB-second.xml")).toBeUndefined();

                const zip = await loadCentralNamedArchive(DUP_DISTINCT_LOCAL);
                await extractPackageParts(zip, dir);
                const staged = path.join(dir, "word", "document.xml");
                expect(existsSync(staged)).toBe(true);
                expect(readFileSync(staged, "utf8")).toBe(SECOND_DOCUMENT);
            });
        });
    });

    describe("centralKeyedArchive", () => {
        it("returns an ordinary archive's bytes unchanged", async () => {
            const bytes = await buildDocx();
            const normalized = centralKeyedArchive(bytes);
            expect(normalized).toBe(bytes);
            const zip = await loadCentralNamedArchive(normalized);
            expect(zip.file("word/document.xml")).not.toBeNull();
        });

        it("leaves malformed, sentinel-offset and truncated bytes alone", async () => {
            for (const bytes of [
                Buffer.from("not a zip"),
                // An entry whose local-header offset lives in a ZIP64 extra field
                // (#479): such entries keep their bytes.
                buildRawZip(
                    [...minimalEntries(), { name: "word/extra.xml", localName: "../evil.xml", content: "<x/>", zip64LocalOffset: true }],
                    { zip64: true },
                ),
                (await buildDocx()).subarray(0, 40),
            ]) {
                expect(centralKeyedArchive(bytes)).toBe(bytes);
            }
        });

        it("bails when the walked directory ends before the record, leaving mismatched names to JSZip (#478)", async () => {
            // A local name JSZip would show instead of the central one, plus
            // a second record whose declared size covers the first. The walk
            // ends at the original record, so `cdEnd !== eocdOffset` and the
            // bytes come back untouched: the gap is what JSZip reads.
            const archive = buildRawZip([...minimalEntries(), { name: "word/extra.xml", localName: "alias.xml", content: "<x/>" }]);
            const eocd = archive.length - 22;
            const appended = Buffer.from(archive.subarray(eocd));
            appended.writeUInt32LE(archive.readUInt32LE(eocd + 12) + 22, 12);
            const bytes = Buffer.concat([archive, appended]);
            expect(await loadedNames(await JSZip.loadAsync(bytes))).toContain("alias.xml");
            expect(centralKeyedArchive(bytes)).toBe(bytes);
            expect(await loadedNames(await loadCentralNamedArchive(bytes))).toContain("alias.xml");
        });

        it("does not modify the input bytes it rewrites", () => {
            const copy = Uint8Array.from(LOCAL_DOTDOT);
            const normalized = centralKeyedArchive(LOCAL_DOTDOT);
            expect(normalized).not.toBe(LOCAL_DOTDOT);
            expect(Uint8Array.from(LOCAL_DOTDOT)).toEqual(copy);
        });
    });
});

// Bytes before the archive: JSZip rebases every stored offset past them
// (`reader.zero`), so the archive it loads is the one behind the prefix.
const PREFIX = Buffer.alloc(42, 0x2a);
// review AD S1: junk before the local-dotdot package.
const PREFIXED_LOCAL_DOTDOT = Buffer.concat([PREFIX, LOCAL_DOTDOT]);
// review AD S2: a whole ZIP, with mismatched names of its own, before it.
const ZIP_A = buildRawZip([{ name: "aa.txt", localName: "bb.txt", content: "a" }]);
const ZIP_PREFIXED_LOCAL_DOTDOT = Buffer.concat([ZIP_A, LOCAL_DOTDOT]);
// review AD S3: an APPNOTE 4.3.13 signature record between the central
// directory and its EOCD, offsets untouched. JSZip reads the gap as a prefix
// and finds no directory, so it refuses the archive.
const withRecordBeforeEocd = (bytes: Uint8Array): Buffer => {
    const buf = Buffer.from(bytes);
    const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    const record = Buffer.from([0x50, 0x4b, 0x05, 0x05, 6, 0, 1, 2, 3, 4, 5, 6]);
    return Buffer.concat([buf.subarray(0, eocd), record, buf.subarray(eocd)]);
};
const GAP_LOCAL_DOTDOT = withRecordBeforeEocd(LOCAL_DOTDOT);

const loadedNames = async (zip: JSZip): Promise<string[]> => Object.keys(zip.files).sort();

describe("the central-keyed view reads the archive JSZip reads (#467, #468, #469)", { timeout: 120_000 }, () => {
    it.each([
        { name: "a junk prefix", bytes: PREFIXED_LOCAL_DOTDOT },
        { name: "a prepended ZIP", bytes: ZIP_PREFIXED_LOCAL_DOTDOT },
    ])("keys the live archive behind $name by its central names", async ({ bytes }) => {
        // JSZip itself loads the live archive under its local name.
        expect(await loadedNames(await JSZip.loadAsync(bytes))).toContain("evil.xml");
        const names = await loadedNames(await loadCentralNamedArchive(bytes));
        expect(names).toContain("word/extra.xml");
        expect(names).toContain("word/document.xml");
        for (const name of ["evil.xml", "aa.txt", "bb.txt"]) expect(names).not.toContain(name);
    });

    it.each([
        { name: "a junk prefix", bytes: PREFIXED_LOCAL_DOTDOT, prefix: PREFIX.length },
        { name: "a prepended ZIP", bytes: ZIP_PREFIXED_LOCAL_DOTDOT, prefix: ZIP_A.length },
    ])("keeps $name verbatim, still a prefix", ({ bytes, prefix }) => {
        const normalized = centralKeyedArchive(bytes);
        expect(normalized).not.toBe(bytes);
        expect(Buffer.from(normalized.subarray(0, prefix))).toEqual(Buffer.from(bytes.subarray(0, prefix)));
        // The stored offsets stay relative to the archive: the rewritten EOCD
        // still puts its directory `prefix` bytes before where it sits.
        const buf = Buffer.from(normalized);
        const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
        expect(eocd - (buf.readUInt32LE(eocd + 16) + buf.readUInt32LE(eocd + 12))).toBe(prefix);
    });

    it.each([
        { name: "a junk prefix", bytes: PREFIXED_LOCAL_DOTDOT },
        { name: "a prepended ZIP", bytes: ZIP_PREFIXED_LOCAL_DOTDOT },
    ])("does not refuse the package for a local-only `..` behind $name", async ({ bytes }) => {
        await withTempDir(async (dir) => {
            const file = specimen(dir, "prefixed.docx", bytes);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                expect(codes(full), profile).not.toContain("package-extract-failed");
                expect(
                    full.issues.filter((issue) => issue.message.includes("evil.xml")),
                    profile,
                ).toEqual([]);
                if (profile === "strict") continue;
                for (const selection of [{ quick: true }, { checks: ["package"] }] as const) {
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    const partial = await validate(file, { profile, ...selection });
                    expect(codes(partial), tag).not.toContain("package-extract-failed");
                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(
                        errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair)),
                        tag,
                    ).toBe(true);
                }
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("never loads an archive JSZip refuses: a record between the directory and the EOCD", async () => {
        await expect(JSZip.loadAsync(GAP_LOCAL_DOTDOT)).rejects.toThrow();
        await expect(loadCentralNamedArchive(GAP_LOCAL_DOTDOT)).rejects.toThrow();
        await withTempDir(async (dir) => {
            const file = specimen(dir, "gap.docx", GAP_LOCAL_DOTDOT);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(false);
                expect(codes(result, "error"), profile).toContain("package-open-failed");
            }
        });
    });

    it("rewrites only the local name and the central extra field: JSZip never reads the local one", () => {
        // `x.xml` both centrally and locally, with 0x7075 in both records:
        // only the central field steers JSZip (`readCentralPart`), so only it
        // is dropped, and the local record keeps its bytes.
        const normalized = centralKeyedArchive(UPATH_INJECTS);
        const field = 9 + "../evil.xml".length;
        expect(normalized.length).toBe(UPATH_INJECTS.length - field);
        // The local name, then the local 0x7075 id: the first record holding both.
        const local = Buffer.from(UPATH_INJECTS).indexOf(Buffer.from([...Buffer.from("x.xml"), 0x75, 0x70]));
        expect(local).toBeGreaterThan(0);
        expect(Buffer.from(normalized.subarray(0, local + 5 + field))).toEqual(Buffer.from(UPATH_INJECTS.subarray(0, local + 5 + field)));
    });
});

// #479: a ZIP64 package. The 32-bit EOCD's directory fields are saturated, so
// JSZip reads the directory through the ZIP64 record and locator; the rewrite
// must rebase those with the directory, or the package view stays keyed by the
// local names — the #445/#446 state.
const LOCAL_DOTDOT_ZIP64 = buildRawZip([...minimalEntries(), { name: "word/extra.xml", localName: "../evil.xml", content: "<x/>" }], {
    zip64: true,
});
const DUP_DISTINCT_LOCAL_ZIP64 = buildRawZip(DUP_DISTINCT_LOCAL_ENTRIES, { zip64: true });
const PREFIXED_LOCAL_DOTDOT_ZIP64 = Buffer.concat([PREFIX, LOCAL_DOTDOT_ZIP64]);
const SENTINEL_LOCAL_DOTDOT_ZIP64 = buildRawZip(
    [...minimalEntries(), { name: "word/extra.xml", localName: "../evil.xml", content: "<x/>", zip64LocalOffset: true }],
    { zip64: true },
);

/** The minimal parts plus the re-keyed extra part, as (name, content) pairs. */
const expectedParts = (): readonly (readonly [string, string])[] => [...Object.entries(MINIMAL_PARTS), ["word/extra.xml", "<x/>"]];

describe("the central-keyed view handles ZIP64 (#479)", { timeout: 120_000 }, () => {
    it("returns a ZIP64 archive with agreeing names and no 0x7075 unchanged", () => {
        const bytes = buildRawZip(minimalEntries(), { zip64: true });
        expect(centralKeyedArchive(bytes)).toBe(bytes);
    });

    it("holds the central names, byte-exact under each", async () => {
        const keyed = centralKeyedArchive(LOCAL_DOTDOT_ZIP64);
        expect(keyed).not.toBe(LOCAL_DOTDOT_ZIP64);
        const zip = await loadCentralNamedArchive(LOCAL_DOTDOT_ZIP64);
        const names = Object.keys(zip.files).sort();
        expect(names).toContain("word/extra.xml");
        expect(names).toContain("word/document.xml");
        expect(names).not.toContain("evil.xml");
        for (const [name, content] of expectedParts()) {
            expect(await zip.file(name)!.async("uint8array"), name).toEqual(new TextEncoder().encode(content));
        }
        expect(zipCentralDirectoryEntries(keyed)!.map((entry) => entry.name)).toEqual(
            zipCentralDirectoryEntries(LOCAL_DOTDOT_ZIP64)!.map((entry) => entry.name),
        );
    });

    it.each(["word-valid", "lenient", "strict"] as const)("does not refuse a local-only `..` under $profile", async (profile) => {
        await withTempDir(async (dir) => {
            const file = specimen(dir, "local-dotdot-zip64.docx", LOCAL_DOTDOT_ZIP64);
            const result = await validate(file, { profile });
            expect(codes(result), profile).not.toContain("package-extract-failed");
            expect(codes(result), profile).not.toContain("package-open-failed");
            expect(
                result.issues.filter((issue) => issue.message.includes("evil.xml")),
                profile,
            ).toEqual([]);
            if (profile !== "word-valid") return;
            expect(result.valid, JSON.stringify(result.issues)).toBe(true);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "rels-unreferenced", severity: "warning", path: "word/extra.xml" }),
            );
        });
    });

    it("reports opc-equivalent-part-names for two central entries under one spelling (#446)", async () => {
        await withTempDir(async (dir) => {
            const file = specimen(dir, "dup-central-distinct-local-zip64.docx", DUP_DISTINCT_LOCAL_ZIP64);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const equivalent = result.issues.filter((issue) => issue.code === "opc-equivalent-part-names");
            expect(equivalent).toHaveLength(1);
            expect(equivalent[0]).toMatchObject({ severity: "error", path: "word/document.xml" });
            expect(codes(result)).not.toContain("package-extract-failed");
        });
        const context = await ValidationContext.fromArchive(DUP_DISTINCT_LOCAL_ZIP64);
        expect(new TextDecoder().decode(context.bytes("word/document.xml"))).toBe(SECOND_DOCUMENT);
        expect(context.bytes("AAAA-first.xml")).toBeUndefined();
        expect(context.bytes("BBBB-second.xml")).toBeUndefined();
    });

    it("keeps a prefix a prefix, re-keying the ZIP64 archive behind it (#467)", async () => {
        const bytes = PREFIXED_LOCAL_DOTDOT_ZIP64;
        const normalized = centralKeyedArchive(bytes);
        expect(normalized).not.toBe(bytes);
        expect(Buffer.from(normalized.subarray(0, PREFIX.length))).toEqual(PREFIX);
        const zip = await loadCentralNamedArchive(bytes);
        expect(Object.keys(zip.files)).toContain("word/extra.xml");
        expect(Object.keys(zip.files)).not.toContain("evil.xml");
        expect(await zip.file("word/document.xml")!.async("string")).toBe(MINIMAL_PARTS["word/document.xml"]);
        // The ZIP64 record's directory size and offset and the locator's
        // record offset are rebased like the 32-bit EOCD's fields: the record
        // and locator still end exactly 76 bytes before the EOCD, and every
        // stored offset stays relative to the archive behind the prefix.
        const buf = Buffer.from(normalized);
        const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
        const locator = eocd - 20;
        const record = Number(buf.readBigUInt64LE(locator + 8)) + PREFIX.length;
        expect(buf.subarray(record, record + 4).equals(Buffer.from([0x50, 0x4b, 0x06, 0x06]))).toBe(true);
        const cdSize = Number(buf.readBigUInt64LE(record + 40));
        const cdOffset = Number(buf.readBigUInt64LE(record + 48));
        expect(eocd - (cdOffset + cdSize + 76)).toBe(PREFIX.length);
    });

    it("keeps the slack a declared directory size spans past its records (#478)", async () => {
        const bytes = buildRawZip([...minimalEntries(), { name: "word/extra.xml", localName: "../evil.xml", content: "<x/>" }], {
            zip64: true,
            cdSlack: 16,
        });
        const zip = await loadCentralNamedArchive(bytes);
        expect(Object.keys(zip.files)).toContain("word/extra.xml");
        expect(Object.keys(zip.files)).not.toContain("evil.xml");
        expect(await zip.file("word/document.xml")!.async("string")).toBe(MINIMAL_PARTS["word/document.xml"]);
    });

    it("bails for an entry whose local-header offset lives in a ZIP64 extra field", async () => {
        expect(centralKeyedArchive(SENTINEL_LOCAL_DOTDOT_ZIP64)).toBe(SENTINEL_LOCAL_DOTDOT_ZIP64);
        // JSZip loads the archive as given, keyed by the local name.
        expect(Object.keys((await JSZip.loadAsync(SENTINEL_LOCAL_DOTDOT_ZIP64)).files)).toContain("evil.xml");
        expect(Object.keys((await loadCentralNamedArchive(SENTINEL_LOCAL_DOTDOT_ZIP64)).files)).toContain("evil.xml");
    });

    it("never loads a ZIP64 archive JSZip refuses: a locator naming a record outside the buffer", async () => {
        const buf = Buffer.from(LOCAL_DOTDOT_ZIP64);
        const locator = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x07]));
        buf.writeUInt32LE(0x7ffffff0, locator + 8);
        await expect(JSZip.loadAsync(buf)).rejects.toThrow();
        await expect(loadCentralNamedArchive(buf)).rejects.toThrow();
    });
});

// #482: JSZip decides an entry's kind after loading, from the name it reads
// (`zipEntry.js` `processAttributes`): the MS-DOS directory attribute, then a
// trailing-`/` fallback. The rewrite stores the central name in the local
// header, so a central name ending in `/` reclassifies a file entry as a
// folder — the Word-faithful reading, since Word reads the central name.
const CENTRAL_FOLDER = buildRawZip([...minimalEntries(), { name: "word/", localName: "word/extra.xml", content: "<x/>" }]);

describe("the entry's kind follows the central name (#482)", { timeout: 120_000 }, () => {
    it("reclassifies a file entry as a folder when the central name ends in '/'", async () => {
        expect(centralKeyedArchive(CENTRAL_FOLDER)).not.toBe(CENTRAL_FOLDER);
        // As given, JSZip reads the file-like local name: a file entry.
        expect((await JSZip.loadAsync(CENTRAL_FOLDER)).files["word/extra.xml"]?.dir).toBe(false);
        const zip = await loadCentralNamedArchive(CENTRAL_FOLDER);
        expect(zip.files["word/"]?.dir).toBe(true);
        expect(zip.files["word/extra.xml"]).toBeUndefined();
    });

    it("drops the part from the package view, without a refusal", async () => {
        const context = await ValidationContext.fromArchive(CENTRAL_FOLDER);
        expect(context.has("word/extra.xml")).toBe(false);
        expect(context.bytes("word/extra.xml")).toBeUndefined();
        await withTempDir(async (dir) => {
            const file = specimen(dir, "central-folder.docx", CENTRAL_FOLDER);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid, JSON.stringify(result.issues)).toBe(true);
            expect(codes(result)).not.toContain("package-extract-failed");
            expect(result.issues.filter((issue) => issue.path === "word/extra.xml")).toEqual([]);
        });
    });
});
