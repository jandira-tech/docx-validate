/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import {
    hasSignature,
    isStructuralEocd,
    readCentralDirectoryView,
    zipBytesBeforeArchive,
    zipCentralDirectoryEntries,
} from "../src/lib/zip-central-directory";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import JSZip from "jszip";

import { absoluteOffsetsAfter, buildRawZip, truncatedCentralStub, saturateLocalHeaderOffsets } from "./support/raw-zip";
import {
    extensibleDataArchive,
    multiVolumeArchive,
    outOfRangeLocatorArchive,
    wrappedLocatorArchive,
    wrappedOffsetArchive,
} from "./support/zip64-specimens";

const names = (bytes: Uint8Array) => zipCentralDirectoryEntries(bytes)?.map((entry) => entry.name);

describe("the archive JSZip loads (#447, #462)", () => {
    const duplicate = () =>
        buildRawZip([
            { name: "a.xml", content: "<1/>" },
            { name: "a.xml", content: "<2/>" },
        ]);
    // A prepended archive: its record is structural when its offsets are read
    // as absolute, the later archive's is not, and JSZip reads the later one.
    const prefix = () =>
        buildRawZip([
            { name: "pad.bin", content: Buffer.alloc(600) },
            { name: "../evil.xml", content: "x" },
        ]);
    const withTrailing = (bytes: Buffer, count: number, fake?: Buffer): Buffer =>
        Buffer.concat(fake ? [bytes, fake, Buffer.alloc(count - fake.length)] : [bytes, Buffer.alloc(count)]);
    // A signature with no readable central directory behind it.
    const fakeEocd = (): Buffer => Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(18)]);
    const loaded = async (bytes: Uint8Array) => Object.keys((await JSZip.loadAsync(bytes)).files).sort();

    it("reads the record however much data follows it", async () => {
        const trailing = withTrailing(duplicate(), 70_000);
        expect(await loaded(trailing)).toEqual(["a.xml"]);
        expect(names(trailing)).toEqual(["a.xml", "a.xml"]);
    });

    it.each([0, 70_000])("reads the later archive of a prefixed one, as JSZip does, behind %i trailing bytes", async (count) => {
        const bytes = withTrailing(Buffer.concat([prefix(), duplicate()]), count);
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "a.xml"]);
    });

    it("reads the later ZIP64 archive of a prefixed one, as JSZip does", async () => {
        const later = buildRawZip(
            [
                { name: "a.xml", content: "<1/>" },
                { name: "a.xml", content: "<2/>" },
            ],
            { zip64: true },
        );
        const bytes = Buffer.concat([prefix(), later]);
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "a.xml"]);
    });

    it("reads the empty directory JSZip reads behind a fake EOCD signature", async () => {
        const trailing = withTrailing(duplicate(), 70_000, fakeEocd());
        expect(await loaded(trailing)).toEqual([]);
        expect(names(trailing)).toEqual([]);
    });

    it("stops at the last signature, as JSZip does, however many precede it (#463)", async () => {
        const block = Buffer.concat([fakeEocd(), Buffer.alloc(4096 - 22)]);
        const bytes = Buffer.concat([duplicate(), ...Array.from({ length: 64 }, () => block)]);
        expect(await loaded(bytes)).toEqual([]);
        expect(names(bytes)).toEqual([]);
    });

    it("reads nothing where JSZip refuses: a later copy of the record, whose offsets JSZip rebases", async () => {
        const bytes = duplicate();
        const trailing = withTrailing(bytes, 70_000, bytes.subarray(bytes.length - 22));
        await expect(JSZip.loadAsync(trailing)).rejects.toThrow();
        expect(names(trailing)).toBeUndefined();
    });
});

const fileNames = (bytes: Uint8Array) =>
    zipCentralDirectoryEntries(bytes)
        ?.filter((entry) => !entry.directory)
        .map((entry) => entry.name);

describe("JSZip's 32-bit field arithmetic (#470)", () => {
    const entries = [
        { name: "a.xml", content: "<1/>" },
        { name: "word%2Fxtra.xml", content: "<x/>" },
    ];
    const loadedRecords = async (bytes: Uint8Array) => Object.keys((await JSZip.loadAsync(bytes)).files);
    // The EOCD declares a 10-byte comment the buffer does not hold.
    const commentOverrun = (zip: Buffer): Buffer => {
        const copy = Buffer.from(zip);
        copy.writeUInt16LE(10, copy.length - 2);
        return copy;
    };

    it("reads the ZIP64 record at the locator offset JSZip wraps to 32 bits", async () => {
        const bytes = wrappedLocatorArchive(entries);
        expect(await loadedRecords(bytes)).toEqual(["a.xml", "word%2Fxtra.xml", "z-clean.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "word%2Fxtra.xml", "z-clean.xml"]);
    });

    it("keeps the duplicates of the record JSZip selects", () => {
        const bytes = wrappedLocatorArchive([...entries, { name: "a.xml", content: "<2/>" }]);
        expect(names(bytes)).toEqual(["a.xml", "word%2Fxtra.xml", "a.xml", "z-clean.xml"]);
    });

    it("wraps the ZIP64 record's directory offset as JSZip does", async () => {
        const bytes = wrappedOffsetArchive(entries);
        expect(await loadedRecords(bytes)).toEqual(["a.xml", "word%2Fxtra.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "word%2Fxtra.xml"]);
    });

    it.each([
        { name: "a locator offset past the buffer", bytes: outOfRangeLocatorArchive(entries) },
        { name: "ZIP64 extensible data", bytes: extensibleDataArchive(entries) },
        { name: "a locator counting two disks", bytes: multiVolumeArchive(entries) },
        { name: "an EOCD comment running past the buffer", bytes: commentOverrun(buildRawZip(entries)) },
    ])("reads nothing where JSZip throws: $name", async ({ bytes }) => {
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
    });
});

describe("zipCentralDirectoryEntries", () => {
    it("lists a JSZip archive's entries, in central-directory order", async () => {
        const bytes = await buildDocx();
        // JSZip stores a folder entry for each implicit directory.
        expect(fileNames(bytes)).toEqual(Object.keys(MINIMAL_PARTS).sort());
    });

    it("keeps duplicate names, in central-directory order", () => {
        const bytes = buildRawZip([
            { name: "b.xml", content: "<b/>" },
            { name: "a.xml", content: "<1/>" },
            { name: "a.xml", content: "<2/>" },
        ]);
        expect(names(bytes)).toEqual(["b.xml", "a.xml", "a.xml"]);
    });

    it("keeps a name resolution would rewrite", () => {
        const bytes = buildRawZip([
            { name: "a/b/../x.xml", content: "<x/>" },
            { name: "a/x.xml", content: "<x/>" },
        ]);
        expect(names(bytes)).toEqual(["a/b/../x.xml", "a/x.xml"]);
    });

    it("decodes a name as UTF-8 when general-purpose bit 11 is set", () => {
        const bytes = buildRawZip([{ name: "café.xml", content: "<x/>", utf8: true }]);
        expect(names(bytes)).toEqual(["café.xml"]);
    });

    it("decodes a name as CP437 without the UTF-8 flag", () => {
        const bytes = buildRawZip([
            { name: "café.xml", nameBytes: Uint8Array.from([0x63, 0x61, 0x66, 0x82, 0x2e, 0x78, 0x6d, 0x6c]), content: "<x/>" },
        ]);
        expect(names(bytes)).toEqual(["café.xml"]);
    });

    it("reads the central directory through the ZIP64 locator and record", () => {
        const bytes = buildRawZip(
            [
                { name: "a.xml", content: "<a/>" },
                { name: "a.xml", content: "<b/>" },
            ],
            { zip64: true },
        );
        expect(names(bytes)).toEqual(["a.xml", "a.xml"]);
    });

    it("marks folder entries, by trailing slash or by the DOS directory bit", () => {
        const bytes = buildRawZip([
            { name: "word/", content: "" },
            { name: "docProps", directory: true },
            { name: "word/a.xml", content: "<a/>" },
        ]);
        expect(zipCentralDirectoryEntries(bytes)?.map((entry) => `${entry.directory}:${entry.name}`)).toEqual([
            "true:word/",
            "true:docProps",
            "false:word/a.xml",
        ]);
    });

    it.each([
        ["no EOCD at all", Buffer.from("not a zip at all")],
        ["a truncated archive", buildRawZip([{ name: "a.xml", content: "<a/>" }]).subarray(0, 12)],
        [
            "a central directory declared past its record",
            (() => {
                const bytes = buildRawZip([
                    { name: "a.xml", content: "<a/>" },
                    { name: "b.xml", content: "<b/>" },
                ]);
                // A directory whose declared span ends 30 bytes after its
                // record sits: the rebase is negative and JSZip throws before
                // it ever walks (`extraBytes < 0` in `readEndOfCentral`).
                bytes.writeUInt32LE(bytes.readUInt32LE(bytes.length - 10) + 30, bytes.length - 10);
                return bytes;
            })(),
        ],
        [
            "a central directory that runs into the EOCD",
            (() => {
                const bytes = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
                // Claim exactly the directory plus four bytes of the record itself.
                bytes.writeUInt32LE(bytes.readUInt32LE(bytes.length - 10) + 4, bytes.length - 10);
                return bytes;
            })(),
        ],
        [
            "a ZIP64 offset without a locator",
            (() => {
                const bytes = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
                // A saturated offset with no locator in place is malformed.
                bytes.writeUInt32LE(0xffffffff, bytes.length - 6);
                return bytes;
            })(),
        ],
    ])("returns undefined on %s", (_label, bytes) => {
        expect(zipCentralDirectoryEntries(bytes)).toBeUndefined();
    });
});

describe("zipBytesBeforeArchive: bytes in front of the archive JSZip loads (#464)", () => {
    const JUNK = Buffer.alloc(4096, 0x20);
    const FAKE_SIG_STUB = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.alloc(4092)]);
    const DOCX_NAMES = Object.keys(MINIMAL_PARTS).sort();
    const zipPrefix = () =>
        buildRawZip([
            { name: "pad.bin", content: Buffer.alloc(600) },
            { name: "ok.xml", content: "<x/>" },
        ]);
    const loaded = async (bytes: Uint8Array) => {
        const files = (await JSZip.loadAsync(bytes)).files;
        return Object.keys(files)
            .filter((name) => !files[name]!.dir)
            .sort();
    };

    it("counts none for an archive that begins at the file's first byte", async () => {
        const bytes = await buildDocx();
        expect(await loaded(bytes)).toEqual(DOCX_NAMES);
        expect(zipBytesBeforeArchive(bytes)).toBe(0);
    });

    // [label, bytes, expected prefix, the names JSZip loads]
    const CASES: readonly (readonly [string, () => Promise<Buffer>, () => number, readonly string[]])[] = [
        [
            "a whole prepended ZIP archive",
            async () => Buffer.concat([zipPrefix(), await buildDocx()]),
            () => zipPrefix().length,
            DOCX_NAMES,
        ],
        ["junk, the archive's offsets left relative", async () => Buffer.concat([JUNK, await buildDocx()]), () => JUNK.length, DOCX_NAMES],
        [
            "junk, the absolute offsets a self-extractor writes",
            async () => Buffer.concat([JUNK, absoluteOffsetsAfter(await buildDocx(), JUNK.length)]),
            () => JUNK.length,
            DOCX_NAMES,
        ],
        ["a single byte", async () => Buffer.concat([Buffer.from([0x0a]), await buildDocx()]), () => 1, DOCX_NAMES],
        [
            "junk whose first bytes fake a local-header signature",
            async () => Buffer.concat([FAKE_SIG_STUB, await buildDocx()]),
            () => FAKE_SIG_STUB.length,
            DOCX_NAMES,
        ],
        ["an empty archive", async () => Buffer.concat([JUNK, buildRawZip([])]), () => JUNK.length, []],
    ];

    it.each(CASES)("counts the prefix in front of the archive behind %s", async (_label, make, prefix, names) => {
        const bytes = await make();
        expect(await loaded(bytes), "JSZip loads the archive behind the prefix").toEqual(names);
        expect(zipBytesBeforeArchive(bytes)).toBe(prefix());
    });

    it("counts the prefix behind a ZIP64 archive", async () => {
        const live = buildRawZip(
            Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
            { zip64: true },
        );
        const bytes = Buffer.concat([JUNK, live]);
        expect(await loaded(bytes)).toEqual(DOCX_NAMES);
        expect(zipBytesBeforeArchive(live)).toBe(0);
        expect(zipBytesBeforeArchive(bytes)).toBe(JUNK.length);
    });

    it("counts a whole document in front of a smaller archive appended after it", async () => {
        const docx = await buildDocx();
        const bytes = Buffer.concat([docx, zipPrefix()]);
        expect(await loaded(bytes)).toEqual(["ok.xml", "pad.bin"]);
        expect(zipBytesBeforeArchive(bytes)).toBe(docx.length);
    });

    it("takes the lowest stored local-header offset among the entries, not the first directory record's", async () => {
        // b.xml is stored first; swapping the two equal-size central records
        // makes the first record name a.xml, whose local header sits later.
        const bytes = buildRawZip([
            { name: "b.xml", content: "<b/>" },
            { name: "a.xml", content: "<a/>" },
        ]);
        const eocd = bytes.length - 22;
        expect(bytes.readUInt16LE(eocd + 10)).toBe(2);
        const cd = bytes.readUInt32LE(eocd + 16);
        const span = 46 + "b.xml".length;
        const first = Buffer.from(bytes.subarray(cd, cd + span));
        bytes.set(bytes.subarray(cd + span, cd + 2 * span), cd);
        bytes.set(first, cd + span);
        const prefixed = Buffer.concat([JUNK, absoluteOffsetsAfter(bytes, JUNK.length)]);
        expect(await loaded(prefixed)).toEqual(["a.xml", "b.xml"]);
        expect(zipBytesBeforeArchive(prefixed)).toBe(JUNK.length);
    });

    it.each([
        ["no EOCD at all", Buffer.from("not a zip at all")],
        ["a truncated archive", buildRawZip([{ name: "a.xml", content: "<a/>" }]).subarray(0, 12)],
        [
            "a later copy of the record, whose offsets JSZip rebases and refuses",
            (() => {
                const archive = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
                return Buffer.concat([archive, Buffer.alloc(70), archive.subarray(archive.length - 22)]);
            })(),
        ],
    ])("returns undefined on %s", (_label, bytes) => {
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });
});

// A central-directory signature the buffer cuts off inside the record's

// A ZIP64 record whose declared total count reads -1 through `readInt(8)`
// (all FF bytes) is no error for JSZip when records exist: `readCentralDir`
// throws only for a nonzero count that finds no record (#492).
describe("the ZIP64 declared count (#492)", () => {
    const loaded = async (bytes: Uint8Array) => Object.keys((await JSZip.loadAsync(bytes)).files).sort();
    const single = () => buildRawZip([{ name: "a.xml", content: "<a/>" }], { zip64: true });
    const zip64Count = (bytes: Buffer, count: bigint): Buffer => {
        const record = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
        const stored = BigInt.asUintN(64, count);
        bytes.writeBigUInt64LE(stored, record + 24);
        bytes.writeBigUInt64LE(stored, record + 32);
        return bytes;
    };

    it.each([-1n, 0n, 2n])("tolerates a declared count of %s, as JSZip does", async (count) => {
        const bytes = zip64Count(single(), count);
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(names(bytes)).toEqual(["a.xml"]);
        expect(zipBytesBeforeArchive(bytes)).toBe(0);
    });

    it("returns undefined where JSZip throws: count -1 and no record found", async () => {
        const bytes = zip64Count(single(), -1n);
        const record = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
        const directory = Number(bytes.readBigUInt64LE(record + 48));
        bytes.writeUInt32LE(0, directory);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });
});

// A saturated local-header offset whose record's ZIP64 extra field (0x0001)
// carries the real value is how JSZip reads the entry (zipEntry.js
// `parseZIP64ExtraField`): `readInt(4)` sees -1, the extra's 8-byte field —
// read with the same 32-bit arithmetic — replaces it (#490). The measure is
// what JSZip loads, and `undefined` exactly where JSZip throws.
describe("the ZIP64 extra field's local-header offset (#490)", () => {
    const loaded = async (bytes: Uint8Array) => {
        const files = (await JSZip.loadAsync(bytes)).files;
        return Object.keys(files)
            .filter((name) => !files[name]!.dir)
            .sort();
    };
    const DOCX_NAMES = Object.keys(MINIMAL_PARTS).sort();
    const JUNK = Buffer.alloc(4096, 0x20);

    it.each([1, 2, undefined] as const)(
        "measures the extra field's value, not the saturated base field, when %s record(s) use the spelling",
        async (count) => {
            const bytes = saturateLocalHeaderOffsets(await buildDocx(), { count });
            expect(await loaded(bytes)).toEqual(DOCX_NAMES);
            expect(zipBytesBeforeArchive(bytes)).toBe(0);
        },
    );

    it("reads the extra field in JSZip's order, after saturated size fields", async () => {
        const bytes = saturateLocalHeaderOffsets(buildRawZip([{ name: "a.xml", content: "<a/>" }]), {
            sizes: true,
        });
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(zipBytesBeforeArchive(bytes)).toBe(0);
    });

    it("measures a real prefix in front of the archive exactly", async () => {
        const archive = saturateLocalHeaderOffsets(await buildDocx(), { count: 1 });
        const bytes = Buffer.concat([JUNK, archive]);
        expect(await loaded(bytes)).toEqual(DOCX_NAMES);
        expect(zipBytesBeforeArchive(bytes)).toBe(JUNK.length);
    });

    it.each([
        [
            "a saturated offset with no extra field",
            async () => {
                const docx = await buildDocx();
                const eocd = docx.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
                docx.writeUInt32LE(0xffffffff, docx.readUInt32LE(eocd + 16) + 42);
                return docx;
            },
        ],
        [
            "an extra field too short for the offset read",
            () => saturateLocalHeaderOffsets(buildRawZip([{ name: "a.xml", content: "<a/>" }]), { fieldLength: 4 }),
        ],
        [
            "a negative effective offset",
            () => saturateLocalHeaderOffsets(buildRawZip([{ name: "a.xml", content: "<a/>" }]), { value: -1n }),
        ],
        [
            "a base field in [0x80000000, 0xFFFFFFFE]",
            () => {
                const bytes = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
                const eocd = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
                bytes.writeUInt32LE(0x80000000, bytes.readUInt32LE(eocd + 16) + 42);
                return bytes;
            },
        ],
    ])("returns undefined where JSZip throws: %s", async (_label, make) => {
        const bytes = await make();
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });
});

// JSZip's `readLocalPart` (zipEntry.js) keeps reading past the local
// signature: the fixed 30-byte header, the local name and extra field, the
// central-declared compressed stream, and a compression method it knows. A
// measure for an archive whose local part throws there claims a prefix for a
// package that never opens (#524).
describe("the local part JSZip reads past the signature (#524)", () => {
    const single = () => buildRawZip([{ name: "a.xml", content: "<a/>" }]);
    const cdOf = (bytes: Buffer): number => bytes.readUInt32LE(bytes.length - 22 + 16);

    it.each([
        [
            "a compressed size past the buffer",
            (() => {
                const bytes = single();
                bytes.writeUInt32LE(0x7fff0000, cdOf(bytes) + 20);
                return bytes;
            })(),
        ],
        [
            "a local signature at the very end of the buffer",
            (() => {
                // The record's local header points at a `50 4B 03 04`
                // appended as the EOCD's four-byte comment: the signature
                // reads, the 30-byte header after it does not.
                const bytes = single();
                bytes.writeUInt32LE(bytes.length, cdOf(bytes) + 42);
                bytes.writeUInt16LE(4, bytes.length - 2);
                return Buffer.concat([bytes, Buffer.from([0x50, 0x4b, 0x03, 0x04])]);
            })(),
        ],
        [
            "an unknown compression method",
            (() => {
                const bytes = single();
                bytes.writeUInt16LE(99, cdOf(bytes) + 10);
                return bytes;
            })(),
        ],
    ])("returns undefined where JSZip throws: %s", async (_label, bytes) => {
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });

    it("still measures the archive whose local parts fit", async () => {
        const bytes = single();
        expect(Object.keys((await JSZip.loadAsync(bytes)).files)).toEqual(["a.xml"]);
        expect(zipBytesBeforeArchive(bytes)).toBe(0);
    });
});

// JSZip's `DataReader.checkIndex` throws on both sides: `zero + index` past
// the buffer and an index below zero. A central compressed size that reads
// negative — the 32-bit field 0x80000000, or a ZIP64 0x0001 extra resolving
// below zero; neither is the -1 the sentinel check catches — sends
// `readLocalPart`'s final `readData(compressedSize)` below zero, so `load`
// rejects: no prefix and no names view may be claimed for it (#555). A small
// negative size whose end stays at or above zero still reads, as in JSZip.
describe("a compressed size that reads negative (#555)", () => {
    const single = () => buildRawZip([{ name: "a.xml", content: "<a/>" }]);
    const cdOf = (bytes: Buffer): number => bytes.readUInt32LE(bytes.length - 22 + 16);
    const withBaseCompressed = (value: number): Buffer => {
        const bytes = single();
        bytes.writeUInt32LE(value, cdOf(bytes) + 20);
        return bytes;
    };
    // Saturate the record's size fields and resolve them from a 0x0001 extra
    // field, in JSZip's order: [uncompressed = the real size][compressed = value].
    const withZip64Compressed = (value: bigint): Buffer => {
        const bytes = single();
        const cd = cdOf(bytes);
        const field = Buffer.alloc(20);
        field.writeUInt16LE(0x0001, 0);
        field.writeUInt16LE(16, 2);
        field.writeBigUInt64LE(BigInt(bytes.readUInt32LE(cd + 24)), 4);
        field.writeBigUInt64LE(BigInt.asUintN(64, value), 12);
        const insert = cd + 46 + bytes.readUInt16LE(cd + 28);
        const grown = Buffer.concat([bytes.subarray(0, insert), field, bytes.subarray(insert)]);
        grown.writeUInt32LE(0xffffffff, cd + 20);
        grown.writeUInt32LE(0xffffffff, cd + 24);
        grown.writeUInt16LE(grown.readUInt16LE(cd + 30) + field.length, cd + 30);
        const eocd = grown.length - 22;
        grown.writeUInt32LE(grown.readUInt32LE(eocd + 12) + field.length, eocd + 12);
        return grown;
    };
    const prefixed = (bytes: Buffer): Buffer => Buffer.concat([Buffer.alloc(64, 0x41), bytes]);

    it.each([
        ["the 32-bit field 0x80000000", () => withBaseCompressed(0x80000000)],
        ["a ZIP64 extra resolving the size below zero", () => withZip64Compressed(-0x80000000n)],
    ])("returns undefined where JSZip throws: %s", async (_label, make) => {
        const bytes = prefixed(make());
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
        expect(zipCentralDirectoryEntries(bytes)).toBeUndefined();
    });

    it("still measures a resolved compressed size whose end stays above zero (-8)", async () => {
        const bytes = prefixed(withZip64Compressed(-8n));
        expect(Object.keys((await JSZip.loadAsync(bytes)).files)).toEqual(["a.xml"]);
        expect(zipBytesBeforeArchive(bytes)).toBe(64);
        expect(names(bytes)).toEqual(["a.xml"]);
    });

    // `readLocalPart` bounds the local name and the extra field on their own
    // reads — `readData(fileNameLength)`, then `skip(localExtraFieldsLength)`
    // — before the compressed read. A small negative resolved size (a ZIP64
    // 0x0001 extra resolving to -8, or the 32-bit field 0xFFFFFFF8, which
    // `readInt(4)` reads as -8) shortens only the combined local-part end, so
    // an overrun of up to |size| bytes past the buffer still throws there
    // (#566).
    describe("a local name/extra overrun a small negative size masks (#566)", () => {
        // The local header sits at archive offset 0 — buffer offset 64 with
        // the prefix — and the declared local extra length ends the local
        // name and extra field `overrun` bytes past the buffer's end.
        const withLocalExtraOverrun = (bytes: Buffer, overrun: number): Buffer => {
            const nameLength = bytes.readUInt16LE(64 + 26);
            bytes.writeUInt16LE(bytes.length + overrun - (64 + 30 + nameLength), 64 + 28);
            return bytes;
        };

        it.each([
            ["a ZIP64 extra resolving the size to -8", () => withZip64Compressed(-8n)],
            ["the 32-bit compressed-size field 0xFFFFFFF8", () => withBaseCompressed(0xfffffff8)],
        ])("returns undefined where JSZip's name/extra reads throw: %s", async (_label, make) => {
            for (const overrun of [1, 8]) {
                const bytes = withLocalExtraOverrun(prefixed(make()), overrun);
                await expect(JSZip.loadAsync(bytes), `overrun=${overrun}`).rejects.toThrow();
                expect(zipBytesBeforeArchive(bytes), `overrun=${overrun}`).toBeUndefined();
                expect(zipCentralDirectoryEntries(bytes), `overrun=${overrun}`).toBeUndefined();
            }
        });
    });
});

// 46-byte fixed header: JSZip's `readCentralPart` reads every field through a
// bounds-checked read, so it throws; the walk must refuse the same bytes
// instead of reading `undefined` fields into a phantom record (#498).
describe("the truncated fixed record header JSZip refuses (#498)", () => {
    it.each(Array.from({ length: 23 }, (_unused, i) => 26 + i))(
        "returns undefined with %i bytes from the signature to the buffer's end, where JSZip throws",
        async (length) => {
            const bytes = truncatedCentralStub(length);
            await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
            expect(readCentralDirectoryView(bytes)).toBeUndefined();
            expect(zipCentralDirectoryEntries(bytes)).toBeUndefined();
            expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
        },
    );

    it("never produces a non-finite prefix count", () => {
        for (const bytes of [truncatedCentralStub(26), buildRawZip([{ name: "a.xml", content: "<a/>" }])]) {
            const prefix = zipBytesBeforeArchive(bytes);
            expect(prefix === undefined || Number.isFinite(prefix)).toBe(true);
        }
    });
});

// JSZip's `readCentralDir` (zipEntries.js): the walk starts at the rebased
// directory offset and reads records while the signature matches. The
// declared size only feeds `reader.zero`, and a count mismatch is tolerated
// unless no record at all is found. The names mirror that exactly: they are
// what JSZip loads when it loads, and `undefined` exactly where JSZip throws
// (#478).
describe("the central-directory walk JSZip runs (#478)", () => {
    const pair = () =>
        buildRawZip([
            { name: "a.xml", content: "<1/>" },
            { name: "a.xml", content: "<2/>" },
        ]);
    const loaded = async (bytes: Uint8Array) => Object.keys((await JSZip.loadAsync(bytes)).files).sort();
    // The EOCD record's fields: the record count JSZip consults is 10 bytes
    // before its end, the directory size 6 bytes before that.
    const withDeclaredCount = (bytes: Buffer, count: number): Buffer => {
        bytes.writeUInt16LE(count, bytes.length - 12);
        return bytes;
    };
    const withDeclaredSize = (bytes: Buffer, size: number): Buffer => {
        bytes.writeUInt32LE(size, bytes.length - 10);
        return bytes;
    };

    it("walks the records and stops at the original record when an appended record overstates the size", async () => {
        const base = pair();
        const eocd = base.length - 22;
        const size = base.readUInt32LE(eocd + 12);
        // A second copy of the record whose directory size covers the first:
        // the walk it describes ends at the original record, not at itself.
        const appended = Buffer.from(base.subarray(eocd));
        appended.writeUInt32LE(size + 22, 12);
        const bytes = Buffer.concat([base, appended]);
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "a.xml"]);
        // The walked directory ends at the original record's offset.
        expect(readCentralDirectoryView(bytes)?.cdEnd).toBe(eocd);
    });

    it.each([5, 1, 0])("tolerates a declared count of %i, as JSZip does", async (count) => {
        const bytes = withDeclaredCount(pair(), count);
        expect(await loaded(bytes)).toEqual(["a.xml"]);
        expect(names(bytes)).toEqual(["a.xml", "a.xml"]);
    });

    it.each([
        [
            "a corrupted first central-header signature",
            (() => {
                const bytes = pair();
                bytes.writeUInt32LE(0x04034b50, bytes.readUInt32LE(bytes.length - 6));
                return bytes;
            })(),
        ],
        [
            "a rebased directory offset that lands mid-record",
            (() => {
                // The declared span shrinks by 8, so the walk starts eight
                // bytes into the first record, where no signature sits.
                const bytes = pair();
                return withDeclaredSize(bytes, bytes.readUInt32LE(bytes.length - 10) - 8);
            })(),
        ],
    ])("reads nothing where JSZip throws on %s with no record found", async (_label, bytes) => {
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
    });

    it.each([
        ["name", 28],
        ["extra field", 30],
        ["comment", 32],
    ])("reads nothing where JSZip throws: a record's %s overruns the buffer", async (_label, field) => {
        const bytes = pair();
        // The second record sits 51 bytes into the directory (46 + "a.xml").
        bytes.writeUInt16LE(0x4000, bytes.readUInt32LE(bytes.length - 6) + 51 + field);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
    });

    it("reads nothing where JSZip throws: fewer than four bytes left for the next signature read", async () => {
        const bytes = pair();
        const eocd = bytes.length - 22;
        const cdOffset = bytes.readUInt32LE(eocd + 16);
        const cdSize = bytes.readUInt32LE(eocd + 12);
        const fake = eocd - 46;
        // A record header ending at the record: its name field is the
        // record's own first 20 bytes, so the walk stops two bytes short of
        // the buffer's end, where a signature cannot be read.
        bytes.writeUInt32LE(0x02014b50, fake);
        bytes.writeUInt16LE(0, fake + 8);
        bytes.writeUInt16LE(20, fake + 28);
        bytes.writeUInt16LE(0, fake + 30);
        bytes.writeUInt16LE(0, fake + 32);
        bytes.writeUInt32LE(cdOffset, eocd + 16);
        bytes.writeUInt32LE(cdSize + 46, eocd + 12);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
    });

    it("reads nothing where JSZip throws: an entry flagged encrypted", async () => {
        const bytes = pair();
        // `readCentralPart` refuses the encrypted bit, in the same walk.
        bytes.writeUInt16LE(1, bytes.readUInt32LE(bytes.length - 6) + 8);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
    });

    it("reads nothing where JSZip throws: an extra field inside a fitting region declaring more data than the buffer holds (#525)", async () => {
        const base = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
        const eocd = base.length - 22;
        const cd = base.readUInt32LE(eocd + 16);
        // An 8-byte extra region — the declared size grows with it — whose
        // first field (0x0001) declares 0xFFF0 bytes: the region fits the
        // buffer, the field's `readData` does not.
        const region = Buffer.from([0x01, 0x00, 0xf0, 0xff, 0x00, 0x00, 0x00, 0x00]);
        const bytes = Buffer.concat([base.subarray(0, eocd), region, base.subarray(eocd)]);
        bytes.writeUInt16LE(region.length, cd + 30);
        const shifted = eocd + region.length;
        bytes.writeUInt32LE(bytes.readUInt32LE(shifted + 12) + region.length, shifted + 12);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow();
        expect(names(bytes)).toBeUndefined();
        // The measure's own extra-field scan already bounds the declared
        // length; only the names view used to disagree.
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });
});

// `parseZIP64ExtraField` reads each saturated size with `readInt(8)`, and
// throws when the 0x0001 field is shorter than that read. The uncompressed
// slot is consumed first, so a short field fails there before the compressed
// slot is reached (#490).
describe("a ZIP64 size the extra field is too short to resolve", () => {
    const withShortZip64Size = (field: 20 | 24): Buffer => {
        const base = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
        const eocd = base.length - 22;
        const cd = base.readUInt32LE(eocd + 16);
        // Four data bytes: the field fits the buffer, the 8-byte read does not.
        const extra = Buffer.from([0x01, 0x00, 0x04, 0x00, 1, 2, 3, 4]);
        const bytes = Buffer.concat([base.subarray(0, eocd), extra, base.subarray(eocd)]);
        bytes.writeUInt32LE(0xffffffff, cd + field);
        bytes.writeUInt16LE(extra.length, cd + 30);
        const shifted = eocd + extra.length;
        bytes.writeUInt32LE(bytes.readUInt32LE(shifted + 12) + extra.length, shifted + 12);
        return bytes;
    };

    it.each([
        ["uncompressed", 24],
        ["compressed", 20],
    ] as const)("returns undefined where JSZip throws: the %s size's ZIP64 value is shorter than 8 bytes", async (_label, field) => {
        const bytes = withShortZip64Size(field);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow(/End of data reached/);
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
        expect(names(bytes)).toBeUndefined();
    });
});

describe("a resolved local offset that is not a local header", () => {
    it("returns undefined where JSZip throws: the offset names the central directory", async () => {
        const bytes = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
        const cd = bytes.readUInt32LE(bytes.length - 22 + 16);
        bytes.writeUInt32LE(cd, cd + 42);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow(/unexpected signature/);
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
        expect(names(bytes)).toBeUndefined();
    });
});

describe("isStructuralEocd", () => {
    it("accepts an empty archive and a directory it can see, and a ZIP64 offset only beside a locator", () => {
        const empty = Buffer.alloc(22);
        empty.writeUInt32LE(0x06054b50, 0);
        expect(isStructuralEocd(empty, 0)).toBe(true);
        expect(hasSignature(empty, -1, 0x06054b50)).toBe(false);

        const bytes = buildRawZip([{ name: "a.xml", content: "A" }]);
        const eocd = bytes.length - 22;
        expect(isStructuralEocd(bytes, eocd)).toBe(true);

        const saturated = Buffer.from(bytes);
        saturated.writeUInt32LE(0xffffffff, eocd + 16);
        expect(isStructuralEocd(saturated, eocd)).toBe(false);

        const located = Buffer.alloc(42);
        located.writeUInt32LE(0x07064b50, 0);
        located.writeUInt32LE(0x06054b50, 20);
        located.writeUInt32LE(0xffffffff, 36);
        expect(isStructuralEocd(located, 20)).toBe(true);
    });

    it("rejects a record that names no directory", () => {
        const record = Buffer.alloc(22);
        record.writeUInt32LE(0x06054b50, 0);
        record.writeUInt16LE(1, 10);
        record.writeUInt32LE(1, 12);
        expect(isStructuralEocd(record, 0)).toBe(false);
    });
});

describe("a ZIP64 locator that names no end record", () => {
    it("returns undefined where JSZip throws: nothing at the locator's offset is a ZIP64 end signature", async () => {
        const base = buildRawZip([{ name: "a.xml", content: "A" }]);
        const eocd = base.length - 22;
        const locator = Buffer.alloc(20);
        locator.writeUInt32LE(0x07064b50, 0);
        locator.writeUInt32LE(0, 8);
        locator.writeUInt32LE(1, 16);
        const bytes = Buffer.concat([base.subarray(0, eocd), locator, base.subarray(eocd)]);
        bytes.writeUInt16LE(0xffff, eocd + locator.length + 10);
        await expect(JSZip.loadAsync(bytes)).rejects.toThrow(/ZIP64 end of central directory/);
        expect(readCentralDirectoryView(bytes)).toBeUndefined();
        expect(zipBytesBeforeArchive(bytes)).toBeUndefined();
    });
});
