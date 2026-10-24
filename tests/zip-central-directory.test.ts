/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { zipCentralDirectoryEntries } from "../src/lib/zip-central-directory";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import JSZip from "jszip";

import { buildRawZip } from "./support/raw-zip";
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
        expect(fileNames(bytes)).toEqual([...Object.keys(MINIMAL_PARTS)].sort());
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
            "a central-directory entry count that lies",
            (() => {
                const bytes = buildRawZip([{ name: "a.xml", content: "<a/>" }]);
                // The EOCD's record counts are 12 bytes and 10 bytes before its end.
                bytes.writeUInt16LE(7, bytes.length - 14);
                bytes.writeUInt16LE(7, bytes.length - 12);
                return bytes;
            })(),
        ],
        [
            "a central directory cut short",
            (() => {
                const bytes = buildRawZip([
                    { name: "a.xml", content: "<a/>" },
                    { name: "b.xml", content: "<b/>" },
                ]);
                // Claim a larger directory than the records hold.
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
