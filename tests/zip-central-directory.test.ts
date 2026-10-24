/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { zipCentralDirectoryEntries, lastStructuralEocdOffset } from "../src/lib/zip-central-directory";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import JSZip from "jszip";

import { buildRawZip } from "./support/raw-zip";

const names = (bytes: Uint8Array) => zipCentralDirectoryEntries(bytes)?.map((entry) => entry.name);

describe("the EOCD record behind more than 64 KiB of trailing data (#447)", () => {
    const duplicate = () =>
        buildRawZip([
            { name: "a.xml", content: "<1/>" },
            { name: "a.xml", content: "<2/>" },
        ]);
    const withTrailing = (bytes: Buffer, count: number, fake?: Buffer): Buffer =>
        Buffer.concat(fake ? [bytes, fake, Buffer.alloc(count - fake.length)] : [bytes, Buffer.alloc(count)]);
    // A signature with no readable central directory behind it.
    const fakeEocd = (): Buffer => Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(18)]);

    it("finds the last structural record however much data follows it", () => {
        const bytes = duplicate();
        const trailing = withTrailing(bytes, 70_000);
        expect(lastStructuralEocdOffset(trailing)).toBe(bytes.length - 22);
        expect(names(trailing)).toEqual(["a.xml", "a.xml"]);
    });

    it("agrees with JSZip, which loads the same archive from that record", async () => {
        const trailing = withTrailing(duplicate(), 70_000);
        const zip = await JSZip.loadAsync(trailing);
        expect(Object.keys(zip.files).sort()).toEqual(["a.xml"]);
        expect(fileNames(trailing)).toEqual(["a.xml", "a.xml"]);
    });

    it("skips a fake EOCD signature in the trailing data", () => {
        const bytes = duplicate();
        const trailing = withTrailing(bytes, 70_000, fakeEocd());
        expect(lastStructuralEocdOffset(trailing)).toBe(bytes.length - 22);
        expect(names(trailing)).toEqual(["a.xml", "a.xml"]);
    });

    it("reads the directory through a later structural record in the trailing data", () => {
        const bytes = duplicate();
        // A later record whose fields still point at the real central directory.
        const copy = bytes.subarray(bytes.length - 22);
        const trailing = withTrailing(bytes, 70_000, copy);
        expect(lastStructuralEocdOffset(trailing)).toBe(bytes.length);
        expect(names(trailing)).toEqual(["a.xml", "a.xml"]);
    });
});

const fileNames = (bytes: Uint8Array) =>
    zipCentralDirectoryEntries(bytes)
        ?.filter((entry) => !entry.directory)
        .map((entry) => entry.name);

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
