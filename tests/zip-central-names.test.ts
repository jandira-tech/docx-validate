/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { type ZipCentralRecord, readCentralDirectoryView, zipCentralDirectoryEntries } from "../src/lib/zip-central-directory";
import { archiveWithoutEntries, centralKeyedArchive, loadCentralNamedArchive } from "../src/lib/zip-central-names";
import { buildRawZip, saturateLocalHeaderOffsets } from "./support/raw-zip";

const HOSTILE = (record: ZipCentralRecord): boolean => record.name.includes("..");

const zipContents = async (zip: JSZip): Promise<Record<string, string>> => {
    const out: Record<string, string> = {};
    for (const [name, entry] of Object.entries(zip.files)) {
        if (!entry.dir) out[name] = await entry.async("string");
    }
    return out;
};

const contents = async (bytes: Uint8Array): Promise<Record<string, string>> => zipContents(await JSZip.loadAsync(bytes));

/** The entries `loadCentralNamedArchive` reads, the loader the validator uses. */
const loadedContents = async (bytes: Uint8Array): Promise<Record<string, string>> => zipContents(await loadCentralNamedArchive(bytes));

/** Three entries, the middle one hostile: dropping it keeps two 5-byte names. */
const zip64WithHostile = (): Buffer =>
    buildRawZip(
        [
            { name: "a.xml", content: "A" },
            { name: "x/../b.xml", content: "HOSTILE" },
            { name: "b.xml", content: "B" },
        ],
        { zip64: true },
    );

/**
 * Junk between the central directory and the EOCD, counted by the EOCD's
 * declared size, as a size spanning past its records does (#478).
 */
const withJunkBeforeEocd = (bytes: Buffer, junk = 4): Buffer => {
    const eocd = bytes.length - 22;
    const out = Buffer.concat([bytes.subarray(0, eocd), Buffer.alloc(junk, 0x5a), bytes.subarray(eocd)]);
    out.writeUInt32LE(out.readUInt32LE(eocd + 12) + junk, eocd + 12);
    return out;
};

/**
 * Junk between the ZIP64 end record and its locator, a tail no rewrite can
 * move (#479). The record's declared directory size counts the junk, so both
 * readers keep finding the EOCD where the record says it is.
 */
const withJunkBetweenZip64RecordAndLocator = (bytes: Buffer, junk = 5): Buffer => {
    const eocd = bytes.length - 22;
    const locator = eocd - 20;
    const record = locator - 56;
    const out = Buffer.concat([bytes.subarray(0, locator), Buffer.alloc(junk, 0x5a), bytes.subarray(locator)]);
    out.writeBigUInt64LE(out.readBigUInt64LE(record + 40) + BigInt(junk), record + 40);
    return out;
};

/** Store `comment` as the EOCD comment of an archive `buildRawZip` wrote. */
const withEocdComment = (bytes: Buffer, comment: string): Buffer => {
    const out = Buffer.concat([bytes, Buffer.from(comment)]);
    out.writeUInt16LE(comment.length, bytes.length - 22 + 20);
    return out;
};

/**
 * Copy this archive's ZIP64 end record into the EOCD comment. JSZip searches
 * for a ZIP64 signature when the locator misses, and that copy is the one the
 * search finds (APPNOTE 4.3.15, #521).
 */
const withZip64EndRecordComment = (bytes: Buffer): Buffer => {
    const eocd = bytes.length - 22;
    const record = eocd - 76;
    const comment = Buffer.from(bytes.subarray(record, record + 56));
    const out = Buffer.concat([bytes, comment]);
    out.writeUInt16LE(comment.length, eocd + 20);
    return out;
};

/** Central records in the opposite order. Local records and the EOCD stay put (APPNOTE 4.4.1.3). */
const reverseCentralRecordOrder = (bytes: Buffer): Buffer => {
    const eocd = bytes.length - 22;
    const cd = bytes.readUInt32LE(eocd + 16);
    const records: Buffer[] = [];
    let offset = cd;
    while (offset < eocd && bytes.readUInt32LE(offset) === 0x02014b50) {
        const length = 46 + bytes.readUInt16LE(offset + 28) + bytes.readUInt16LE(offset + 30) + bytes.readUInt16LE(offset + 32);
        records.push(Buffer.from(bytes.subarray(offset, offset + length)));
        offset += length;
    }
    return Buffer.concat([bytes.subarray(0, cd), ...records.reverse(), bytes.subarray(offset)]);
};

/** The locator's absolute offset names the ZIP64 end record that ends where the locator begins. */
const expectZip64LocatorNamesAdjacentRecord = (bytes: Uint8Array): void => {
    const keyed = Buffer.from(bytes);
    const view = readCentralDirectoryView(keyed);
    expect(view?.zip64).toBe(true);
    const locator = view!.eocdOffset - 20;
    expect(keyed.readUInt32LE(locator)).toBe(0x07064b50);
    const named = Number(keyed.readBigUInt64LE(locator + 8));
    expect(named).toBe(locator - 56);
    expect(keyed.readUInt32LE(named)).toBe(0x06064b50);
};

describe("archiveWithoutEntries (#629)", () => {
    it("drops the matching central records and keeps every other entry's own bytes", async () => {
        const bytes = buildRawZip([
            { name: "a.xml", content: "A" },
            { name: "word/x/../b.xml", content: "HOSTILE" },
            { name: "b.xml", content: "B" },
        ]);
        const subset = archiveWithoutEntries(bytes, HOSTILE);
        expect(subset).toBeDefined();
        expect(zipCentralDirectoryEntries(subset!)?.map((entry) => entry.name)).toEqual(["a.xml", "b.xml"]);
        expect(await contents(subset!)).toEqual({ "a.xml": "A", "b.xml": "B" });
    });

    it("brings back the entry a later hostile name hid in JSZip's view", async () => {
        const bytes = buildRawZip([
            { name: "word/document.xml", content: "REAL" },
            { name: "word/x/../document.xml", content: "HOSTILE" },
        ]);
        // JSZip resolves the `..` name onto the real one, and the later entry wins.
        expect((await contents(bytes))["word/document.xml"]).toBe("HOSTILE");
        expect(await contents(archiveWithoutEntries(bytes, HOSTILE)!)).toEqual({ "word/document.xml": "REAL" });
    });

    it("keeps the bytes before the archive and the EOCD comment", async () => {
        const archive = withEocdComment(
            buildRawZip([
                { name: "a.xml", content: "A" },
                { name: "/abs.xml", content: "HOSTILE" },
            ]),
            "kept comment",
        );
        const prefixed = Buffer.concat([Buffer.from("PREFIX"), archive]);
        const subset = archiveWithoutEntries(prefixed, (record) => record.name.startsWith("/"));
        expect(subset).toBeDefined();
        expect(Buffer.from(subset!.subarray(0, 6)).toString()).toBe("PREFIX");
        expect(Buffer.from(subset!.subarray(subset!.length - 12)).toString()).toBe("kept comment");
        expect(await contents(subset!)).toEqual({ "a.xml": "A" });
    });

    it("hands the predicate every central record, folder records included", () => {
        const bytes = buildRawZip([
            { name: "word/", directory: true },
            { name: "word/a.xml", content: "A" },
        ]);
        const seen: [string, boolean][] = [];
        archiveWithoutEntries(bytes, (record) => {
            seen.push([record.name, record.directory]);
            return false;
        });
        expect(seen).toEqual([
            ["word/", true],
            ["word/a.xml", false],
        ]);
    });

    it("returns the bytes as given when no record matches", () => {
        const bytes = buildRawZip([{ name: "a.xml", content: "A" }]);
        expect(archiveWithoutEntries(bytes, HOSTILE)).toBe(bytes);
    });

    it("returns a ZIP64 archive as given when no record matches", () => {
        const bytes = buildRawZip([{ name: "a.xml", content: "A" }], { zip64: true });
        expect(archiveWithoutEntries(bytes, HOSTILE)).toBe(bytes);
    });

    it("moves a ZIP64 tail with the directory it shrinks (#479)", async () => {
        const subset = archiveWithoutEntries(zip64WithHostile(), HOSTILE)!;
        expect(zipCentralDirectoryEntries(subset)?.map((entry) => entry.name)).toEqual(["a.xml", "b.xml"]);
        expect(await loadedContents(subset)).toEqual({ "a.xml": "A", "b.xml": "B" });
        const eocd = subset.length - 22;
        const locator = eocd - 20;
        // The locator names the record's new absolute position (#521), where the record's signature sits.
        const record = Number(Buffer.from(subset.subarray(locator, eocd)).readBigUInt64LE(8));
        const recordBytes = Buffer.from(subset.subarray(record, record + 56));
        expect(recordBytes.readUInt32LE(0)).toBe(0x06064b50);
        // Both entry counts name the two kept records, and the size names the
        // directory's shrunk span: where the record now sits minus where the
        // unchanged directory-offset field says the directory starts.
        expect(recordBytes.readBigUInt64LE(24)).toBe(2n);
        expect(recordBytes.readBigUInt64LE(32)).toBe(2n);
        expect(recordBytes.readBigUInt64LE(40)).toBe(BigInt(record - Number(recordBytes.readBigUInt64LE(48))));
        // The saturated 32-bit EOCD fields stay as stored; JSZip reads them
        // only to decide to follow the ZIP64 structures.
        expect(Buffer.from(subset.subarray(eocd)).readUInt16LE(8)).toBe(0xffff);
        expect(Buffer.from(subset.subarray(eocd)).readUInt16LE(10)).toBe(0xffff);
        expect(Buffer.from(subset.subarray(eocd)).readUInt32LE(12)).toBe(0xffffffff);
    });

    it("rewrites the 32-bit EOCD fields a ZIP64 archive kept unsaturated", async () => {
        const bytes = zip64WithHostile();
        // An unsaturated total-entries field: the ZIP64 record still owns the count both readers use.
        bytes.writeUInt16LE(3, bytes.length - 22 + 10);
        const subset = archiveWithoutEntries(bytes, HOSTILE)!;
        const eocd = Buffer.from(subset.subarray(subset.length - 22));
        expect(eocd.readUInt16LE(8)).toBe(0xffff);
        expect(eocd.readUInt16LE(10)).toBe(2);
        expect(eocd.readUInt32LE(12)).toBe(0xffffffff);
        expect(await loadedContents(subset)).toEqual({ "a.xml": "A", "b.xml": "B" });
    });

    it("rewrites an unsaturated disk-entry count and directory size on a ZIP64 archive", async () => {
        const bytes = zip64WithHostile();
        const eocdAt = bytes.length - 22;
        // Unsaturated disk-entries and directory size. The total stays saturated, so the ZIP64 record still owns it.
        bytes.writeUInt16LE(3, eocdAt + 8);
        bytes.writeUInt32LE(100, eocdAt + 12);
        const subset = archiveWithoutEntries(bytes, HOSTILE)!;
        const eocd = Buffer.from(subset.subarray(subset.length - 22));
        const locator = subset.length - 22 - 20;
        const record = Number(Buffer.from(subset.subarray(locator, subset.length - 22)).readBigUInt64LE(8));
        const recordBytes = Buffer.from(subset.subarray(record, record + 56));
        expect(eocd.readUInt16LE(8)).toBe(2);
        expect(eocd.readUInt16LE(10)).toBe(0xffff);
        expect(eocd.readUInt32LE(12)).toBe(Number(recordBytes.readBigUInt64LE(40)));
        expect(await loadedContents(subset)).toEqual({ "a.xml": "A", "b.xml": "B" });
    });

    it("keeps a prefix before a ZIP64 archive, rebasing its offsets (#467)", async () => {
        const prefixed = Buffer.concat([Buffer.from("PREFIX"), zip64WithHostile()]);
        const subset = archiveWithoutEntries(prefixed, HOSTILE)!;
        expect(Buffer.from(subset.subarray(0, 6)).toString()).toBe("PREFIX");
        expect(await loadedContents(subset)).toEqual({ "a.xml": "A", "b.xml": "B" });
    });

    it("returns undefined for an envelope it cannot rewrite", async () => {
        const entries = [
            { name: "a.xml", content: "A" },
            { name: "x/../a.xml", content: "HOSTILE" },
        ];
        // Bytes between the walked directory and the EOCD, which moving the EOCD up would drop (#468).
        expect(archiveWithoutEntries(withJunkBeforeEocd(buildRawZip(entries)), HOSTILE)).toBeUndefined();
        // A ZIP64 record not adjacent to its locator, which no rewrite can move (#479); the archive still opens.
        const immovable = withJunkBetweenZip64RecordAndLocator(buildRawZip(entries, { zip64: true }));
        await expect(JSZip.loadAsync(immovable)).resolves.toBeDefined();
        expect(archiveWithoutEntries(immovable, HOSTILE)).toBeUndefined();
        expect(archiveWithoutEntries(Buffer.from("not a zip archive"), HOSTILE)).toBeUndefined();
    });
});

/** Insert `chunk` immediately before the EOCD and count it in the directory size. */
const insertBeforeEocd = (bytes: Buffer, chunk: Buffer): Buffer => {
    const eocd = bytes.length - 22;
    const out = Buffer.concat([bytes.subarray(0, eocd), chunk, bytes.subarray(eocd)]);
    const shifted = eocd + chunk.length;
    out.writeUInt32LE(out.readUInt32LE(shifted + 12) + chunk.length, shifted + 12);
    return out;
};

const jszipNames = async (bytes: Uint8Array): Promise<string[]> => Object.keys((await JSZip.loadAsync(bytes)).files).sort();

describe("centralKeyedArchive (#445)", () => {
    it("returns the bytes as given when the central directory cannot be read", () => {
        const bytes = Buffer.from("not a zip");
        expect(centralKeyedArchive(bytes)).toBe(bytes);
    });

    it("returns the bytes as given when bytes sit between the directory and the EOCD", async () => {
        // A second EOCD whose declared size covers the first. The walk stops at
        // the first EOCD, so cdEnd is not that record (#478).
        const archive = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const eocd = archive.length - 22;
        const appended = Buffer.from(archive.subarray(eocd));
        appended.writeUInt32LE(archive.readUInt32LE(eocd + 12) + 22, 12);
        const bytes = Buffer.concat([archive, appended]);
        const view = readCentralDirectoryView(bytes);
        expect(view?.zip64).toBe(false);
        expect(view?.cdEnd).not.toBe(view?.eocdOffset);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
        expect(await loadedContents(bytes)).toEqual({ "local.xml": "A" });
    });

    it("returns the bytes as given when a local-header offset is the ZIP64 sentinel", async () => {
        const bytes = saturateLocalHeaderOffsets(buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]));
        expect(readCentralDirectoryView(bytes)?.records[0]?.localHeaderOffset).toBe(0xffffffff);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
    });

    it("rebuilds a local header so JSZip loads the central name, and keeps a non-unicode extra field", async () => {
        const built = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const cd = built.readUInt32LE(built.length - 22 + 16);
        const extra = Buffer.from([0x01, 0x00, 0x00, 0x00]);
        const bytes = insertBeforeEocd(built, extra);
        bytes.writeUInt16LE(extra.length, cd + 30);
        const keyed = centralKeyedArchive(bytes);
        expect(keyed).not.toBe(bytes);
        expect(Buffer.from(readCentralDirectoryView(keyed)!.records[0]!.extraBytes)).toEqual(extra);
        expect(await jszipNames(keyed)).toEqual(["central.xml"]);
        expect(await loadedContents(bytes)).toEqual({ "central.xml": "A" });
    });

    it("moves a ZIP64 tail when it rebuilds the local name", async () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }], { zip64: true });
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
        expect(await loadedContents(bytes)).toEqual({ "central.xml": "A" });
        const keyed = centralKeyedArchive(bytes);
        expect(keyed).not.toBe(bytes);
        expectZip64LocatorNamesAdjacentRecord(keyed);
    });

    it("names the moved ZIP64 end record when the EOCD comment carries another one (#521)", async () => {
        const bytes = withZip64EndRecordComment(
            buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }], { zip64: true }),
        );
        expect(bytes.readUInt32LE(bytes.length - 56)).toBe(0x06064b50);
        const keyed = centralKeyedArchive(bytes);
        expect(await loadedContents(bytes)).toEqual({ "central.xml": "A" });
        expect(keyed).not.toBe(bytes);
        expectZip64LocatorNamesAdjacentRecord(keyed);
    });

    it("strips an Info-ZIP unicode path field when it rebuilds the local name", async () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", unicodePath: "unicode.xml", content: "A" }]);
        const keyed = centralKeyedArchive(bytes);
        expect(keyed).not.toBe(bytes);
        expect(readCentralDirectoryView(keyed)!.records[0]!.extraBytes).toHaveLength(0);
        expect(await jszipNames(keyed)).toEqual(["central.xml"]);
    });

    it("rebuilds every local name when the streams do not overlap", async () => {
        const bytes = buildRawZip([
            { name: "central.xml", localName: "local.xml", content: "A" },
            { name: "b.xml", content: "B" },
        ]);
        const keyed = centralKeyedArchive(bytes);
        expect(keyed).not.toBe(bytes);
        expect(await jszipNames(keyed)).toEqual(["b.xml", "central.xml"]);
        expect(await loadedContents(bytes)).toEqual({ "b.xml": "B", "central.xml": "A" });
    });

    it("rebuilds local names when central records are not in local-header order (APPNOTE 4.4.1.3)", async () => {
        const ordered = buildRawZip([
            { name: "central.xml", localName: "local.xml", content: "A" },
            { name: "b.xml", content: "B" },
        ]);
        const bytes = reverseCentralRecordOrder(ordered);
        const records = readCentralDirectoryView(bytes)!.records;
        expect(records.map((record) => record.name)).toEqual(["b.xml", "central.xml"]);
        expect(records[0]!.localHeaderOffset).toBeGreaterThan(records[1]!.localHeaderOffset);
        const keyed = centralKeyedArchive(bytes);
        expect(await jszipNames(keyed)).toEqual(["b.xml", "central.xml"]);
        expect(await loadedContents(bytes)).toEqual({ "b.xml": "B", "central.xml": "A" });
        expect(keyed).not.toBe(bytes);
    });

    it("returns the bytes as given when a compressed stream overruns the directory", () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const cd = bytes.readUInt32LE(bytes.length - 22 + 16);
        // Ten bytes into the central directory, still inside the buffer.
        bytes.writeUInt32LE(bytes.readUInt32LE(cd + 20) + 10, cd + 20);
        expect(readCentralDirectoryView(bytes)?.records).toHaveLength(1);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
    });

    it("returns the bytes as given when two local streams overlap", async () => {
        const bytes = buildRawZip([
            { name: "central.xml", localName: "local.xml", content: "AAAA" },
            { name: "b.xml", content: "B" },
        ]);
        const first = bytes.readUInt32LE(bytes.length - 22 + 16);
        bytes.writeUInt32LE(bytes.readUInt32LE(first + 20) + 10, first + 20);
        // The central size claims ten bytes of the next local header. Both streams
        // still end before the directory, so the overlap — not an overrun — is what
        // keeps the archive as JSZip read it.
        expect(readCentralDirectoryView(bytes)?.records).toHaveLength(2);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["b.xml", "local.xml"]);
    });

    it("returns the bytes as given when the local record starts past the central directory", () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const cd = bytes.readUInt32LE(bytes.length - 22 + 16);
        const planted = Buffer.alloc(30);
        planted.writeUInt32LE(0x04034b50, 0);
        const out = Buffer.concat([bytes, planted]);
        // The central offset names the planted header, and the stream is empty, so the name itself overruns the directory.
        out.writeUInt32LE(bytes.length, cd + 42);
        out.writeUInt32LE(0, cd + 20);
        expect(readCentralDirectoryView(out)?.records).toHaveLength(1);
        expect(centralKeyedArchive(out)).toBe(out);
        expect(out.includes("local.xml")).toBe(true);
    });

    it("leaves a non-ZIP64 archive whose directory ends in a ZIP64 locator signature to JSZip", async () => {
        // The locator sits in the central record's comment, so the walked directory still ends at the EOCD.
        const comment = Buffer.alloc(20);
        comment.writeUInt32LE(0x07064b50, 0);
        const built = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const cd = built.readUInt32LE(built.length - 22 + 16);
        const bytes = insertBeforeEocd(built, comment);
        bytes.writeUInt16LE(comment.length, cd + 32);
        expect(readCentralDirectoryView(bytes)?.zip64).toBe(false);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
    });

    it("keeps a truncated extra field's unparseable tail verbatim", async () => {
        // Declared data length 10, one payload byte: the field does not fit, so the tail stays, 0x7075 included.
        const tail = Buffer.from([0x75, 0x70, 0x0a, 0x00, 0xff]);
        const agreed = buildRawZip([{ name: "a.xml", content: "A" }]);
        const agreedCd = agreed.readUInt32LE(agreed.length - 22 + 16);
        const agreedBytes = insertBeforeEocd(agreed, tail);
        agreedBytes.writeUInt16LE(tail.length, agreedCd + 30);
        expect(centralKeyedArchive(agreedBytes)).toBe(agreedBytes);

        const differed = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }]);
        const differedCd = differed.readUInt32LE(differed.length - 22 + 16);
        const differedBytes = insertBeforeEocd(differed, tail);
        differedBytes.writeUInt16LE(tail.length, differedCd + 30);
        const keyed = centralKeyedArchive(differedBytes);
        expect(keyed).not.toBe(differedBytes);
        expect(Buffer.from(readCentralDirectoryView(keyed)!.records[0]!.extraBytes)).toEqual(tail);
        expect(await jszipNames(keyed)).toEqual(["central.xml"]);
    });

    it("returns the bytes as given when two central records share one local header", async () => {
        const bytes = buildRawZip([
            { name: "central.xml", localName: "local.xml", content: "AAAA" },
            { name: "b.xml", content: "B" },
        ]);
        const first = bytes.readUInt32LE(bytes.length - 22 + 16);
        const second = first + 46 + bytes.readUInt16LE(first + 28) + bytes.readUInt16LE(first + 30) + bytes.readUInt16LE(first + 32);
        bytes.writeUInt32LE(bytes.readUInt32LE(first + 42), second + 42);
        bytes.writeUInt32LE(0, second + 20);
        expect(readCentralDirectoryView(bytes)?.records).toHaveLength(2);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
    });

    it("does not move a ZIP64 record that declares fewer than its fixed 44 bytes", async () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }], { zip64: true });
        const record = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
        // The declared size is 20. The directory-size field grows by the same 24 so the walk still starts on the records.
        bytes.writeUInt32LE(20, record + 4);
        bytes.writeBigUInt64LE(bytes.readBigUInt64LE(record + 40) + 24n, record + 40);
        expect(readCentralDirectoryView(bytes)?.zip64).toBe(true);
        expect(centralKeyedArchive(bytes)).toBe(bytes);
        expect(await jszipNames(bytes)).toEqual(["local.xml"]);
    });

    it("does not move a ZIP64 tail whose last locator signature sits in the EOCD comment", async () => {
        const bytes = buildRawZip([{ name: "central.xml", localName: "local.xml", content: "A" }], { zip64: true });
        const eocd = bytes.length - 22;
        const record = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
        const fake = Buffer.alloc(20);
        fake.writeUInt32LE(0x07064b50, 0);
        fake.writeBigUInt64LE(BigInt(record), 8);
        fake.writeUInt32LE(1, 16);
        const out = Buffer.concat([bytes, fake]);
        out.writeUInt16LE(fake.length, eocd + 20);
        expect(readCentralDirectoryView(out)?.zip64).toBe(true);
        expect(centralKeyedArchive(out)).toBe(out);
        expect(await jszipNames(out)).toEqual(["local.xml"]);
    });
});
