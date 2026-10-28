/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { type ZipCentralRecord, zipCentralDirectoryEntries } from "../src/lib/zip-central-directory";
import { archiveWithoutEntries, loadCentralNamedArchive } from "../src/lib/zip-central-names";
import { buildRawZip } from "./support/raw-zip";

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
