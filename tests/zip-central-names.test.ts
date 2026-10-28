/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { type ZipCentralRecord, zipCentralDirectoryEntries } from "../src/lib/zip-central-directory";
import { archiveWithoutEntries } from "../src/lib/zip-central-names";
import { buildRawZip } from "./support/raw-zip";

const HOSTILE = (record: ZipCentralRecord): boolean => record.name.includes("..");

const contents = async (bytes: Uint8Array): Promise<Record<string, string>> => {
    const zip = await JSZip.loadAsync(bytes);
    const out: Record<string, string> = {};
    for (const [name, entry] of Object.entries(zip.files)) {
        if (!entry.dir) out[name] = await entry.async("string");
    }
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

    it("returns undefined for an envelope it cannot rewrite", () => {
        const entries = [
            { name: "a.xml", content: "A" },
            { name: "x/../a.xml", content: "HOSTILE" },
        ];
        expect(archiveWithoutEntries(buildRawZip(entries, { zip64: true }), HOSTILE)).toBeUndefined();
        expect(archiveWithoutEntries(Buffer.from("not a zip archive"), HOSTILE)).toBeUndefined();
    });
});
