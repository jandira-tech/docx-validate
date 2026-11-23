/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Unit tests for the resource ceilings (`src/lib/resource-ceilings.ts`) and
 * the pre-inflate ZIP gate in `validate()`: the entry-count, per-part-size,
 * compression-ratio and cumulative-size ceilings are judged from the central
 * directory's declared fields before any entry is inflated.
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
    assertResourceCeilings,
    DEFAULT_RESOURCE_CEILINGS,
    resolveResourceCeilings,
    ZIP_COMPRESSION_RATIO_CEILING,
    ZIP_ENTRY_COUNT_CEILING,
    ZIP_PART_SIZE_CEILING,
    ZIP_TOTAL_UNCOMPRESSED_CEILING,
    zipCeilingIssues,
} from "../src/lib/resource-ceilings";
import { withTempDir } from "../src/lib/run-cli";
import { runValidateFromArgv, validate } from "../src/scripts/office/validate";
import { readCentralDirectoryView, type ZipCentralRecord } from "../src/lib/zip-central-directory";
import { ValidationContext, PackageViewError } from "../src/lib/validation-context";

const record = (name: string, compressedSize: number, uncompressedSize: number, directory = false): ZipCentralRecord => ({
    name,
    nameBytes: new TextEncoder().encode(name),
    extraBytes: new Uint8Array(0),
    commentBytes: new Uint8Array(0),
    recordOffset: 0,
    localHeaderOffset: 0,
    compressedSize,
    uncompressedSize,
    declaredCompressedSize: compressedSize,
    diskNumberStart: 0,
    directory,
});

describe("resolveResourceCeilings / assertResourceCeilings", () => {
    it("fills every ceiling with the default when nothing is given", () => {
        expect(resolveResourceCeilings(undefined)).toEqual(DEFAULT_RESOURCE_CEILINGS);
        expect(resolveResourceCeilings({})).toEqual(DEFAULT_RESOURCE_CEILINGS);
    });

    it("overrides individual ceilings and keeps the rest", () => {
        expect(resolveResourceCeilings({ maxEntries: 5, maxDepth: 100 })).toEqual({
            ...DEFAULT_RESOURCE_CEILINGS,
            maxEntries: 5,
            maxDepth: 100,
        });
    });

    it("rejects non-positive, non-finite and non-integer values", () => {
        expect(() => assertResourceCeilings({ maxEntries: 0 })).toThrow(RangeError);
        expect(() => assertResourceCeilings({ maxPartBytes: Number.NaN })).toThrow(RangeError);
        expect(() => assertResourceCeilings({ maxDepth: 2.5 })).toThrow(RangeError);
        expect(() => assertResourceCeilings({ maxCompressionRatio: 1.5 })).not.toThrow();
        expect(() => assertResourceCeilings("nope" as unknown as object)).toThrow(RangeError);
    });
});

describe("zipCeilingIssues", () => {
    it("reports nothing when every declared size is under the ceilings", () => {
        const records = [record("word/document.xml", 100, 500), record("word/", 0, 0, true)];
        expect(zipCeilingIssues(records, resolveResourceCeilings(undefined))).toEqual([]);
    });

    it("reports nothing when the directory cannot be read (the open failure reports it)", () => {
        expect(zipCeilingIssues(undefined, resolveResourceCeilings(undefined))).toEqual([]);
    });

    it("names the entry-count ceiling", () => {
        const records = [record("a.xml", 1, 10), record("b.xml", 1, 10), record("c/", 0, 0, true)];
        const issues = zipCeilingIssues(records, resolveResourceCeilings({ maxEntries: 1 }));
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: ZIP_ENTRY_COUNT_CEILING });
        expect(issues[0]!.message).toContain("2 file entries");
        expect(issues[0]!.message).toContain("maxEntries");
    });

    it("names the per-part size ceiling with the entry", () => {
        const records = [record("word/document.xml", 10, 1000), record("word/styles.xml", 10, 100)];
        const issues = zipCeilingIssues(records, resolveResourceCeilings({ maxPartBytes: 500 }));
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: ZIP_PART_SIZE_CEILING });
        expect(issues[0]!.message).toContain('"word/document.xml"');
        expect(issues[0]!.message).toContain("maxPartBytes");
    });

    it("names the compression-ratio ceiling only for entries over it", () => {
        const records = [record("word/document.xml", 10, 5000), record("[Content_Types].xml", 100, 200)];
        const issues = zipCeilingIssues(records, resolveResourceCeilings({ maxCompressionRatio: 100 }));
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: ZIP_COMPRESSION_RATIO_CEILING });
        expect(issues[0]!.message).toContain('"word/document.xml"');
        expect(issues[0]!.message).not.toContain("[Content_Types].xml");
    });

    it("names the cumulative uncompressed ceiling", () => {
        const records = [record("a.xml", 10, 600), record("b.xml", 10, 600)];
        const issues = zipCeilingIssues(records, resolveResourceCeilings({ maxTotalUncompressedBytes: 1000 }));
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: ZIP_TOTAL_UNCOMPRESSED_CEILING });
        expect(issues[0]!.message).toContain("1200");
    });
});

/** A minimal valid-enough docx for the gate tests. */
const buildDocx = async (entries: Record<string, string | Uint8Array>): Promise<Uint8Array> => {
    const zip = new JSZip();
    zip.file(
        "[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    );
    for (const [name, content] of Object.entries(entries)) zip.file(name, content);
    return zip.generateAsync({ type: "uint8array" });
};

describe("validate() ZIP ceiling gate (metadata before inflate)", () => {
    it("refuses an entry flood with the named ceiling, before staging", async () => {
        const entries: Record<string, string> = {};
        for (let i = 0; i < 50; i += 1) entries[`word/filler${i}.xml`] = "<a/>";
        const bytes = await buildDocx(entries);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "flood.docx");
            await writeFile(file, bytes);
            const result = await validate(file, { ceilings: { maxEntries: 10 } });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.code === ZIP_ENTRY_COUNT_CEILING && issue.severity === "error")).toBe(true);
            // A bounded refusal runs no checks and never stages the package.
            expect(result.metadata?.executedChecks).toEqual([]);
        });
    });

    it("refuses a part whose declared size is over the ceiling, before inflation", async () => {
        const bytes = await buildDocx({
            "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'.repeat(200),
        });
        await withTempDir(async (dir) => {
            const file = path.join(dir, "doc.docx");
            await writeFile(file, bytes);
            const ok = await validate(file, { ceilings: { maxPartBytes: 1_000_000_000 } });
            expect(ok.issues.some((issue) => issue.code?.startsWith("zip-"))).toBe(false);
            const refused = await validate(file, { ceilings: { maxPartBytes: 1000 } });
            expect(refused.valid).toBe(false);
            expect(
                refused.issues.some((issue) => issue.code === ZIP_PART_SIZE_CEILING && issue.message.includes("word/document.xml")),
            ).toBe(true);
        });
    });

    it("leaves a package under the ceilings alone", async () => {
        const bytes = await buildDocx({
            "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        });
        await withTempDir(async (dir) => {
            const file = path.join(dir, "doc.docx");
            await writeFile(file, bytes);
            const result = await validate(file);
            expect(result.issues.some((issue) => issue.code?.startsWith("zip-"))).toBe(false);
        });
    });

    it("rejects an invalid ceiling before reading anything", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "doc.docx");
            await writeFile(file, await buildDocx({}));
            await expect(validate(file, { ceilings: { maxEntries: 0 } })).rejects.toThrow(RangeError);
        });
    });
});

describe("ValidationContext.fromArchive pre-inflate gate", () => {
    it("refuses a breaching archive before inflating entries", async () => {
        const entries: Record<string, string> = {};
        for (let i = 0; i < 20; i += 1) entries[`word/filler${i}.xml`] = "<a/>";
        const bytes = await buildDocx(entries);
        // The central directory alone is enough: readCentralDirectoryView sees
        // the declared count without inflating anything.
        const view = readCentralDirectoryView(bytes);
        expect(view).toBeDefined();
        await expect(ValidationContext.fromArchive(bytes, { ceilings: { maxEntries: 5 } })).rejects.toThrow(PackageViewError);
        await expect(ValidationContext.fromArchive(bytes, { ceilings: { maxEntries: 5 } })).rejects.toThrow(/maxEntries/);
        // Under the default the same archive loads.
        const ctx = await ValidationContext.fromArchive(bytes);
        expect(ctx.partNames.length).toBeGreaterThan(0);
    });
});

/**
 * One stored entry, hand-built so the central record's size fields (and an
 * optional ZIP64 extra field) can declare anything. The local header stays
 * honest, so JSZip-style resolution (signed low 32 bits) still opens it.
 */
const storedArchive = (data: Uint8Array, central: { uncompressed: number; compressed: number; zip64?: bigint[] }): Uint8Array => {
    const name = new TextEncoder().encode("a.xml");
    const u16 = (v: number) => [v & 0xff, (v >>> 8) & 0xff];
    const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
    const u64 = (v: bigint) => [...u32(Number(v & 0xffffffffn)), ...u32(Number(v >> 32n))];
    const local = [
        ...u32(0x04034b50),
        ...u16(20),
        ...u16(0),
        ...u16(0),
        ...u16(0),
        ...u16(0),
        ...u32(0),
        ...u32(data.length),
        ...u32(data.length),
        ...u16(name.length),
        ...u16(0),
        ...name,
        ...data,
    ];
    const extra = central.zip64 ? [...u16(1), ...u16(central.zip64.length * 8), ...central.zip64.flatMap(u64)] : [];
    const cd = [
        ...u32(0x02014b50),
        ...u16(45),
        ...u16(45),
        ...u16(0),
        ...u16(0),
        ...u16(0),
        ...u16(0),
        ...u32(0),
        ...u32(central.compressed),
        ...u32(central.uncompressed),
        ...u16(name.length),
        ...u16(extra.length),
        ...u16(0),
        ...u16(0),
        ...u16(0),
        ...u32(0),
        ...u32(0),
        ...name,
        ...extra,
    ];
    const eocd = [...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(1), ...u16(1), ...u32(cd.length), ...u32(local.length), ...u16(0)];
    return new Uint8Array([...local, ...cd, ...eocd]);
};

describe("declared sizes are read unsigned and full width (r4651164162, r4651164138)", () => {
    const data = new TextEncoder().encode("<a/>");

    it("a 32-bit uncompressed size of 0x80000000 breaches maxPartBytes instead of reading negative", () => {
        const view = readCentralDirectoryView(storedArchive(data, { uncompressed: 0x80000000, compressed: data.length }));
        expect(view?.records[0]?.uncompressedSize).toBe(0x80000000);
        const codes = zipCeilingIssues(view?.records, { ...DEFAULT_RESOURCE_CEILINGS, maxPartBytes: 1_000_000 }).map((i) => i.code);
        expect(codes).toContain(ZIP_PART_SIZE_CEILING);
    });

    it("a ZIP64 uncompressed size of 2^32 + 1 keeps its high bits", () => {
        const view = readCentralDirectoryView(
            storedArchive(data, { uncompressed: 0xffffffff, compressed: data.length, zip64: [0x100000001n] }),
        );
        expect(view?.records[0]?.uncompressedSize).toBe(0x100000001);
        const codes = zipCeilingIssues(view?.records, DEFAULT_RESOURCE_CEILINGS).map((i) => i.code);
        expect(codes).toContain(ZIP_PART_SIZE_CEILING);
        expect(codes).toContain(ZIP_TOTAL_UNCOMPRESSED_CEILING);
    });

    it("the compression ratio uses the ZIP64-resolved compressed size", () => {
        const view = readCentralDirectoryView(
            storedArchive(data, { uncompressed: 0xffffffff, compressed: 0xffffffff, zip64: [5000n, BigInt(data.length)] }),
        );
        expect(view?.records[0]?.declaredCompressedSize).toBe(data.length);
        const codes = zipCeilingIssues(view?.records, { ...DEFAULT_RESOURCE_CEILINGS, maxCompressionRatio: 100 }).map((i) => i.code);
        expect(codes).toContain(ZIP_COMPRESSION_RATIO_CEILING);
    });
});

describe("CLI --max-* operands (r4651164168)", () => {
    it.each(["", "   "])("rejects the empty operand %j instead of falling back to the default", async (operand) => {
        const err: string[] = [];
        const stderrWrite = process.stderr.write.bind(process.stderr);
        process.stderr.write = ((chunk: string | Uint8Array): boolean => {
            err.push(typeof chunk === "string" ? chunk : Buffer.from(chunk).toString("utf-8"));
            return true;
        }) as typeof process.stderr.write;
        let code: number;
        try {
            code = await runValidateFromArgv(["does-not-exist.docx", "--max-entries", operand]);
        } finally {
            process.stderr.write = stderrWrite;
        }
        expect(code).toBe(1);
        expect(err.join("")).toContain("Invalid --max-entries");
    });
});
