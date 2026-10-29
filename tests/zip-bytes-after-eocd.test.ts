/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { countZipBytesAfterEocd } from "../src/scripts/office/zip-bytes-after-eocd";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.zip-bytes-after-eocd.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.zip-no-bytes-after-eocd.docx");
const CONCATENATED = path.join(HERE, "fixtures", "broken", "package.zip-concatenated-copy.docx");
const CONCATENATED_REMOVED = path.join(HERE, "fixtures", "working", "package.zip-concatenated-copy-removed.docx");

describe("word-valid ZIP bytes after the end of central directory", { timeout: 30000 }, () => {
    it("rejects the package Word recovers", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "zip-bytes-after-eocd")).toBe(true);
    });

    it("does not reject the same package once those bytes are removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not reject the same trailing byte once the ZIP comment length declares it", async () => {
        const clean = readFileSync(CLEAN);
        const declared = Buffer.concat([clean, Buffer.from([0x0a])]);
        declared.writeUInt16LE(1, clean.length - 2);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "declared-comment.docx");
            writeFileSync(file, declared);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
            expect(result.valid).toBe(true);
        });
    });

    // A concatenated package is the prefixed shape of #464: JSZip loads the
    // appended copy, the original archive becomes the prefix, and one check
    // reports it — zip-bytes-before-archive, never also zip-bytes-after-eocd.
    it("reports a second copy of the archive appended after the end-of-central-directory record once, as bytes before the archive", async () => {
        const result = await validate(CONCATENATED, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code)).toEqual([
            "zip-bytes-before-archive",
        ]);
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
    });

    it("does not reject that package once the appended copy is deleted", async () => {
        const result = await validate(CONCATENATED_REMOVED, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
        expect(result.valid).toBe(true);
    });

    // A prefix record is the record plus the bytes it actually holds: its
    // declared comment length reaching into the live archive does not make
    // the live archive trailing data (#491).
    const commentOverrunPrefix = async (commentLength: number): Promise<Buffer> => {
        const prefix = buildRawZip([{ name: "pad.bin", content: Buffer.alloc(100) }]);
        prefix.writeUInt16LE(commentLength, prefix.length - 2);
        return Buffer.concat([prefix, await buildDocx(MINIMAL_PARTS)]);
    };

    it("reports a prefix whose EOCD comment length reaches into the live archive once, as bytes before the archive", async () => {
        const bytes = await commentOverrunPrefix(40);
        expect(countZipBytesAfterEocd(bytes)).toBe(0);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "comment-overrun-prefix.docx");
            writeFileSync(file, bytes);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
            const prefix = buildRawZip([{ name: "pad.bin", content: Buffer.alloc(100) }]);
            const before = result.issues.filter((issue) => issue.code === "zip-bytes-before-archive");
            expect(before).toHaveLength(1);
            // The same count as the comment-0 control.
            expect(before[0]?.message).toContain(`${prefix.length} byte(s) precede the ZIP archive`);
        });
    });

    it("still reports genuine trailing data behind a comment-overrun prefix, measured from the loaded record", async () => {
        const bytes = Buffer.concat([await commentOverrunPrefix(40), Buffer.alloc(64)]);
        expect(countZipBytesAfterEocd(bytes)).toBe(64);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "overrun-prefix-and-trailing.docx");
            writeFileSync(file, bytes);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const after = result.issues.find((issue) => issue.code === "zip-bytes-after-eocd");
            expect(after?.severity).toBe("error");
            expect(after?.message).toContain("64 undeclared byte(s)");
        });
    });

    it("treats a prefix EOCD comment that runs past the buffer end as a prefix, not an overrun", async () => {
        const bytes = Buffer.concat([await commentOverrunPrefix(0xffff), Buffer.alloc(64)]);
        expect(countZipBytesAfterEocd(bytes)).toBe(64);
    });

    it("finds the earlier EOCD through a Uint8Array view with a nonzero byte offset", () => {
        const concatenated = readFileSync(CONCATENATED);
        const padded = new Uint8Array(concatenated.length + 32);
        padded.set(concatenated, 16);
        const view = padded.subarray(16, 16 + concatenated.length);
        // The earlier archive lies wholly in front of the archive JSZip loads,
        // so it is zip-bytes-before-archive's finding, not trailing data (#464).
        expect(countZipBytesAfterEocd(view)).toBe(0);
        expect(countZipBytesAfterEocd(readFileSync(CONCATENATED_REMOVED))).toBe(0);
    });

    it("does not report the concatenated copy on lenient", async () => {
        const result = await validate(CONCATENATED, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd" || issue.code === "zip-bytes-before-archive")).toBe(
            false,
        );
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the concatenated copy on strict", async () => {
        const result = await validate(CONCATENATED, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "zip-bytes-before-archive")).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "zip-bytes-after-eocd")).toBe(true);
    });
});
