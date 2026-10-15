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

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.zip-bytes-after-eocd.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.zip-no-bytes-after-eocd.docx");
const CONCATENATED = path.join(HERE, "fixtures", "broken", "package.zip-concatenated-copy.docx");
const CONCATENATED_REMOVED = path.join(HERE, "fixtures", "working", "package.zip-concatenated-copy-removed.docx");

describe("word-valid ZIP bytes after the end of central directory", () => {
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

    it("rejects a second copy of the archive appended after the end-of-central-directory record", async () => {
        const result = await validate(CONCATENATED, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "zip-bytes-after-eocd")).toBe(true);
    });

    it("does not reject that package once the appended copy is deleted", async () => {
        const result = await validate(CONCATENATED_REMOVED, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("finds the earlier EOCD through a Uint8Array view with a nonzero byte offset", () => {
        const concatenated = readFileSync(CONCATENATED);
        const padded = new Uint8Array(concatenated.length + 32);
        padded.set(concatenated, 16);
        const view = padded.subarray(16, 16 + concatenated.length);
        expect(countZipBytesAfterEocd(view)).toBe(concatenated.length / 2);
        expect(countZipBytesAfterEocd(readFileSync(CONCATENATED_REMOVED))).toBe(0);
    });

    it.each(["lenient", "strict"] as const)("does not report the concatenated copy on %s", async (profile) => {
        const result = await validate(CONCATENATED, { profile });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "zip-bytes-after-eocd")).toBe(false);
    });
});
