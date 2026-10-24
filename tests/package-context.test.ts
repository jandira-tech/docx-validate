/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { contextFromArchive, contextFromDirectory } from "../src/scripts/office/package-context";

afterEach(() => {
    vi.restoreAllMocks();
});

describe("contextFromDirectory", () => {
    it("names each part by its forward-slash path under the directory", async () => {
        await withTempDir(async (dir) => {
            await mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await writeFile(path.join(dir, "word", "document.xml"), "<w:document/>");
            await writeFile(path.join(dir, "word", "_rels", "document.xml.rels"), "<Relationships/>");
            const context = contextFromDirectory(dir);
            expect(context.source).toBe("directory");
            expect([...context.partNames].sort()).toEqual(["word/_rels/document.xml.rels", "word/document.xml"]);
            expect(new TextDecoder().decode(context.bytes("word/document.xml"))).toBe("<w:document/>");
        });
    });
});

describe("contextFromArchive", () => {
    it("opens a readable archive with no open error", async () => {
        const zip = new JSZip();
        zip.file("word/document.xml", "<w:document/>");
        const { context, openError } = await contextFromArchive(await zip.generateAsync({ type: "uint8array" }));
        expect(openError).toBeUndefined();
        expect(context.source).toBe("archive");
        expect(context.has("word/document.xml")).toBe(true);
    });

    it("keeps unreadable bytes, with no parts, and the reason the archive did not open", async () => {
        const bytes = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
        const { context, openError } = await contextFromArchive(bytes);
        expect(openError).toEqual(expect.any(String));
        expect(openError!.length).toBeGreaterThan(0);
        expect(context.partNames).toEqual([]);
        expect(context.archiveBytes).toBe(bytes);
    });

    it("reports a rejection that is not an Error by its string form", async () => {
        vi.spyOn(ValidationContext, "fromArchive").mockRejectedValue("not a zip");
        const { context, openError } = await contextFromArchive(new Uint8Array([1, 2, 3]));
        expect(openError).toBe("not a zip");
        expect(context.partNames).toEqual([]);
    });
});
