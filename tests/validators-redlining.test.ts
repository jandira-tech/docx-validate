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

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli.ts";
import { NS } from "../src/lib/types.ts";
import { parseXml, serializeXml } from "../src/lib/xml-helpers.ts";
import {
    RedliningValidator,
    extractTextContent,
    getGitWordDiff,
    removeAuthorTrackedChanges,
    validateRedlining,
} from "../src/scripts/office/validators/redlining.ts";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function wrapDoc(body: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
}

async function writeUnpacked(dir: string, body: string): Promise<void> {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir, { recursive: true });
    await fs.writeFile(path.join(wordDir, "document.xml"), wrapDoc(body), "utf8");
}

async function buildOriginalDocx(dir: string, body: string): Promise<string> {
    const zip = new JSZip();
    zip.file("word/document.xml", wrapDoc(body));
    const buf = await zip.generateAsync({ type: "nodebuffer" });
    const docxPath = path.join(dir, "original.docx");
    await fs.writeFile(docxPath, buf);
    return docxPath;
}

describe("validateRedlining", () => {
    it("returns an error when modified document.xml is missing", async () => {
        await withTempDir(async (dir) => {
            const docx = await buildOriginalDocx(dir, "<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>");
            const unpacked = path.join(dir, "unpacked");
            await fs.mkdir(unpacked, { recursive: true });
            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            expect(result.valid).toBe(false);
            expect(result.issues[0].message).toMatch(/Modified document\.xml not found/);
        });
    });

    it("passes when no tracked changes by author exist", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body>";
            const docx = await buildOriginalDocx(dir, body);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, body);

            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("passes when author insertions and deletions strip down to original text", async () => {
        await withTempDir(async (dir) => {
            const original = "<w:body><w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p></w:body>";

            const modified =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">The </w:t></w:r>' +
                '<w:del w:id="1" w:author="Claude" w:date="2024-01-01T00:00:00Z">' +
                '<w:r><w:delText xml:space="preserve">quick </w:delText></w:r>' +
                "</w:del>" +
                '<w:ins w:id="2" w:author="Claude" w:date="2024-01-01T00:00:00Z">' +
                '<w:r><w:t xml:space="preserve">slow </w:t></w:r>' +
                "</w:ins>" +
                "<w:r><w:t>brown fox</w:t></w:r>" +
                "</w:p></w:body>";

            const docx = await buildOriginalDocx(dir, original);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, modified);

            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("fails with a git-style word diff when the modified text drifts past tracked-change wrappers", async () => {
        await withTempDir(async (dir) => {
            const original = "<w:body><w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p></w:body>";
            const modified =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">The </w:t></w:r>' +
                '<w:ins w:id="1" w:author="Claude"><w:r><w:t xml:space="preserve">very </w:t></w:r></w:ins>' +
                "<w:r><w:t>quick red fox</w:t></w:r>" +
                "</w:p></w:body>";

            const docx = await buildOriginalDocx(dir, original);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, modified);

            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("redlining/mismatch");
            expect(result.issues[0].message).toMatch(/Document text doesn't match/);
            expect(result.issues[0].message).toContain("Differences:");
        });
    });

    it("ignores tracked changes by other authors", async () => {
        await withTempDir(async (dir) => {
            const original = "<w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body>";
            const modified =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t xml:space="preserve">cruel </w:t></w:r></w:ins>' +
                "<w:r><w:t>world</w:t></w:r>" +
                "</w:p></w:body>";

            const docx = await buildOriginalDocx(dir, original);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, modified);

            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            expect(result.valid).toBe(true);
        });
    });
});

describe("removeAuthorTrackedChanges", () => {
    it("drops author insertions and unwraps author deletions converting w:delText to w:t", () => {
        const xml = wrapDoc(
            "<w:body><w:p>" +
                "<w:r><w:t>keep </w:t></w:r>" +
                '<w:ins w:author="Claude"><w:r><w:t>drop-ins</w:t></w:r></w:ins>' +
                '<w:del w:author="Claude"><w:r><w:delText>restored</w:delText></w:r></w:del>' +
                "</w:p></w:body>",
        );
        const doc = parseXml(xml);
        removeAuthorTrackedChanges(doc.documentElement!, "Claude");
        const after = serializeXml(doc);
        expect(after).not.toMatch(/<w:ins/);
        expect(after).not.toMatch(/<w:del[^a-zA-Z]/);
        expect(after).not.toMatch(/<w:delText/);
        expect(after).toMatch(/<w:t[^>]*>restored<\/w:t>/);
        expect(after).toContain("keep ");
    });

    it("leaves other authors' tracked changes intact", () => {
        const xml = wrapDoc(
            "<w:body><w:p>" +
                '<w:ins w:author="Alice"><w:r><w:t>alice-ins</w:t></w:r></w:ins>' +
                '<w:del w:author="Alice"><w:r><w:delText>alice-del</w:delText></w:r></w:del>' +
                "</w:p></w:body>",
        );
        const doc = parseXml(xml);
        removeAuthorTrackedChanges(doc.documentElement!, "Claude");
        const after = serializeXml(doc);
        expect(after).toMatch(/w:author="Alice"/);
        expect(after).toContain("alice-ins");
        expect(after).toContain("alice-del");
    });
});

describe("extractTextContent", () => {
    it("joins per-paragraph text values with newlines and skips empty paragraphs", () => {
        const xml = wrapDoc(
            "<w:body>" +
                "<w:p><w:r><w:t>first</w:t></w:r><w:r><w:t> line</w:t></w:r></w:p>" +
                "<w:p></w:p>" +
                "<w:p><w:r><w:t>second</w:t></w:r></w:p>" +
                "</w:body>",
        );
        const doc = parseXml(xml);
        expect(extractTextContent(doc.documentElement!)).toBe("first line\nsecond");
    });

    it("traverses nested w:t elements within tracked-change wrappers", () => {
        const xml = wrapDoc(
            "<w:body><w:p>" +
                '<w:ins w:author="Alice"><w:r><w:t>hi </w:t></w:r></w:ins>' +
                "<w:r><w:t>there</w:t></w:r>" +
                "</w:p></w:body>",
        );
        const doc = parseXml(xml);
        expect(extractTextContent(doc.documentElement!)).toBe("hi there");
    });
});

describe("getGitWordDiff", () => {
    it("returns null for identical inputs", async () => {
        const result = await getGitWordDiff("hello world\n", "hello world\n");
        expect(result).toBeNull();
    });

    it("returns diff content when the two strings differ", async () => {
        const result = await getGitWordDiff("the quick brown fox\n", "the quick red fox\n");
        expect(result).not.toBeNull();
        expect(result!).toMatch(/\[-.+-\]/);
        expect(result!).toMatch(/\{\+.+\+\}/);
    });
});

describe("redlining XML helpers", () => {
    it("uses the WordprocessingML namespace exposed by lib/types", () => {
        expect(NS.W).toBe("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    });
});

describe("RedliningValidator class (parity with Python class-based API)", () => {
    it("validate() delegates to validateRedlining and returns the same result", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body>";
            const docx = await buildOriginalDocx(dir, body);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, body);

            const validator = new RedliningValidator({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            const result = await validator.validate();
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("validate() surfaces mismatch errors via the class API", async () => {
        await withTempDir(async (dir) => {
            const original = "<w:body><w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p></w:body>";
            const modified =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">The </w:t></w:r>' +
                '<w:ins w:id="1" w:author="Claude"><w:r><w:t>slow </w:t></w:r></w:ins>' +
                "<w:r><w:t>red fox</w:t></w:r>" +
                "</w:p></w:body>";

            const docx = await buildOriginalDocx(dir, original);
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, modified);

            const validator = new RedliningValidator({
                unpackedDir: unpacked,
                originalDocx: docx,
                author: "Claude",
            });
            const result = await validator.validate();
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("redlining/mismatch");
        });
    });

    it("repair() returns 0 (no-op, matching Python behaviour)", async () => {
        const validator = new RedliningValidator({
            unpackedDir: "/tmp/unused",
            originalDocx: "/tmp/unused.docx",
            author: "Claude",
        });
        const repairs = await validator.repair();
        expect(repairs).toBe(0);
    });
});
