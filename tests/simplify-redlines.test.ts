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

import { withTempDir } from "../src/lib/run-cli";
import { inferAuthor, simplifyRedlines } from "../src/scripts/office/helpers/simplify-redlines";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function writeDoc(dir: string, body: string): Promise<string> {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir, { recursive: true });
    const docPath = path.join(wordDir, "document.xml");
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
    await fs.writeFile(docPath, xml, "utf8");
    return docPath;
}

async function buildDocx(target: string, body: string): Promise<void> {
    const zip = new JSZip();
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`);
    const buf = await zip.generateAsync({ type: "nodebuffer" });
    await fs.writeFile(target, buf);
}

describe("simplifyRedlines", () => {
    it("returns an error when document.xml is missing", async () => {
        await withTempDir(async (dir) => {
            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
            expect(result.message).toMatch(/Error:.*document\.xml not found/);
        });
    });

    it("merges adjacent w:ins from the same author", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice" w:date="2025-01-01T00:00:00Z">' +
                "<w:r><w:t>foo</w:t></w:r>" +
                "</w:ins>" +
                '<w:ins w:id="2" w:author="Alice" w:date="2025-02-02T00:00:00Z">' +
                "<w:r><w:t>bar</w:t></w:r>" +
                "</w:ins>" +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(1);
            expect(result.message).toBe("Simplified 1 tracked changes");

            const after = await fs.readFile(docPath, "utf8");
            const insMatches = after.match(/<w:ins[ >]/g) ?? [];
            expect(insMatches.length).toBe(1);
            expect(after).toMatch(/foo/);
            expect(after).toMatch(/bar/);
        });
    });

    it("does not merge w:ins from different authors", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                '<w:ins w:id="2" w:author="Bob"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            const insMatches = after.match(/<w:ins[ >]/g) ?? [];
            expect(insMatches.length).toBe(2);
        });
    });

    it("does not merge w:ins separated by another element", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                "<w:r><w:t>middle</w:t></w:r>" +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
        });
    });

    it("merges adjacent w:del independently of w:ins", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:del w:id="1" w:author="Alice"><w:r><w:delText>x</w:delText></w:r></w:del>' +
                '<w:del w:id="2" w:author="Alice"><w:r><w:delText>y</w:delText></w:r></w:del>' +
                "</w:p></w:body>";
            await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(1);
        });
    });
});

describe("inferAuthor", () => {
    it("returns the default when there are no tracked changes in the modified doc", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(dir, "original.docx");
            await buildDocx(docxPath, "<w:body><w:p><w:r><w:t>none</w:t></w:r></w:p></w:body>");
            await writeDoc(dir, "<w:body><w:p><w:r><w:t>none</w:t></w:r></w:p></w:body>");

            const author = await inferAuthor(dir, docxPath, "Ritapolis");
            expect(author).toBe("Ritapolis");
        });
    });

    it("returns the lone new author when only one author added changes", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(dir, "original.docx");
            await buildDocx(docxPath, "<w:body><w:p><w:r><w:t>orig</w:t></w:r></w:p></w:body>");

            const body =
                "<w:body><w:p>" +
                "<w:r><w:t>orig</w:t></w:r>" +
                '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>new</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            await writeDoc(dir, body);

            const author = await inferAuthor(dir, docxPath, "Ritapolis");
            expect(author).toBe("Mallory");
        });
    });

    it("does not double-count authors already present in the original", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(dir, "original.docx");
            const origBody = "<w:body><w:p>" + '<w:ins w:id="1" w:author="Alice"><w:r><w:t>old</w:t></w:r></w:ins>' + "</w:p></w:body>";
            await buildDocx(docxPath, origBody);

            const body = "<w:body><w:p>" + '<w:ins w:id="1" w:author="Alice"><w:r><w:t>old</w:t></w:r></w:ins>' + "</w:p></w:body>";
            await writeDoc(dir, body);

            const author = await inferAuthor(dir, docxPath, "Ritapolis");
            expect(author).toBe("Ritapolis");
        });
    });

    it("throws when multiple authors added new changes", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(dir, "original.docx");
            await buildDocx(docxPath, "<w:body><w:p><w:r><w:t>orig</w:t></w:r></w:p></w:body>");

            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>a</w:t></w:r></w:ins>' +
                '<w:ins w:id="2" w:author="Bob"><w:r><w:t>b</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            await writeDoc(dir, body);

            await expect(inferAuthor(dir, docxPath, "Ritapolis")).rejects.toThrow(/Multiple authors/);
        });
    });
});
