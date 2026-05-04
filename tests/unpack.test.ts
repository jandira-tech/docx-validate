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
import { unpack } from "../src/scripts/office/unpack.ts";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

interface DocxParts {
    document: string;
    contentTypes?: string;
    rels?: string;
}

async function buildDocx(target: string, parts: DocxParts): Promise<void> {
    const zip = new JSZip();
    zip.file(
        "[Content_Types].xml",
        parts.contentTypes ??
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                '<Default Extension="xml" ContentType="application/xml"/>' +
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
                "</Types>",
    );
    zip.file(
        "_rels/.rels",
        parts.rels ??
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
                "</Relationships>",
    );
    zip.file(
        "word/document.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${parts.document}</w:document>`,
    );
    const buf = await zip.generateAsync({ type: "nodebuffer" });
    await fs.writeFile(target, buf);
}

describe("unpack", () => {
    it("returns an error when the input file does not exist", async () => {
        await withTempDir(async (dir) => {
            const result = await unpack(path.join(dir, "missing.docx"), path.join(dir, "out"));
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/does not exist/);
        });
    });

    it("rejects unsupported file extensions", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "note.txt");
            await fs.writeFile(inputFile, "hello");
            const result = await unpack(inputFile, path.join(dir, "out"));
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/must be a \.docx, \.pptx, or \.xlsx/);
        });
    });

    it("rejects non-zip files masquerading as docx", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "fake.docx");
            await fs.writeFile(inputFile, "this is not a zip");
            const result = await unpack(inputFile, path.join(dir, "out"));
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/not a valid Office file/);
        });
    });

    it("extracts a DOCX, pretty-prints XML, merges runs, simplifies redlines", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "input.docx");
            const outputDir = path.join(dir, "out");

            const body =
                "<w:body><w:p>" +
                "<w:r><w:t>foo</w:t></w:r>" +
                "<w:r><w:t>bar</w:t></w:r>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t></w:r></w:ins>' +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>y</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            await buildDocx(inputFile, { document: body });

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);
            expect(result.message).toMatch(/Unpacked .*input\.docx \(\d+ XML files\)/);
            expect(result.message).toMatch(/simplified 1 tracked changes/);
            expect(result.message).toMatch(/merged 2 runs/);

            const docXml = await fs.readFile(path.join(outputDir, "word", "document.xml"), "utf8");
            expect((docXml.match(/<w:ins[ >]/g) ?? []).length).toBe(1);
            // After pretty-printing the source XML, mergeRuns combines two
            // <w:t> children into one whose data carries the indent whitespace
            // that the pretty-printer left between them. Accept either compact
            // output or whitespace-separated output.
            expect(docXml).toMatch(/foo\s*bar/);
            expect(docXml).toMatch(/x\s*y/);

            const contentTypes = await fs.readFile(path.join(outputDir, "[Content_Types].xml"), "utf8");
            expect(contentTypes).toMatch(/<Types/);
            const rels = await fs.readFile(path.join(outputDir, "_rels", ".rels"), "utf8");
            expect(rels).toMatch(/<Relationships/);
        });
    });

    it("escapes smart quotes after pretty-printing", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "smart.docx");
            const outputDir = path.join(dir, "out");

            const body = "<w:body><w:p><w:r><w:t>“hi” ‘there’</w:t></w:r></w:p></w:body>";
            await buildDocx(inputFile, { document: body });

            const result = await unpack(inputFile, outputDir, {
                mergeRuns: false,
                simplifyRedlines: false,
            });
            expect(result.ok).toBe(true);

            const docXml = await fs.readFile(path.join(outputDir, "word", "document.xml"), "utf8");
            expect(docXml).toMatch(/&#x201C;hi&#x201D; &#x2018;there&#x2019;/);
            expect(docXml).not.toMatch(/[“”‘’]/);
        });
    });

    it("skips DOCX-only post-processing for PPTX inputs", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "deck.pptx");
            const outputDir = path.join(dir, "out");

            const zip = new JSZip();
            zip.file(
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    "</Types>",
            );
            zip.file("ppt/presentation.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><presentation/>');
            const buf = await zip.generateAsync({ type: "nodebuffer" });
            await fs.writeFile(inputFile, buf);

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);
            expect(result.message).not.toMatch(/merged \d+ runs/);
            expect(result.message).not.toMatch(/simplified \d+ tracked changes/);
        });
    });

    it("respects --merge-runs false / --simplify-redlines false", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "input.docx");
            const outputDir = path.join(dir, "out");

            const body = "<w:body><w:p>" + "<w:r><w:t>foo</w:t></w:r>" + "<w:r><w:t>bar</w:t></w:r>" + "</w:p></w:body>";
            await buildDocx(inputFile, { document: body });

            const result = await unpack(inputFile, outputDir, {
                mergeRuns: false,
                simplifyRedlines: false,
            });
            expect(result.ok).toBe(true);
            expect(result.message).not.toMatch(/merged/);
            expect(result.message).not.toMatch(/simplified/);

            const docXml = await fs.readFile(path.join(outputDir, "word", "document.xml"), "utf8");
            expect((docXml.match(/<w:r[ >]/g) ?? []).length).toBe(2);
        });
    });
});
