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
import { pack } from "../src/scripts/office/pack";
import { unpack } from "../src/scripts/office/unpack";

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

async function writeUnpacked(dir: string, parts: DocxParts): Promise<void> {
    await fs.mkdir(path.join(dir, "_rels"), { recursive: true });
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.writeFile(
        path.join(dir, "[Content_Types].xml"),
        parts.contentTypes ??
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                '<Default Extension="xml" ContentType="application/xml"/>' +
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
                "</Types>",
        "utf8",
    );
    await fs.writeFile(
        path.join(dir, "_rels", ".rels"),
        parts.rels ??
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
                "</Relationships>",
        "utf8",
    );
    await fs.writeFile(
        path.join(dir, "word", "document.xml"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${parts.document}</w:document>`,
        "utf8",
    );
}

describe("pack", () => {
    it("returns an error when input directory does not exist", async () => {
        await withTempDir(async (dir) => {
            const result = await pack(path.join(dir, "missing"), path.join(dir, "out.docx"), {
                validate: false,
            });
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/is not a directory/);
        });
    });

    it("rejects unsupported output extensions", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>",
            });
            const result = await pack(inputDir, path.join(dir, "out.txt"), {
                validate: false,
            });
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/must be a \.docx, \.pptx, or \.xlsx/);
        });
    });

    it("packs a directory into a .docx zip without validation", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");

            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body>",
            });

            const result = await pack(inputDir, outputFile, { validate: false });
            expect(result.ok).toBe(true);
            expect(result.message).toMatch(/Successfully packed .* to .*out\.docx/);

            const buf = await fs.readFile(outputFile);
            const zip = await JSZip.loadAsync(buf);
            expect(zip.file("word/document.xml")).not.toBeNull();
            expect(zip.file("[Content_Types].xml")).not.toBeNull();
            expect(zip.file("_rels/.rels")).not.toBeNull();

            const docXml = await zip.file("word/document.xml")!.async("string");
            expect(docXml).toMatch(/<w:t>hello<\/w:t>/);
        });
    });

    it("condenses whitespace-only text and strips comments from non-:t elements", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");

            const docBody =
                "<w:body>\n  <w:p>\n    <!-- this is a comment -->\n    <w:r>\n      <w:t>keep this    </w:t>\n    </w:r>\n  </w:p>\n</w:body>";
            await writeUnpacked(inputDir, { document: docBody });

            const result = await pack(inputDir, outputFile, { validate: false });
            expect(result.ok).toBe(true);

            const zip = await JSZip.loadAsync(await fs.readFile(outputFile));
            const docXml = await zip.file("word/document.xml")!.async("string");
            expect(docXml).not.toMatch(/<!--/);
            // <w:t> contents preserved verbatim, including trailing whitespace.
            expect(docXml).toMatch(/<w:t[^>]*>keep this {4}<\/w:t>/);
            // Non-:t whitespace-only text nodes between elements get stripped, so
            // <w:body><w:p>...</w:p></w:body> are now adjacent without indentation.
            expect(docXml).toMatch(/<w:body><w:p>/);
            expect(docXml).toMatch(/<\/w:p><\/w:body>/);
        });
    });

    it("round-trips through unpack(pack(unpacked)) for a tiny synthetic DOCX", async () => {
        await withTempDir(async (dir) => {
            const originalDocx = path.join(dir, "original.docx");
            const unpackedDir = path.join(dir, "unpacked");
            const repackedDocx = path.join(dir, "repacked.docx");
            const finalUnpacked = path.join(dir, "final");

            await buildDocx(originalDocx, {
                document: "<w:body><w:p><w:r><w:t>roundtrip text</w:t></w:r></w:p></w:body>",
            });

            const u1 = await unpack(originalDocx, unpackedDir, {
                mergeRuns: false,
                simplifyRedlines: false,
            });
            expect(u1.ok).toBe(true);

            const p = await pack(unpackedDir, repackedDocx, { validate: false });
            expect(p.ok).toBe(true);

            const u2 = await unpack(repackedDocx, finalUnpacked, {
                mergeRuns: false,
                simplifyRedlines: false,
            });
            expect(u2.ok).toBe(true);

            const finalDoc = await fs.readFile(path.join(finalUnpacked, "word", "document.xml"), "utf8");
            expect(finalDoc).toMatch(/roundtrip text/);
        });
    });

    it("creates parent directories for the output file as needed", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>",
            });
            const nestedOut = path.join(dir, "deeply", "nested", "out.docx");

            const result = await pack(inputDir, nestedOut, { validate: false });
            expect(result.ok).toBe(true);

            await expect(fs.stat(nestedOut)).resolves.toBeDefined();
        });
    });

    it("propagates the validate flag (skips validation when --validate false)", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");
            const original = path.join(dir, "missing-original.docx");

            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>",
            });

            // original does not exist; pack with validate=false must succeed.
            const result = await pack(inputDir, outputFile, {
                original,
                validate: false,
            });
            expect(result.ok).toBe(true);
            expect(result.validationLog).toBeUndefined();
        });
    });

    it("invokes inferAuthorFunc callback instead of default inferAuthor when provided", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");
            const originalFile = path.join(dir, "original.docx");

            await buildDocx(originalFile, {
                document: "<w:body><w:p><w:r><w:t>text</w:t></w:r></w:p></w:body>",
            });
            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>text</w:t></w:r></w:p></w:body>",
            });

            let callCount = 0;
            let capturedUnpackedDir: string | undefined;
            let capturedOriginalDocx: string | undefined;

            const customInferAuthor = (unpackedDir: string, originalDocx: string): string => {
                callCount += 1;
                capturedUnpackedDir = unpackedDir;
                capturedOriginalDocx = originalDocx;
                return "CustomAuthor";
            };

            const result = await pack(inputDir, outputFile, {
                original: originalFile,
                validate: true,
                inferAuthorFunc: customInferAuthor,
            });

            expect(result.ok).toBe(true);
            expect(callCount).toBe(1);
            expect(capturedUnpackedDir).toBe(path.resolve(inputDir));
            expect(capturedOriginalDocx).toBe(path.resolve(originalFile));
        });
    });
});
