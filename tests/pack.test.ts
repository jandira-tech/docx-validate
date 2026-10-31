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
import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { pack, runPackFromArgv } from "../src/scripts/office/pack";
import { unpack } from "../src/scripts/office/unpack";
import { validate } from "../src/scripts/office/validate";
import { buildDocx as buildMinimalDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** A part stored in UTF-16, with the byte order mark that says which way round it is. */
function utf16Part(text: string, endian: "le" | "be"): Buffer {
    const body = Buffer.from(text, "utf16le");
    if (endian === "be") body.swap16();
    return Buffer.concat([Buffer.from(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]), body]);
}

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

    it.each(["file", "directory"] as const)("refuses a symbolic link to a %s without writing output (#1199)", async (kind) => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");
            for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
                const file = path.join(inputDir, name);
                await fs.mkdir(path.dirname(file), { recursive: true });
                await fs.writeFile(file, content);
            }
            await fs.writeFile(
                path.join(inputDir, "[Content_Types].xml"),
                MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                    "</Types>",
                    '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
                ),
            );
            const mediaDir = path.join(inputDir, "word", "media");
            await fs.mkdir(mediaDir);
            const target = kind === "file" ? "data.bin" : "data";
            if (kind === "file") {
                await fs.writeFile(path.join(mediaDir, target), "data");
            } else {
                await fs.mkdir(path.join(mediaDir, target));
                await fs.writeFile(path.join(mediaDir, target, "data.bin"), "data");
            }
            const link = kind === "file" ? "alias.bin" : "alias";
            await fs.symlink(target, path.join(mediaDir, link), kind === "file" ? "file" : "dir");

            const result = await pack(inputDir, outputFile, { validate: false });

            expect(result.ok).toBe(false);
            expect(result.message).toContain(`word/media/${link}`);
            expect(result.message).toMatch(/^Error: .*symbolic link/);
            await expect(fs.stat(outputFile)).rejects.toMatchObject({ code: "ENOENT" });
        });
    });

    it("refuses a symbolic link created after the preflight scan (#1199, codex on #1200)", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");
            for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
                const file = path.join(inputDir, name);
                await fs.mkdir(path.dirname(file), { recursive: true });
                await fs.writeFile(file, content);
            }
            await fs.writeFile(path.join(dir, "outside.bin"), "outside");
            // The second listing of the input root is staging's: a writer adds the link between the two.
            const readdir = fs.readdir;
            let listings = 0;
            fs.readdir = (async (target: string, ...rest: unknown[]) => {
                if (target === inputDir && ++listings === 2)
                    await fs.symlink(path.join(dir, "outside.bin"), path.join(inputDir, "alias.bin"));
                return (readdir as (...args: unknown[]) => unknown).call(fs, target, ...rest);
            }) as typeof fs.readdir;
            try {
                const result = await pack(inputDir, outputFile, { validate: false });

                expect(listings).toBe(2);
                expect(result.ok).toBe(false);
                expect(result.message).toBe("Error: alias.bin is a symbolic link");
                await expect(fs.stat(outputFile)).rejects.toMatchObject({ code: "ENOENT" });
            } finally {
                fs.readdir = readdir;
            }
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

    it("condenses package-root relationships like document relationships (#1198)", async () => {
        await withTempDir(async (dir) => {
            const inputDir = path.join(dir, "src");
            const outputFile = path.join(dir, "out.docx");
            const rels =
                '<?xml version="1.0" encoding="UTF-8"?>\n' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
                '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
                "</Relationships>";
            await writeUnpacked(inputDir, {
                document: "<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body>",
                rels,
            });
            await fs.mkdir(path.join(inputDir, "word", "_rels"), { recursive: true });
            await fs.writeFile(path.join(inputDir, "word", "_rels", "document.xml.rels"), rels, "utf8");

            const result = await pack(inputDir, outputFile, { validate: false });
            expect(result.ok).toBe(true);

            const zip = await JSZip.loadAsync(await fs.readFile(outputFile));
            const documentRels = await zip.file("word/_rels/document.xml.rels")!.async("string");
            expect(documentRels).toMatch(/<Relationships[^>]*><Relationship /);
            expect(documentRels).toMatch(/\/><\/Relationships>/);
            const rootRels = await zip.file("_rels/.rels")!.async("string");
            expect(rootRels).toBe(documentRels);
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

    it.each(["le", "be"] as const)("round-trips a UTF-16%s package through unpack and pack (#866)", async (endian) => {
        await withTempDir(async (dir) => {
            const original = path.join(dir, "utf16.docx");
            const unpackedDir = path.join(dir, "unpacked");
            const packed = path.join(dir, "packed.docx");
            const document = MINIMAL_PARTS["word/document.xml"]!.replace('encoding="UTF-8"', 'encoding="UTF-16"');
            await fs.writeFile(original, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": utf16Part(document, endian) }));

            expect((await unpack(original, unpackedDir)).ok).toBe(true);
            expect((await pack(unpackedDir, packed)).ok).toBe(true);

            expect((await validate(packed, { profile: "word-valid" })).valid).toBe(true);

            const zip = await JSZip.loadAsync(await fs.readFile(packed));
            const bytes = Buffer.from(await zip.file("word/document.xml")!.async("uint8array"));
            expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]);
            const body = Buffer.from(bytes.subarray(2));
            if (endian === "be") body.swap16();
            expect(body.toString("utf16le")).toContain("Hello");
        });
    });

    it.each(["le", "be"] as const)(
        "packs a UTF-16%s part without a byte order mark under a declaration naming its byte order (#878)",
        async (endian) => {
            await withTempDir(async (dir) => {
                const inputDir = path.join(dir, "src");
                const packed = path.join(dir, "packed.docx");
                const declared = endian === "le" ? "UTF-16LE" : "UTF-16BE";
                await writeUnpacked(inputDir, { document: "<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>" });
                const docPath = path.join(inputDir, "word", "document.xml");
                const document = (await fs.readFile(docPath, "utf8")).replace('encoding="UTF-8"', `encoding="${declared}"`);
                await fs.writeFile(docPath, utf16Part(document, endian).subarray(2));

                expect((await pack(inputDir, packed, { validate: false })).ok).toBe(true);

                const zip = await JSZip.loadAsync(await fs.readFile(packed));
                const bytes = Buffer.from(await zip.file("word/document.xml")!.async("uint8array"));
                expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0x3c, 0x00] : [0x00, 0x3c]);
                if (endian === "be") bytes.swap16();
                const text = bytes.toString("utf16le");
                expect(text).toMatch(new RegExp(`^<\\?xml[^>]*encoding="${declared}"`));
                expect(text).toContain("Hello");
            });
        },
    );
});

/** Integration fixtures: exercise the real archive, filesystem and validators together. */
async function writePackageParts(dir: string, parts: Readonly<Record<string, string>> = MINIMAL_PARTS): Promise<void> {
    for (const [name, content] of Object.entries(parts)) {
        const file = path.join(dir, name);
        await fs.mkdir(path.dirname(file), { recursive: true });
        await fs.writeFile(file, content);
    }
}

describe("pack validation integration", () => {
    afterEach(() => vi.restoreAllMocks());

    it("rejects a regular file as the input directory", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input.docx");
            const output = path.join(dir, "out.docx");
            await fs.writeFile(input, await buildMinimalDocx());
            expect(await pack(input, output)).toEqual({ ok: false, message: `Error: ${input} is not a directory` });
            await expect(fs.access(output)).rejects.toThrow();
        });
    });

    it.each(["missing", "directory"])("refuses a %s original before packing", async (kind) => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const original = path.join(dir, "original.docx");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input);
            if (kind === "directory") await fs.mkdir(original);
            expect(await pack(input, output, { original })).toEqual({
                ok: false,
                message: `Error: original file not found: ${original}`,
            });
            await expect(fs.access(output)).rejects.toThrow();
        });
    });

    it("uses the explicit author as the default inference fallback", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const original = path.join(dir, "original.docx");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input);
            await fs.writeFile(original, await buildMinimalDocx());
            const result = await pack(input, output, { original, author: "Alice" });
            expect(result.ok).toBe(true);
            expect(result.validationLog).toBe("All validations PASSED!");
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("word/document.xml")!.async("string")).toContain("Hello");
        });
    });

    it("refuses an original comparison with no inferred or explicit author", async () => {
        await withTempDir(async (dir) => {
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
            const input = path.join(dir, "input");
            const original = path.join(dir, "original.docx");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input);
            await fs.writeFile(original, await buildMinimalDocx());
            const result = await pack(input, output, { original });
            expect(result.ok).toBe(false);
            expect(result.message).toBe(`Error: Validation failed for ${input}`);
            expect(result.validationLog).toContain("Could not infer author and `author` was not provided.");
            expect(stderr).toHaveBeenCalledWith(expect.stringContaining("Could not infer author"));
            await expect(fs.access(output)).rejects.toThrow();
        });
    });

    it.each([new Error("inference unavailable"), "inference unavailable"])("reports a rejected author callback (%s)", async (error) => {
        await withTempDir(async (dir) => {
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
            const input = path.join(dir, "input");
            const original = path.join(dir, "original.docx");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input);
            await fs.writeFile(original, await buildMinimalDocx());
            const result = await pack(input, output, {
                original,
                inferAuthorFunc: () => Promise.reject(error),
            });
            expect(result.ok).toBe(false);
            expect(result.validationLog).toContain("inference unavailable");
            expect(stderr).toHaveBeenCalledWith(expect.stringContaining("inference unavailable"));
            await expect(fs.access(output)).rejects.toThrow();
        });
    });

    it("returns the validator's issues and leaves the output absent when validation fails", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input, {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:t>Hello</w:t>", '<w:commentReference w:id="99"/>'),
            });
            const result = await pack(input, output);
            expect(result.ok).toBe(false);
            expect(result.message).toBe(`Error: Validation failed for ${input}`);
            expect(result.issues).toContainEqual(expect.objectContaining({ code: "comment-marker-missing", severity: "error" }));
            expect(result.validationLog).toBeUndefined();
            await expect(fs.access(output)).rejects.toThrow();
        });
    });

    it("auto-repairs whitespace preservation before packaging and reports the repair", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input, {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:t>Hello</w:t>", "<w:t> Hello </w:t>"),
            });
            const result = await pack(input, output);
            expect(result.ok).toBe(true);
            expect(result.validationLog).toBe("Auto-repaired 1 issue(s)\nAll validations PASSED!");
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("word/document.xml")!.async("string")).toContain('<w:t xml:space="preserve"> Hello </w:t>');
        });
    });

    it("validates a PPTX with its presentation validator", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.pptx");
            await writePackageParts(input, MINIMAL_PPTX_PARTS);
            const result = await pack(input, output);
            expect(result.ok).toBe(true);
            expect(result.validationLog).toBe("All validations PASSED!");
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("ppt/presentation.xml")!.async("string")).toContain("<p:presentation ");
        });
    });
});

describe("pack XML condensation integration", () => {
    afterEach(() => vi.restoreAllMocks());

    it("preserves unprefixed SpreadsheetML t text and binary parts when packing XLSX", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.xlsx");
            await writePackageParts(input, {
                "xl/sharedStrings.xml":
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n <si><t>   </t></si>\n</sst>',
            });
            const binary = Buffer.from([0, 0xff, 0x80, 0x42]);
            await fs.writeFile(path.join(input, "xl", "image.bin"), binary);
            const result = await pack(input, output);
            expect(result.ok).toBe(true);
            expect(result.validationLog).toBeUndefined();
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("xl/sharedStrings.xml")!.async("string")).toContain("<si><t>   </t></si></sst>");
            expect(await zip.file("xl/image.bin")!.async("nodebuffer")).toEqual(binary);
        });
    });

    it("leaves XML in an unsupported encoding byte-for-byte untouched", async () => {
        await withTempDir(async (dir) => {
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.docx");
            const document = MINIMAL_PARTS["word/document.xml"]!.replace('encoding="UTF-8"', 'encoding="macintosh"');
            await writePackageParts(input, { ...MINIMAL_PARTS, "word/document.xml": document });
            expect((await pack(input, output, { validate: false })).ok).toBe(true);
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("word/document.xml")!.async("nodebuffer")).toEqual(Buffer.from(document));
        });
    });

    it("reports malformed XML and rejects without producing an archive", async () => {
        await withTempDir(async (dir) => {
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input, { ...MINIMAL_PARTS, "word/document.xml": '<w:document xmlns:w="urn:w"><w:body></w:document>' });
            await expect(pack(input, output, { validate: false })).rejects.toThrow();
            expect(stderr).toHaveBeenCalledWith(expect.stringContaining("ERROR: Failed to parse document.xml:"));
            await expect(fs.access(output)).rejects.toThrow();
        });
    });
});

describe("pack CLI integration", () => {
    afterEach(() => vi.restoreAllMocks());

    it.each([
        [[], 1, "missing required argument"],
        [["--help"], 0, "Usage: pack"],
    ] as const)("returns a clean exit code for %j", async (argv, code, message) => {
        const stdout = vi.spyOn(process.stdout, "write").mockImplementation(() => true);
        const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
        expect(await runPackFromArgv(argv)).toBe(code);
        expect([...stdout.mock.calls, ...stderr.mock.calls].map(([text]) => String(text)).join("")).toContain(message);
    });

    it.each(["FALSE", "TrUe"])("parses --validate %s and prints the packing result", async (flag) => {
        await withTempDir(async (dir) => {
            const stdout = vi.spyOn(process.stdout, "write").mockImplementation(() => true);
            const input = path.join(dir, "input");
            const output = path.join(dir, "out.docx");
            await writePackageParts(input);
            expect(await runPackFromArgv([input, output, "--validate", flag])).toBe(0);
            expect(stdout).toHaveBeenCalledWith(`Successfully packed ${input} to ${output}\n`);
            if (flag === "TrUe") expect(stdout).toHaveBeenCalledWith("All validations PASSED!\n");
            else expect(stdout).toHaveBeenCalledTimes(1);
            const zip = await JSZip.loadAsync(await fs.readFile(output));
            expect(await zip.file("word/document.xml")!.async("string")).toContain("Hello");
        });
    });

    it("prints a failed packing result and returns exit code 1", async () => {
        await withTempDir(async (dir) => {
            const stdout = vi.spyOn(process.stdout, "write").mockImplementation(() => true);
            const input = path.join(dir, "missing");
            expect(await runPackFromArgv([input, path.join(dir, "out.docx")])).toBe(1);
            expect(stdout).toHaveBeenCalledWith(`Error: ${input} is not a directory\n`);
        });
    });
});
