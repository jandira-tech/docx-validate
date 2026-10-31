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
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import { pack } from "../src/scripts/office/pack";
import { unpack } from "../src/scripts/office/unpack";
import { buildDocx as buildMinimalDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const MINIMAL_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!;

/** The minimal main part with its declaration naming the encoding it is stored in. */
function minimalDocument(encoding: string): string {
    return MINIMAL_DOCUMENT.replace('encoding="UTF-8"', `encoding="${encoding}"`);
}

/** A part stored in UTF-16, with the byte order mark that says which way round it is. */
function utf16Part(text: string, endian: "le" | "be"): Buffer {
    const body = Buffer.from(text, "utf16le");
    if (endian === "be") body.swap16();
    return Buffer.concat([Buffer.from(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]), body]);
}

/** The text of a UTF-16 part, decoded in the byte order it was stored in. */
function decodeUtf16(bytes: Buffer, endian: "le" | "be"): string {
    const body = Buffer.from(bytes.subarray(2));
    if (endian === "be") body.swap16();
    return body.toString("utf16le");
}

/** A part's document, without the XML declaration that names its encoding. */
function documentBody(text: string): string {
    return text.slice(text.indexOf("<w:document"));
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

describe("unpack", () => {
    it.each([" <!--c--> ", "\u00a0<!--c-->"])("preserves text beside comments through unpack and pack (#1219): %j", async (content) => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "text.docx");
            const outputDir = path.join(dir, "out");
            const packedFile = path.join(dir, "round-trip.docx");
            const text = `<w:t xml:space="preserve">${content}</w:t>`;
            await buildDocx(inputFile, { document: `<w:body><w:p><w:r>${text}</w:r></w:p></w:body>` });
            const textNodes = (xml: string) =>
                getElementsByTagNameNSAll(parseXml(xml), "*", "t").map((element) =>
                    Array.from(element.childNodes)
                        .filter((child) => child.nodeType === 3)
                        .map((child) => child.nodeValue),
                );
            const original = await JSZip.loadAsync(await fs.readFile(inputFile));
            const expected = textNodes(await original.file("word/document.xml")!.async("string"));

            expect((await unpack(inputFile, outputDir, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);
            const unpacked = textNodes(await fs.readFile(path.join(outputDir, "word/document.xml"), "utf8"));
            expect((await pack(outputDir, packedFile, { validate: false })).ok).toBe(true);
            const packed = await JSZip.loadAsync(await fs.readFile(packedFile));
            expect({ unpacked, packed: textNodes(await packed.file("word/document.xml")!.async("string")) }).toEqual({
                unpacked: expected,
                packed: expected,
            });
        });
    });

    it("preserves w:t text bytes through unpack and pack (#1219)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "text.docx");
            const outputDir = path.join(dir, "out");
            const packedFile = path.join(dir, "round-trip.docx");
            const textElements = [
                '<w:t xml:space="preserve"> a  b </w:t>',
                "<w:t>Hello</w:t>",
                '<w:t xml:space="preserve"> </w:t>',
                '<w:t xml:space="preserve">\t a\n  b\t </w:t>',
                "<w:t>&amp; &lt; &gt;</w:t>",
            ];
            await buildDocx(inputFile, {
                document: `<w:body><w:p>${textElements.map((text) => `<w:r>${text}</w:r>`).join("")}</w:p></w:body>`,
            });

            expect((await unpack(inputFile, outputDir, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);
            const unpackedXml = await fs.readFile(path.join(outputDir, "word/document.xml"), "utf8");
            // Compare lexical bytes, including whitespace and entity escapes,
            // rather than normalized DOM text or the outer ZIP envelope.
            const textBytes = (xml: string) => Buffer.from((xml.match(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g) ?? []).join(""));
            const expected = Buffer.from(textElements.join(""));
            expect(textBytes(unpackedXml)).toEqual(expected);

            expect((await pack(outputDir, packedFile, { validate: false })).ok).toBe(true);
            const zip = await JSZip.loadAsync(await fs.readFile(packedFile));
            expect(textBytes(await zip.file("word/document.xml")!.async("string"))).toEqual(expected);
        });
    });

    it("keeps raw JSZip's UTF-8 fallback for flag-less CP437 names (#480)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "cp437.docx");
            const outputDir = path.join(dir, "out");
            await fs.writeFile(
                inputFile,
                buildRawZip([
                    ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                    { name: "café.xml", nameBytes: Uint8Array.from([0x63, 0x61, 0x66, 0x82, 0x2e, 0x78, 0x6d, 0x6c]), content: "<x/>" },
                ]),
            );
            expect((await unpack(inputFile, outputDir, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);
            expect(await fs.readdir(outputDir)).toContain("caf�.xml");
            expect(await fs.readdir(outputDir)).not.toContain("café.xml");
        });
    });

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
            // Pretty-printing must not introduce text between merged runs.
            expect(docXml).toContain("<w:t>foobar</w:t>");
            expect(docXml).toContain("<w:t>xy</w:t>");

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

    it("refuses to extract a zip entry that escapes the output dir", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            await buildDocx(inputFile, { document: "<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>" });

            const originalLoadAsync = JSZip.loadAsync;
            JSZip.loadAsync = async function loadAsyncPatched(data, options) {
                const zip = await originalLoadAsync.call(this, data, options);
                zip.files["../../escaped.txt"] = {
                    name: "../../escaped.txt",
                    dir: false,
                    async: async () => Buffer.from("malicious"),
                } as unknown as JSZip.JSZipObject;
                return zip;
            };

            try {
                const result = await unpack(inputFile, path.join(dir, "out"), {
                    mergeRuns: false,
                    simplifyRedlines: false,
                });
                expect(result.ok).toBe(false);
                expect(result.message).toMatch(/Refusing to extract entry outside output dir/);
            } finally {
                JSZip.loadAsync = originalLoadAsync;
            }
        });
    });

    /** A minimal package plus one entry stored under `name`, written by a raw archiver. */
    const withEntry = async (target: string, name: string): Promise<void> => {
        const entries = Object.entries(MINIMAL_PARTS).map(([part, content]) => ({ name: part, content }));
        await fs.writeFile(target, buildRawZip([...entries, { name, content: "x" }]));
    };

    it("leaves no output directory behind when it created one and the extraction is refused (#675)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            await withEntry(inputFile, "../evil.txt");
            const out = path.join(dir, "new", "out");

            const result = await unpack(inputFile, out);
            expect(result).toEqual({ ok: false, message: "Error unpacking: Refusing to extract entry outside output dir: ../evil.txt" });
            await expect(fs.access(path.join(dir, "new"))).rejects.toThrow();
            await expect(fs.access(path.join(dir, "evil.txt"))).rejects.toThrow();
        });
    });

    it("keeps an existing output directory and its files when the extraction is refused (#675)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            await withEntry(inputFile, "word/../../evil.txt");
            const out = path.join(dir, "out");
            await fs.mkdir(out);
            await fs.writeFile(path.join(out, "notes.txt"), "mine");

            expect((await unpack(inputFile, out)).ok).toBe(false);
            expect(await fs.readdir(out)).toEqual(["notes.txt"]);
        });
    });

    // JSZip keeps no unsafeOriginalName for a folder entry, and keys its view
    // by the resolved name, so a later entry under that name replaces a
    // traversal record: only the central directory still holds the stored
    // names (#676 Codex P2s).
    it("refuses a stored folder name that leaves the output directory, writing nothing (#676)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            await withEntry(inputFile, "../../outside/");
            const out = path.join(dir, "out");

            const result = await unpack(inputFile, out);
            expect(result).toEqual({
                ok: false,
                message: "Error unpacking: Refusing to extract entry outside output dir: ../../outside/",
            });
            await expect(fs.access(out)).rejects.toThrow();
        });
    });

    it("refuses a traversal record a later entry under its resolved name shadows (#676)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            const entries = Object.entries(MINIMAL_PARTS).map(([part, content]) => ({ name: part, content }));
            await fs.writeFile(inputFile, buildRawZip([{ name: "../../word/document.xml", content: "x" }, ...entries]));
            const out = path.join(dir, "out");

            const result = await unpack(inputFile, out);
            expect(result).toEqual({
                ok: false,
                message: "Error unpacking: Refusing to extract entry outside output dir: ../../word/document.xml",
            });
            await expect(fs.access(out)).rejects.toThrow();
        });
    });

    it("prints a refused name's control characters as escapes (#675)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "slip.docx");
            await withEntry(inputFile, "../\u001b]0;owned\u0007\r\nevil.txt");

            const result = await unpack(inputFile, path.join(dir, "out"));
            expect(result.ok).toBe(false);
            expect(result.message).toBe(
                "Error unpacking: Refusing to extract entry outside output dir: ../\\u{1b}]0;owned\\u{7}\\u{d}\\u{a}evil.txt",
            );
        });
    });

    it("leaves no output directory behind for a file that is not a ZIP archive (#675)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "fake.docx");
            await fs.writeFile(inputFile, "not a zip");
            const out = path.join(dir, "out");

            expect((await unpack(inputFile, out)).ok).toBe(false);
            await expect(fs.access(out)).rejects.toThrow();
        });
    });

    it.each(["le", "be"] as const)("keeps a UTF-16%s main part in its own encoding (#866)", async (endian) => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "utf16.docx");
            const outputDir = path.join(dir, "out");
            const utf8File = path.join(dir, "utf8.docx");
            const utf8Out = path.join(dir, "utf8-out");
            await fs.writeFile(
                inputFile,
                await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": utf16Part(minimalDocument("UTF-16"), endian) }),
            );
            await fs.writeFile(utf8File, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": MINIMAL_DOCUMENT }));

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);
            expect((await unpack(utf8File, utf8Out)).ok).toBe(true);

            const bytes = await fs.readFile(path.join(outputDir, "word", "document.xml"));
            expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]);
            const decoded = decodeUtf16(bytes, endian);
            expect(decoded).not.toContain("\uFFFD");
            // The part reads exactly as the same package stored in UTF-8 does.
            const utf8Text = await fs.readFile(path.join(utf8Out, "word", "document.xml"), "utf8");
            expect(documentBody(decoded)).toBe(documentBody(utf8Text));
        });
    });

    it.each(["le", "be"] as const)("keeps a UTF-16%s main part's text when both helpers are off (#866)", async (endian) => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "utf16.docx");
            const outputDir = path.join(dir, "out");
            const utf8File = path.join(dir, "utf8.docx");
            const utf8Out = path.join(dir, "utf8-out");
            await fs.writeFile(
                inputFile,
                await buildMinimalDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf16Part(minimalDocument("UTF-16"), endian),
                }),
            );
            await fs.writeFile(utf8File, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": MINIMAL_DOCUMENT }));

            const result = await unpack(inputFile, outputDir, { mergeRuns: false, simplifyRedlines: false });
            expect(result.ok).toBe(true);
            expect((await unpack(utf8File, utf8Out, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);

            const bytes = await fs.readFile(path.join(outputDir, "word", "document.xml"));
            expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]);
            const decoded = decodeUtf16(bytes, endian);
            expect(decoded).not.toContain("\uFFFD");
            const utf8Text = await fs.readFile(path.join(utf8Out, "word", "document.xml"), "utf8");
            expect(documentBody(decoded)).toBe(documentBody(utf8Text));
        });
    });

    it("escapes a smart quote inside a UTF-16 part exactly as in a UTF-8 part (#866)", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:t>“hi” ‘there’</w:t>";
            const utf16File = path.join(dir, "utf16.docx");
            const utf8File = path.join(dir, "utf8.docx");
            await fs.writeFile(
                utf16File,
                await buildMinimalDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf16Part(minimalDocument("UTF-16").replace("<w:t>Hello</w:t>", body), "le"),
                }),
            );
            await fs.writeFile(
                utf8File,
                await buildMinimalDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_DOCUMENT.replace("<w:t>Hello</w:t>", body),
                }),
            );

            const utf16Out = path.join(dir, "utf16-out");
            const utf8Out = path.join(dir, "utf8-out");
            expect((await unpack(utf16File, utf16Out, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);
            expect((await unpack(utf8File, utf8Out, { mergeRuns: false, simplifyRedlines: false })).ok).toBe(true);

            const utf16Text = decodeUtf16(await fs.readFile(path.join(utf16Out, "word", "document.xml")), "le");
            const utf8Text = await fs.readFile(path.join(utf8Out, "word", "document.xml"), "utf8");
            // The declarations name the encoding each part is stored in; the
            // bodies must read the same.
            expect(utf16Text.slice(utf16Text.indexOf("<w:document"))).toBe(utf8Text.slice(utf8Text.indexOf("<w:document")));
            expect(utf16Text).toContain("&#x201C;hi&#x201D; &#x2018;there&#x2019;");
            expect(utf16Text).not.toMatch(/[“”‘’]/);
        });
    });

    it("leaves a part in an encoding it cannot decode byte-for-byte (#866)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "macintosh.docx");
            const outputDir = path.join(dir, "out");
            const document = minimalDocument("macintosh");
            await fs.writeFile(inputFile, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": document }));

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);

            const bytes = await fs.readFile(path.join(outputDir, "word", "document.xml"));
            expect(bytes.equals(Buffer.from(document, "utf8"))).toBe(true);
        });
    });

    it("leaves such a part byte-for-byte however long its declaration is (#881)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "macintosh-long.docx");
            const outputDir = path.join(dir, "out");
            // "Café" as Mac Roman stores it (é is 0x8E), behind a declaration longer than 200 bytes.
            const document = Buffer.from(
                MINIMAL_DOCUMENT.replace('encoding="UTF-8"', `${" ".repeat(240)}encoding="macintosh"`).replace("Hello", "Caf\u008e"),
                "latin1",
            );
            await fs.writeFile(inputFile, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": document }));

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);

            const bytes = await fs.readFile(path.join(outputDir, "word", "document.xml"));
            expect(bytes.equals(document)).toBe(true);
        });
    });

    // #1139: mergeRuns and simplifyRedlines walk with their own stack, so a
    // main document nested deeper than the call stack no longer fails unpack
    // (their own rows pin the 50 000-deep walk). #1159: such a part is past
    // libxml2's depth limit, so unpack now leaves it as stored instead.
    it("unpacks a main document nested 50 000 elements deep, leaving it as stored (#1139, #1159)", async () => {
        await withTempDir(async (dir) => {
            const inputFile = path.join(dir, "deep.docx");
            const outputDir = path.join(dir, "out");
            const depth = 50_000;
            const document = MINIMAL_DOCUMENT.replace(
                "<w:sectPr",
                `${"<w:customXml>".repeat(depth)}<w:p>` +
                    "<w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r>" +
                    '<w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t></w:r></w:ins>' +
                    '<w:ins w:id="2" w:author="Alice"><w:r><w:t>y</w:t></w:r></w:ins>' +
                    `</w:p>${"</w:customXml>".repeat(depth)}<w:sectPr`,
            );
            await fs.writeFile(inputFile, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": document }));

            const result = await unpack(inputFile, outputDir);
            expect(result.ok).toBe(true);
            expect(result.message).toContain("left word/document.xml as stored");
            expect(result.message).not.toMatch(/simplified|merged/);
            expect(await fs.readFile(path.join(outputDir, "word", "document.xml"), "utf-8")).toBe(document);
        });
    });
});

/**
 * #1159: libxml2 refuses a part nested deeper than 256 elements (the root
 * counts as one), so `xml-syntax` fails it and no validation reaches its DOM.
 * `unpack` used to parse, pretty-print and serialize it anyway, which grows
 * with the square of the depth and ran out of heap at 50 000. A part over the
 * limit is now left as stored, and the message names it.
 */
describe("unpack leaves a part deeper than the xml-syntax limit as stored (#1159)", () => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    /** A main part `depth` elements deep, the root included. */
    const nested = (depth: number, perElementNamespace: boolean): string => {
        const open = perElementNamespace ? `<w:customXml ${NS}>` : "<w:customXml>";
        return (
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>` +
            open.repeat(depth - 2) +
            "</w:customXml>".repeat(depth - 2) +
            "</w:body></w:document>"
        );
    };
    const unpackDocument = async (document: string) =>
        withTempDir(async (dir) => {
            const inputFile = path.join(dir, "deep.docx");
            const outputDir = path.join(dir, "out");
            await fs.writeFile(inputFile, await buildMinimalDocx({ ...MINIMAL_PARTS, "word/document.xml": document }));
            const result = await unpack(inputFile, outputDir);
            const stored = await fs.readFile(path.join(outputDir, "word", "document.xml"), "utf-8");
            const styles = await fs.readFile(path.join(outputDir, "word", "styles.xml"), "utf-8");
            return { result, stored, styles };
        });

    it.each([
        ["per-element xmlns", true],
        ["plain", false],
    ] as const)(
        "returns at 50 000 deep with the %s part byte-identical",
        async (_shape, perElementNamespace) => {
            const document = nested(50_000, perElementNamespace);
            const { result, stored } = await unpackDocument(document);
            expect(result.ok).toBe(true);
            expect(stored).toBe(document);
        },
        60_000,
    );

    it("leaves a part one element over the limit as stored, names it, and still processes the others", async () => {
        const document = nested(257, false);
        const { result, stored, styles } = await unpackDocument(document);
        expect(result.ok).toBe(true);
        expect(stored).toBe(document);
        expect(result.message).toContain("left word/document.xml as stored: nested deeper than 256 elements");
        expect(result.message).not.toContain("merged");
        expect(result.message).not.toContain("simplified");
        expect(styles).not.toBe(MINIMAL_PARTS["word/styles.xml"]);
    });

    it("still pretty-prints a part at the limit", async () => {
        const document = nested(256, false);
        const { result, stored } = await unpackDocument(document);
        expect(result.ok).toBe(true);
        expect(stored).not.toBe(document);
        expect(stored).toContain("\n  <w:body>");
        expect(result.message).not.toContain("as stored");
        expect(result.message).toContain("merged");
    });
});
