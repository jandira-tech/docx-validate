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
import { getTrackedChangeAuthorsFromPath, inferAuthor, simplifyRedlines } from "../src/scripts/office/helpers/simplify-redlines";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function writeDoc(dir: string, body: string): Promise<string> {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir, { recursive: true });
    const docPath = path.join(wordDir, "document.xml");
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
    await fs.writeFile(docPath, xml, "utf8");
    return docPath;
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

    it.each(["le", "be"] as const)("merges tracked changes in a UTF-16%s main part and keeps its encoding (#866)", async (endian) => {
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
            const document = `<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            await fs.writeFile(docPath, utf16Part(document, endian));

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(1);

            const bytes = await fs.readFile(docPath);
            expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]);
            const after = decodeUtf16(bytes, endian);
            expect((after.match(/<w:ins[ >]/g) ?? []).length).toBe(1);
            expect(after).toMatch(/foo/);
            expect(after).toMatch(/bar/);
        });
    });

    // #1139: the walk keeps its own stack, so a part nested deeper than the
    // call stack is simplified as a shallow one is.
    it("merges the tracked changes of a main document nested 50 000 elements deep (#1139)", async () => {
        await withTempDir(async (dir) => {
            const depth = 50_000;
            const docPath = await writeDoc(
                dir,
                `<w:body>${"<w:customXml>".repeat(depth)}<w:p>` +
                    '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                    '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                    `</w:p>${"</w:customXml>".repeat(depth)}</w:body>`,
            );
            expect(await simplifyRedlines(dir)).toEqual({ count: 1, message: "Simplified 1 tracked changes" });
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<w:ins[ >]/g) ?? []).length).toBe(1);
        });
    });

    // #1139: the stack walk starts at the root, as the recursive one did, so
    // a part whose root is a paragraph has its tracked changes merged.
    it("merges the tracked changes of a part whose root is a paragraph (#1139)", async () => {
        await withTempDir(async (dir) => {
            await fs.mkdir(path.join(dir, "word"), { recursive: true });
            const docPath = path.join(dir, "word", "document.xml");
            await fs.writeFile(
                docPath,
                `<?xml version="1.0" encoding="UTF-8"?>\n<w:p ${W}>` +
                    '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                    '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins></w:p>',
                "utf8",
            );
            expect(await simplifyRedlines(dir)).toEqual({ count: 1, message: "Simplified 1 tracked changes" });
        });
    });

    it("leaves an undecodable main part untouched and reports zero tracked changes", async () => {
        await withTempDir(async (dir) => {
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            const stored = Buffer.from(`<?xml version="1.0" encoding="macintosh"?>\n<w:document ${W}>${body}</w:document>`, "latin1");
            await fs.writeFile(docPath, stored);

            expect(await simplifyRedlines(dir)).toEqual({ count: 0, message: "Simplified 0 tracked changes" });
            expect((await fs.readFile(docPath)).equals(stored)).toBe(true);
        });
    });

    it("returns an error message for a malformed main part", async () => {
        await withTempDir(async (dir) => {
            const docPath = await writeDoc(dir, "<w:body><w:p>");
            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
            expect(result.message).toMatch(/^Error: /);
            expect(await fs.readFile(docPath, "utf8")).toContain("<w:body><w:p>");
        });
    });

    it("does not merge w:ins separated by non-whitespace text", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                "stray" +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<w:ins[ >]/g) ?? []).length).toBe(2);
        });
    });

    it("merges w:ins separated by whitespace-only text", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                "\n      " +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(1);
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<w:ins[ >]/g) ?? []).length).toBe(1);
            expect(after).toContain('<w:ins w:id="1" w:author="Alice"><w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r></w:ins>');
        });
    });

    it("does not merge a w:ins that carries no author attribute", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:ins w:id="1"><w:r><w:t>foo</w:t></w:r></w:ins>' +
                '<w:ins w:id="2" w:author="Alice"><w:r><w:t>bar</w:t></w:r></w:ins>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<w:ins[ >]/g) ?? []).length).toBe(2);
        });
    });

    it("merges tracked changes written in the default Word namespace", async () => {
        await withTempDir(async (dir) => {
            const defaultW = 'xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
            const body =
                "<w:body><w:p>" +
                `<ins ${defaultW} author="Alice"><w:r><w:t>foo</w:t></w:r></ins>` +
                `<ins ${defaultW} author="Alice"><w:r><w:t>bar</w:t></w:r></ins>` +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(1);
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<ins[ >]/g) ?? []).length).toBe(1);
            expect(after).toMatch(/foo/);
            expect(after).toMatch(/bar/);
        });
    });

    it("does not merge default-namespace tracked changes from different authors", async () => {
        await withTempDir(async (dir) => {
            const defaultW = 'xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
            const body =
                "<w:body><w:p>" +
                `<ins ${defaultW} author="Alice"><w:r><w:t>foo</w:t></w:r></ins>` +
                `<ins ${defaultW} author="Bob"><w:r><w:t>bar</w:t></w:r></ins>` +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await simplifyRedlines(dir);
            expect(result.count).toBe(0);
            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<ins[ >]/g) ?? []).length).toBe(2);
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

    it("infers from the modified document alone when the original docx is missing", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>new</w:t></w:r></w:ins>' + "</w:p></w:body>";
            await writeDoc(dir, body);

            expect(await inferAuthor(dir, path.join(dir, "missing.docx"), "Ritapolis")).toBe("Mallory");
        });
    });

    it("infers from the modified document alone when the original docx is not a zip", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(dir, "garbage.docx");
            await fs.writeFile(docxPath, Buffer.from("not a zip archive", "utf8"));
            await writeDoc(
                dir,
                "<w:body><w:p>" + '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>new</w:t></w:r></w:ins>' + "</w:p></w:body>",
            );

            expect(await inferAuthor(dir, docxPath, "Ritapolis")).toBe("Mallory");
        });
    });

    it("infers from the modified document alone when the original docx carries no word/document.xml", async () => {
        await withTempDir(async (dir) => {
            const zip = new JSZip();
            zip.file(
                "[Content_Types].xml",
                '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
            );
            const docxPath = path.join(dir, "empty.docx");
            await fs.writeFile(docxPath, await zip.generateAsync({ type: "nodebuffer" }));
            await writeDoc(
                dir,
                "<w:body><w:p>" + '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>new</w:t></w:r></w:ins>' + "</w:p></w:body>",
            );

            expect(await inferAuthor(dir, docxPath, "Ritapolis")).toBe("Mallory");
        });
    });

    it("infers from the modified document alone when the original document part is undecodable", async () => {
        await withTempDir(async (dir) => {
            const zip = new JSZip();
            zip.file(
                "word/document.xml",
                Buffer.from(
                    `<?xml version="1.0" encoding="macintosh"?>\n<w:document ${W}><w:body><w:p>` +
                        '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>old</w:t></w:r></w:ins>' +
                        "</w:p></w:body></w:document>",
                    "latin1",
                ),
            );
            const docxPath = path.join(dir, "macintosh.docx");
            await fs.writeFile(docxPath, await zip.generateAsync({ type: "nodebuffer" }));
            await writeDoc(
                dir,
                "<w:body><w:p>" + '<w:ins w:id="1" w:author="Mallory"><w:r><w:t>new</w:t></w:r></w:ins>' + "</w:p></w:body>",
            );

            expect(await inferAuthor(dir, docxPath, "Ritapolis")).toBe("Mallory");
        });
    });
});

describe("getTrackedChangeAuthorsFromPath", () => {
    it("returns an empty map for a missing file", async () => {
        await withTempDir(async (dir) => {
            expect(await getTrackedChangeAuthorsFromPath(path.join(dir, "word", "document.xml"))).toEqual(new Map());
        });
    });

    it("returns an empty map for a part whose encoding it cannot decode", async () => {
        await withTempDir(async (dir) => {
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            await fs.writeFile(
                docPath,
                Buffer.from(
                    `<?xml version="1.0" encoding="macintosh"?>\n<w:document ${W}><w:body><w:p>` +
                        '<w:ins w:id="1" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:ins>' +
                        "</w:p></w:body></w:document>",
                    "latin1",
                ),
            );
            expect(await getTrackedChangeAuthorsFromPath(docPath)).toEqual(new Map());
        });
    });

    it("returns an empty map for a malformed part", async () => {
        await withTempDir(async (dir) => {
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            await fs.writeFile(docPath, `<?xml version="1.0"?>\n<w:document ${W}><w:body>`, "utf8");
            expect(await getTrackedChangeAuthorsFromPath(docPath)).toEqual(new Map());
        });
    });

    it("skips tracked changes with no author attribute and counts w:del authors", async () => {
        await withTempDir(async (dir) => {
            const docPath = await writeDoc(
                dir,
                "<w:body><w:p>" +
                    '<w:ins w:id="1"><w:r><w:t>anon</w:t></w:r></w:ins>' +
                    '<w:del w:id="2" w:author="Dana"><w:r><w:delText>gone</w:delText></w:r></w:del>' +
                    "</w:p></w:body>",
            );
            expect(await getTrackedChangeAuthorsFromPath(docPath)).toEqual(new Map([["Dana", 1]]));
        });
    });
});
