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

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { mergeRuns } from "../src/scripts/office/helpers/merge-runs";

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

describe("mergeRuns", () => {
    it("returns an error when document.xml is missing", async () => {
        await withTempDir(async (dir) => {
            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);
            expect(result.message).toMatch(/Error:.*document\.xml not found/);
        });
    });

    it("merges two adjacent runs that share identical empty rPr", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + "<w:r><w:t>Hello </w:t></w:r>" + "<w:r><w:t>world</w:t></w:r>" + "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);
            expect(result.message).toBe("Merged 1 runs");

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(1);
            expect(after).toMatch(/Hello world/);
        });
    });

    it("preserves leading/trailing whitespace via xml:space when merging", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">trailing </w:t></w:r>' +
                '<w:r><w:t xml:space="preserve">space </w:t></w:r>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toMatch(/trailing space /);
            expect(after).toMatch(/xml:space="preserve"/);
        });
    });

    it("merges adjacent runs that share identical rPr", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:r w:rsidR="ABC"><w:rPr><w:b/></w:rPr><w:t>foo</w:t></w:r>' +
                '<w:r w:rsidR="DEF"><w:rPr><w:b/></w:rPr><w:t>bar</w:t></w:r>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(1);
            expect(after).toMatch(/foobar/);
            expect(after).not.toMatch(/w:rsidR/);
        });
    });

    it("does not merge runs with differing rPr", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>" +
                "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>" +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(2);
        });
    });

    it("removes proofErr elements before merging", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                "<w:r><w:t>foo</w:t></w:r>" +
                '<w:proofErr w:type="spellStart"/>' +
                "<w:r><w:t>bar</w:t></w:r>" +
                '<w:proofErr w:type="spellEnd"/>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).not.toMatch(/proofErr/);
            expect(after).toMatch(/foobar/);
        });
    });

    it.each(["le", "be"] as const)("merges runs in a UTF-16%s main part and keeps its encoding (#866)", async (endian) => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + "<w:r><w:t>Hello </w:t></w:r>" + "<w:r><w:t>world</w:t></w:r>" + "</w:p></w:body>";
            const document = `<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            await fs.writeFile(docPath, utf16Part(document, endian));

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const bytes = await fs.readFile(docPath);
            expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]);
            const after = decodeUtf16(bytes, endian);
            expect((after.match(/<w:r[ >]/g) ?? []).length).toBe(1);
            expect(after).toMatch(/Hello world/);
        });
    });

    it.each(["le", "be"] as const)(
        "keeps a UTF-16%s main part without a byte order mark, its declaration naming the byte order (#878)",
        async (endian) => {
            await withTempDir(async (dir) => {
                const body = "<w:body><w:p>" + "<w:r><w:t>Hello </w:t></w:r>" + "<w:r><w:t>world</w:t></w:r>" + "</w:p></w:body>";
                const declared = endian === "le" ? "UTF-16LE" : "UTF-16BE";
                const document = `<?xml version="1.0" encoding="${declared}" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
                const wordDir = path.join(dir, "word");
                await fs.mkdir(wordDir, { recursive: true });
                const docPath = path.join(wordDir, "document.xml");
                await fs.writeFile(docPath, utf16Part(document, endian).subarray(2));

                const result = await mergeRuns(dir);
                expect(result.count).toBe(1);

                const bytes = await fs.readFile(docPath);
                expect([...bytes.subarray(0, 2)]).toEqual(endian === "le" ? [0x3c, 0x00] : [0x00, 0x3c]);
                const after = decodeUtf16(Buffer.concat([Buffer.alloc(2), bytes]), endian);
                expect(after).toMatch(new RegExp(`^<\\?xml[^>]*encoding="${declared}"`));
                expect(after).toMatch(/Hello world/);
            });
        },
    );

    it("leaves a main part whose long declaration names an encoding it cannot decode untouched (#881)", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + "<w:r><w:t>Caf</w:t></w:r>" + "<w:r><w:t>\u008e</w:t></w:r>" + "</w:p></w:body>";
            // "Café" as Mac Roman stores it (é is 0x8E), behind a declaration longer than 200 bytes.
            const document = Buffer.from(
                `<?xml version="1.0"${" ".repeat(240)} encoding="macintosh" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`,
                "latin1",
            );
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            const docPath = path.join(wordDir, "document.xml");
            await fs.writeFile(docPath, document);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            expect((await fs.readFile(docPath)).equals(document)).toBe(true);
        });
    });

    // #1139: the walk keeps its own stack, so a part nested deeper than the
    // call stack is merged as a shallow one is.
    it("merges the runs of a main document nested 50 000 elements deep (#1139)", async () => {
        await withTempDir(async (dir) => {
            const depth = 50_000;
            const docPath = await writeDoc(
                dir,
                `<w:body>${"<w:customXml>".repeat(depth)}` +
                    "<w:p><w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r></w:p>" +
                    `${"</w:customXml>".repeat(depth)}</w:body>`,
            );
            expect(await mergeRuns(dir)).toEqual({ count: 1, message: "Merged 1 runs" });
            expect(await fs.readFile(docPath, "utf8")).toContain("<w:r><w:t>foobar</w:t></w:r>");
        });
    });

    // #1139: the stack walk starts at the root, as the recursive one did, so
    // a part whose root is a run has that run's rsid attributes stripped.
    it("strips the rsid attributes of a part whose root is a run (#1139)", async () => {
        await withTempDir(async (dir) => {
            await fs.mkdir(path.join(dir, "word"), { recursive: true });
            const docPath = path.join(dir, "word", "document.xml");
            await fs.writeFile(docPath, `<?xml version="1.0" encoding="UTF-8"?>\n<w:r ${W} w:rsidR="00AB0001"><w:t>x</w:t></w:r>`, "utf8");
            expect(await mergeRuns(dir)).toEqual({ count: 0, message: "Merged 0 runs" });
            expect(await fs.readFile(docPath, "utf8")).not.toContain("rsid");
        });
    });

    it("returns an error message for a malformed main part", async () => {
        await withTempDir(async (dir) => {
            const docPath = await writeDoc(dir, "<w:body><w:p>");
            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);
            expect(result.message).toMatch(/^Error: /);
            expect(await fs.readFile(docPath, "utf8")).toContain("<w:body><w:p>");
        });
    });

    it("merges runs separated by stray paragraph text, keeping that text", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>lead <w:r><w:t>a</w:t></w:r> tail<w:r><w:t>b</w:t></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("ab");
            expect(after).toContain("lead ");
            expect(after).toContain(" tail");
            expect((after.match(/<w:r[ >]/g) ?? []).length).toBe(1);
        });
    });

    it("does not merge a styled run into an unstyled one", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + "<w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r>" + "<w:r><w:t>b</w:t></w:r>" + "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            expect((after.match(/<w:r[ >]/g) ?? []).length).toBe(2);
            expect(after).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r>");
        });
    });

    it("keeps w:t elements separated by another element inside one run", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("<w:tab/>");
            expect((after.match(/<w:t>/g) ?? []).length).toBe(2);
        });
    });

    it("does not consolidate w:t elements separated by non-whitespace text", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>a</w:t>stray<w:t>b</w:t></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("stray");
            expect((after.match(/<w:t>/g) ?? []).length).toBe(2);
        });
    });

    it("consolidates w:t elements separated by whitespace inside one run", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>a</w:t>\n      <w:t>b</w:t></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("<w:t>ab</w:t>");
            expect((after.match(/<w:t>/g) ?? []).length).toBe(1);
        });
    });

    it("fills an empty leading w:t when its run merges", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t/></w:r><w:r><w:t>data</w:t></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("<w:t>data</w:t>");
            expect(after).not.toContain("<w:t/>");
            expect((after.match(/<w:t>/g) ?? []).length).toBe(1);
        });
    });

    it("drops a trailing empty w:t when its run merges", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p><w:r><w:t>a</w:t></w:r><w:r><w:t/></w:r></w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("<w:t>a</w:t>");
            expect(after).not.toContain("<w:t/>");
            expect((after.match(/<w:t>/g) ?? []).length).toBe(1);
        });
    });

    it("removes xml:space from a merged w:t that no longer borders a space", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + '<w:r><w:t xml:space="preserve">ab</w:t></w:r>' + "<w:r><w:t>cd</w:t></w:r>" + "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toContain("abcd");
            expect(after).not.toContain("xml:space");
        });
    });
});
