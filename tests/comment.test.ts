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

import * as crypto from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("node:crypto", async (importOriginal) => {
    const actual = await importOriginal<typeof import("node:crypto")>();
    return {
        ...actual,
        randomInt: vi.fn(actual.randomInt),
    };
});

import { withTempDir } from "../src/lib/run-cli";
import { parseXml } from "../src/lib/xml-helpers";
import { SECURE_LONG_HEX_EXCLUSIVE_MAX } from "../src/lib/secure-id";
import { addComment, generateHexId } from "../src/scripts/comment";

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const RELATIONSHIPS_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

async function makeFixture(dir: string): Promise<void> {
    await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
    await fs.writeFile(path.join(dir, "word", "_rels", "document.xml.rels"), RELS_XML, "utf-8");
    await fs.writeFile(path.join(dir, "[Content_Types].xml"), CT_XML, "utf-8");
}

/** A part stored in UTF-16, with the byte order mark that says which way round it is. */
function utf16Part(text: string): Buffer {
    return Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, "utf16le")]);
}

/** Every file under `dir`, keyed by its relative path, with its bytes. */
async function readTree(dir: string): Promise<Map<string, Buffer>> {
    const files = new Map<string, Buffer>();
    const walk = async (current: string): Promise<void> => {
        for (const entry of await fs.readdir(current, { withFileTypes: true })) {
            const full = path.join(current, entry.name);
            if (entry.isDirectory()) {
                await walk(full);
            } else {
                files.set(path.relative(dir, full), await fs.readFile(full));
            }
        }
    };
    await walk(dir);
    return files;
}

/** Rewrite a part's XML declaration to name an encoding this package cannot decode (#866). */
async function declareUndecodable(file: string): Promise<void> {
    const text = await fs.readFile(file, "utf-8");
    const undecodable = text.includes('encoding="UTF-8"')
        ? text.replace('encoding="UTF-8"', 'encoding="ISO-8859-1"')
        : text.replace("?>", ' encoding="ISO-8859-1"?>');
    await fs.writeFile(file, undecodable, "utf-8");
}

describe("generateHexId", () => {
    afterEach(() => {
        vi.mocked(crypto.randomInt).mockReset();
    });

    // [MS-OI29500] 2.6.2.3: paraId is an ST_LongHexNumber whose value MUST be
    // greater than 0 and less than 0x80000000. The generator feeds both w14:paraId
    // and w16cid:durableId, so a zero would be an out-of-spec id we wrote ourselves.
    it("draws from crypto.randomInt in the ST_LongHexNumber-safe range", () => {
        vi.mocked(crypto.randomInt).mockImplementation(() => 1);
        expect(generateHexId()).toBe("00000001");
        expect(crypto.randomInt).toHaveBeenCalledWith(1, SECURE_LONG_HEX_EXCLUSIVE_MAX);
    });

    it("stays below the ST_LongHexNumber cap at the RNG maximum", () => {
        vi.mocked(crypto.randomInt).mockImplementation(() => 0x7ffffffe);
        const value = Number.parseInt(generateHexId(), 16);
        expect(value).toBe(0x7ffffffe);
        expect(value).toBeGreaterThan(0);
        expect(value).toBeLessThan(0x80000000);
        // Also below the tighter durableId cap, since the same id seeds durableId.
        expect(value).toBeLessThan(SECURE_LONG_HEX_EXCLUSIVE_MAX);
    });

    it("returns a zero-padded 8-digit uppercase hex string", () => {
        vi.mocked(crypto.randomInt).mockImplementation(() => 1);
        expect(generateHexId()).toMatch(/^[0-9A-F]{8}$/);
    });
});

describe("addComment", () => {
    it("adds a top-level comment, seeds parts, and registers rels + content types", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            expect(result.paraId).toMatch(/^[0-9A-F]{8}$/);

            for (const name of ["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"]) {
                await fs.access(path.join(dir, "word", name));
            }

            const commentsXml = await fs.readFile(path.join(dir, "word", "comments.xml"), "utf-8");
            const commentsDom = parseXml(commentsXml);
            const comments = commentsDom.getElementsByTagName("w:comment");
            expect(comments.length).toBe(1);
            const c0 = comments.item(0)!;
            expect(c0.getAttribute("w:id")).toBe("0");
            expect(c0.getAttribute("w:author")).toBe("Test");
            expect(c0.getAttribute("w:initials")).toBe("C");
            expect(c0.getAttribute("w:date")).toBe("2026-05-03T12:00:00Z");
            expect(commentsXml).toContain("Hello world");

            const relsXml = await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8");
            expect(relsXml).toContain('Target="comments.xml"');
            expect(relsXml).toContain('Target="commentsExtended.xml"');
            expect(relsXml).toContain('Target="commentsIds.xml"');
            expect(relsXml).toContain('Target="commentsExtensible.xml"');

            const ctXml = await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8");
            expect(ctXml).toContain('PartName="/word/comments.xml"');
            expect(ctXml).toContain('PartName="/word/commentsExtended.xml"');

            const reply = await addComment({
                unpackedDir: dir,
                commentId: 1,
                text: "Reply text",
                parent: 0,
                author: "Reviewer",
                initials: "R",
                date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
            });
            expect(reply.message).toMatch(/^Added reply 1 \(para_id=[0-9A-F]{8}\)$/);

            const extXml = await fs.readFile(path.join(dir, "word", "commentsExtended.xml"), "utf-8");
            const extDom = parseXml(extXml);
            const exItems = extDom.getElementsByTagName("w15:commentEx");
            expect(exItems.length).toBe(2);
            const replyEx = exItems.item(1)!;
            expect(replyEx.getAttribute("w15:paraId")).toBe(reply.paraId);
            expect(replyEx.getAttribute("w15:paraIdParent")).toBe(result.paraId);
        });
    });

    it("escapes XML special characters in author, initials, and text", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: 'a & b < c > d "quoted"',
                author: "R&D <team>",
                initials: 'J"R',
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            // A raw '&' or '<' in any templated field produces malformed XML and
            // makes addComment throw while re-parsing comments.xml. A clean
            // paraId proves every field was escaped before templating.
            expect(result.paraId).toMatch(/^[0-9A-F]{8}$/);

            const xml = await fs.readFile(path.join(dir, "word", "comments.xml"), "utf-8");
            // Must still parse as well-formed XML.
            const dom = parseXml(xml);
            const comment = dom.getElementsByTagName("w:comment").item(0)!;
            expect(comment.getAttribute("w:author")).toBe("R&D <team>");
            expect(comment.getAttribute("w:initials")).toBe('J"R');
            const t = dom.getElementsByTagName("w:t").item(0)!;
            expect(t.textContent).toBe('a & b < c > d "quoted"');
        });
    });

    it("keeps a UTF-16 package's parts in their own encoding (#866)", async () => {
        await withTempDir(async (dir) => {
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(dir, "word", "_rels", "document.xml.rels"),
                utf16Part(RELS_XML.replace('encoding="UTF-8"', 'encoding="UTF-16"')),
            );
            await fs.writeFile(path.join(dir, "[Content_Types].xml"), utf16Part(CT_XML.replace('encoding="UTF-8"', 'encoding="UTF-16"')));

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            for (const name of ["word/_rels/document.xml.rels", "[Content_Types].xml"]) {
                const bytes = await fs.readFile(path.join(dir, name));
                expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
                const text = Buffer.from(bytes.subarray(2)).toString("utf16le");
                expect(text).toContain('encoding="UTF-16"');
            }
            const rels = Buffer.from((await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"))).subarray(2)).toString(
                "utf16le",
            );
            expect(rels).toContain('Target="comments.xml"');
        });
    });

    it.each([
        [
            "types .rels parts by Default",
            CT_XML.replace(
                '<Default Extension="xml"',
                `<Default Extension="rels" ContentType="${RELATIONSHIPS_TYPE}"/>\n  <Default Extension="xml"`,
            ),
            false,
        ],
        ["types no .rels part", CT_XML, true],
    ])(
        "relates the comments parts through a relationships part it creates when the main document has none, and the content types %s (#895)",
        async (_label, ctXml, needsOverride) => {
            await withTempDir(async (dir) => {
                await fs.mkdir(path.join(dir, "word"), { recursive: true });
                await fs.writeFile(path.join(dir, "[Content_Types].xml"), ctXml, "utf-8");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                const rels = parseXml(await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8"));
                const relationships = rels.getElementsByTagName("Relationship");
                const targets: string[] = [];
                for (let i = 0; i < relationships.length; i += 1) {
                    const rel = relationships.item(i)!;
                    targets.push(`${rel.getAttribute("Id")} ${rel.getAttribute("Target")}`);
                }
                expect(targets).toEqual([
                    "rId1 comments.xml",
                    "rId2 commentsExtended.xml",
                    "rId3 commentsIds.xml",
                    "rId4 commentsExtensible.xml",
                ]);

                // The new part must be typed, or the package gains an uncovered part.
                const ctAfter = await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8");
                expect(ctAfter.includes(`<Override PartName="/word/_rels/document.xml.rels" ContentType="${RELATIONSHIPS_TYPE}"/>`)).toBe(
                    needsOverride,
                );
                expect(ctAfter).toContain('PartName="/word/comments.xml"');
            });
        },
    );

    it("refuses a first comment when [Content_Types].xml is missing, changing nothing (#895)", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);
            await fs.rm(path.join(dir, "[Content_Types].xml"));

            const before = await readTree(dir);
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });

            expect(result.paraId).toBe("");
            expect(result.message).toBe(`Error: ${path.join(dir, "[Content_Types].xml")} not found`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it("reports an error when the unpacked DOCX has no word/ directory", async () => {
        await withTempDir(async (dir) => {
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });
            expect(result.paraId).toBe("");
            expect(result.message).toMatch(/^Error: .*word not found$/);
        });
    });

    it("reports an error when the parent comment id is unknown", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);
            await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "first",
                author: "TestAuthor",
                initials: "TA",
            });
            const reply = await addComment({
                unpackedDir: dir,
                commentId: 1,
                text: "orphan reply",
                parent: 99,
                author: "TestAuthor",
                initials: "TA",
            });
            expect(reply.paraId).toBe("");
            expect(reply.message).toBe("Error: Parent comment 99 not found");
        });
    });

    it.each(["word/_rels/document.xml.rels", "[Content_Types].xml"])(
        "refuses the first comment when %s declares an encoding it cannot decode, changing nothing (#879)",
        async (part) => {
            await withTempDir(async (dir) => {
                await makeFixture(dir);
                await declareUndecodable(path.join(dir, part));

                const before = await readTree(dir);
                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "x",
                    author: "Test",
                    initials: "T",
                });

                expect(result.paraId).toBe("");
                expect(result.message).toBe(`Error: ${path.join(dir, part)} is stored in an encoding this tool cannot decode`);
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );

    it.each(["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"])(
        "refuses a later comment when word/%s declares an encoding it cannot decode, changing nothing (#879)",
        async (name) => {
            await withTempDir(async (dir) => {
                await makeFixture(dir);
                await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "first",
                    author: "Test",
                    initials: "T",
                });

                const part = path.join(dir, "word", name);
                await declareUndecodable(part);
                const before = await readTree(dir);
                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 1,
                    text: "second",
                    author: "Test",
                    initials: "T",
                });

                expect(result.paraId).toBe("");
                expect(result.message).toBe(`Error: ${part} is stored in an encoding this tool cannot decode`);
                expect(await readTree(dir)).toEqual(before);

                // The comment id was not consumed: write the same text back
                // with a decodable UTF-8 declaration and comment 1 lands.
                await fs.writeFile(part, (await fs.readFile(part, "utf-8")).replace('encoding="ISO-8859-1"', 'encoding="UTF-8"'), "utf-8");
                const retry = await addComment({
                    unpackedDir: dir,
                    commentId: 1,
                    text: "second",
                    author: "Test",
                    initials: "T",
                });
                expect(retry.message).toMatch(/^Added comment 1 \(para_id=[0-9A-F]{8}\)$/);
                const commentsXml = await fs.readFile(path.join(dir, "word", "comments.xml"), "utf-8");
                expect(commentsXml).toContain('w:id="1"');
            });
        },
    );
});
