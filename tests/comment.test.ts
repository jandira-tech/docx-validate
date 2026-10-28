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
import { canonicalPartName, CONTENT_TYPES_NAMESPACE, parseContentTypes, resolveContentType } from "../src/lib/content-types";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import { SECURE_LONG_HEX_EXCLUSIVE_MAX } from "../src/lib/secure-id";
import { addComment, generateHexId } from "../src/scripts/comment";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const RELATIONSHIPS_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const MAIN_DOCUMENT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const OFFICE_DOCUMENT_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STRICT_OFFICE_DOCUMENT_RELATIONSHIP_TYPE = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const COMMENTS_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const STRICT_COMMENTS_RELATIONSHIP_TYPE = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const HYPERLINK_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const COMMENTS_EXTENDED_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_IDS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";

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

/** The package relationships part holding one relationship with `attributes`. */
const packageRels = (attributes: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" ${attributes}/>\n</Relationships>`;

const officeDocumentRels = (target: string): string => packageRels(`Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="${target}"`);

/** The four relationships a first comment appends, with ids counting from `firstRid`. */
const firstCommentRelationships = (firstRid: number): string[][] => [
    [`rId${firstRid}`, COMMENTS_RELATIONSHIP_TYPE, "comments.xml", ""],
    [`rId${firstRid + 1}`, COMMENTS_EXTENDED_RELATIONSHIP_TYPE, "commentsExtended.xml", ""],
    [`rId${firstRid + 2}`, COMMENTS_IDS_RELATIONSHIP_TYPE, "commentsIds.xml", ""],
    [`rId${firstRid + 3}`, COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE, "commentsExtensible.xml", ""],
];

/** A content-types stream whose `.rels` Default names `relsType`, with `overrides` spliced in (#901). */
const contentTypes = (relsType: string, overrides: readonly string[] = []): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="${relsType}"/>${overrides.join("")}</Types>`;

/** A content-types stream whose namespace is bound to the `ct:` prefix, with a `.rels` Default naming `relsType` and `overrides` spliced in (#923). */
const prefixedContentTypes = (relsType: string, overrides: readonly string[] = []): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ct:Types xmlns:ct="${CONTENT_TYPES_NAMESPACE}"><ct:Default Extension="rels" ContentType="${relsType}"/>${overrides.join("")}</ct:Types>`;

/** The four comments parts a first comment types, each with the type it gets (#923). */
const COMMENT_PART_TYPES: readonly [string, string][] = [
    ["/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"],
    ["/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"],
    ["/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"],
    ["/word/commentsExtensible.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"],
];

/** The minimal main document with a paired marker triple for comment 0. */
const markedMainDocument = (): string =>
    MINIMAL_PARTS["word/document.xml"]!.replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>',
    );

/** Zip an unpacked tree into a DOCX at `file`. */
async function packTree(dir: string, file: string): Promise<void> {
    const parts: Record<string, Buffer> = {};
    for (const [name, bytes] of await readTree(dir)) {
        parts[name.split(path.sep).join("/")] = bytes;
    }
    await fs.writeFile(file, await buildDocx(parts));
}

/**
 * An unpacked tree whose `_rels/.rels` holds `relsXml`, with `word/`, a
 * `[Content_Types].xml` covering the stored parts, and the minimal main
 * document stored at `mainPart` when given. Without `relsDefault` the package
 * relationships part is typed by Override, so a relationships part a first
 * comment creates must be typed too (#895, #906).
 */
async function makePackageFixture(dir: string, relsXml: string, mainPart?: string, options: { relsDefault?: boolean } = {}): Promise<void> {
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.mkdir(path.join(dir, "_rels"), { recursive: true });
    await fs.writeFile(path.join(dir, "_rels", ".rels"), relsXml, "utf-8");
    const relsCoverage =
        options.relsDefault === false
            ? `<Override PartName="/_rels/.rels" ContentType="${RELATIONSHIPS_TYPE}"/>`
            : `<Default Extension="rels" ContentType="${RELATIONSHIPS_TYPE}"/>`;
    const mainOverride = mainPart === undefined ? "" : `<Override PartName="/${mainPart}" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`;
    await fs.writeFile(
        path.join(dir, "[Content_Types].xml"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${relsCoverage}${mainOverride}</Types>`,
        "utf-8",
    );
    if (mainPart !== undefined) {
        await fs.mkdir(path.dirname(path.join(dir, ...mainPart.split("/"))), { recursive: true });
        await fs.writeFile(path.join(dir, ...mainPart.split("/")), MINIMAL_PARTS["word/document.xml"]!, "utf-8");
    }
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

/** A part's text with its closing root tag dropped, so it decodes but does not parse (#905). */
function malformed(text: string): string {
    return text.replace(/<\/[^\s>]+>\s*$/, "");
}

/** Every [PartName, ContentType] pair of the Override elements naming `partName` under OPC equivalence. */
function overridesNaming(dom: ReturnType<typeof parseXml>, partName: string): [string, string][] {
    const found: [string, string][] = [];
    const overrides = dom.getElementsByTagName("Override");
    for (let i = 0; i < overrides.length; i += 1) {
        const override = overrides.item(i)!;
        if (canonicalPartName(override.getAttribute("PartName") ?? "") === canonicalPartName(partName)) {
            found.push([override.getAttribute("PartName") ?? "", override.getAttribute("ContentType") ?? ""]);
        }
    }
    return found;
}

/** Every [namespaceURI, prefix, PartName, ContentType] row of the content-types Override elements naming `partName` under OPC equivalence (#923). */
function namespacedOverridesNaming(dom: ReturnType<typeof parseXml>, partName: string): [string, string, string, string][] {
    const found: [string, string, string, string][] = [];
    for (const override of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override")) {
        if (canonicalPartName(override.getAttribute("PartName") ?? "") === canonicalPartName(partName)) {
            found.push([
                override.namespaceURI ?? "",
                override.prefix ?? "",
                override.getAttribute("PartName") ?? "",
                override.getAttribute("ContentType") ?? "",
            ]);
        }
    }
    return found;
}

/** Every [Id, Type, Target, TargetMode] of the relationships a relationships part holds. */
function relationshipRows(dom: ReturnType<typeof parseXml>): string[][] {
    const rows: string[][] = [];
    const relationships = dom.getElementsByTagName("Relationship");
    for (let i = 0; i < relationships.length; i += 1) {
        const rel = relationships.item(i)!;
        rows.push([
            rel.getAttribute("Id") ?? "",
            rel.getAttribute("Type") ?? "",
            rel.getAttribute("Target") ?? "",
            rel.getAttribute("TargetMode") ?? "",
        ]);
    }
    return rows;
}

/** Every [Extension, ContentType] pair of the Default elements naming `extension`. */
function defaultsFor(dom: ReturnType<typeof parseXml>, extension: string): [string, string][] {
    const found: [string, string][] = [];
    const defaults = dom.getElementsByTagName("Default");
    for (let i = 0; i < defaults.length; i += 1) {
        const def = defaults.item(i)!;
        if ((def.getAttribute("Extension") ?? "").toLowerCase() === extension.toLowerCase()) {
            found.push([def.getAttribute("Extension") ?? "", def.getAttribute("ContentType") ?? ""]);
        }
    }
    return found;
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

    it.each([
        ["word/document.xml", "/word/_rels/document.xml.rels"],
        ["word/main.xml", "/word/_rels/main.xml.rels"],
    ])(
        "answers a .rels Default of another type with an Override for the relationships part it creates beside %s (#901)",
        async (mainPart, relsPartName) => {
            await withTempDir(async (dir) => {
                // The review's tree: the `.rels` Default names another type, so
                // coverage alone left the created part mistyped while
                // addComment reported success.
                await makePackageFixture(dir, officeDocumentRels(mainPart), mainPart);
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    contentTypes("application/xml", [
                        `<Override PartName="/_rels/.rels" ContentType="${RELATIONSHIPS_TYPE}"/>`,
                        `<Override PartName="/${mainPart}" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    ]),
                    "utf-8",
                );

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The Default of another type is left alone and answered with
                // exactly one Override under the part's own name.
                const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
                expect(overridesNaming(dom, relsPartName)).toEqual([[relsPartName, RELATIONSHIPS_TYPE]]);
                expect(defaultsFor(dom, "rels")).toEqual([["rels", "application/xml"]]);
            });
        },
    );

    it.each(["/word/_rels/document.xml.rels", "/Word/_rels/Document.xml.rels"])(
        "corrects an Override of another type naming the created part as %s, adding no second Override (#901)",
        async (partName) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    contentTypes(RELATIONSHIPS_TYPE, [
                        `<Override PartName="${partName}" ContentType="application/xml"/>`,
                        `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    ]),
                    "utf-8",
                );

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The Override is rewritten in place, whatever OPC-equivalent
                // spelling names the part, never answered with a second one.
                const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
                expect(overridesNaming(dom, partName)).toEqual([[partName, RELATIONSHIPS_TYPE]]);
                expect(defaultsFor(dom, "rels")).toEqual([["rels", RELATIONSHIPS_TYPE]]);
            });
        },
    );

    it("adds no Override when a .rels Default already names the relationships media type in another case (#901)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
            const oddCase = "Application/Vnd.Openxmlformats-Package.Relationships+XML";
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(oddCase, [`<Override PartName="/word/main.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`]),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The media type already matches, whatever its case: the created
            // part needs no Override, the Default keeps its spelling, and the
            // comments parts are still typed.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(overridesNaming(dom, "/word/_rels/main.xml.rels")).toEqual([]);
            expect(defaultsFor(dom, "rels")).toEqual([["rels", oddCase]]);
            expect(overridesNaming(dom, "/word/comments.xml")).toHaveLength(1);
        });
    });

    it("never retypes a relationships part that already existed, however it is typed (#901)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/_rels/document.xml.rels" ContentType="application/xml"/>`,
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                ]),
                "utf-8",
            );
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await fs.writeFile(path.join(dir, "word", "_rels", "document.xml.rels"), RELS_XML, "utf-8");

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The part existed, so its typing is the package's own business:
            // only the comments parts are typed.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(overridesNaming(dom, "/word/_rels/document.xml.rels")).toEqual([["/word/_rels/document.xml.rels", "application/xml"]]);
        });
    });

    it(
        "draws no content-type finding for the relationships part it types over a wrong .rels Default (#901)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                // A control, not the defect: the content-type checks judge
                // whether a part is covered, not by which type, so they accept
                // the mistyped part too. The wrong-Default tree, with the comment
                // referenced from the main document so the packed package reads it.
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    contentTypes("application/xml", [
                        `<Override PartName="/_rels/.rels" ContentType="${RELATIONSHIPS_TYPE}"/>`,
                        `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    ]),
                    "utf-8",
                );

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "wrong-rels-default.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "strict" });
                expect(
                    validation.issues.filter(
                        (issue) =>
                            (issue.code ?? "").startsWith("ct-") &&
                            canonicalPartName(issue.path ?? "") === canonicalPartName("word/_rels/document.xml.rels"),
                    ),
                ).toEqual([]);
            });
        },
    );

    it("types the comments parts in the content-types namespace, whatever prefix the stream binds (#923)", { timeout: 30000 }, async () => {
        await withTempDir(async (dir) => {
            // The stream binds its namespace to `ct:`; before, the four
            // Overrides landed in no namespace, where the read model
            // cannot see them, so the comments parts resolved to the `xml`
            // Default's generic type.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                prefixedContentTypes(RELATIONSHIPS_TYPE, [
                    `<ct:Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                ]),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // Each comment part is typed by exactly one Override, in the
            // content-types namespace under the root's prefix, so the
            // read model resolves each part to its own type.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            const table = parseContentTypes(dom);
            for (const [partName, contentType] of COMMENT_PART_TYPES) {
                expect(namespacedOverridesNaming(dom, partName)).toEqual([[CONTENT_TYPES_NAMESPACE, "ct", partName, contentType]]);
                expect(resolveContentType(table, partName)).toEqual({ contentType, source: "override" });
            }

            // The typed, packed tree draws none of the findings the
            // un-namespaced Overrides drew: no schema error on the
            // stream itself, no generic typing on the comments parts.
            await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "prefixed-ct.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "strict" });
            expect(
                validation.issues.filter(
                    (issue) => (issue.code === "xsd-error" || issue.code === "xsd-summary") && issue.path === "[Content_Types].xml",
                ),
            ).toEqual([]);
            expect(validation.issues.filter((issue) => issue.code === "word-part-content-type-generic")).toEqual([]);
        });
    });

    it("corrects a prefixed Override of another type naming the relationships part it creates (#923)", async () => {
        await withTempDir(async (dir) => {
            // No main-document relationships part, and the prefixed Override
            // naming the part addComment will create carries the wrong type:
            // #901's rewrite must reach it through the prefix.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                prefixedContentTypes(RELATIONSHIPS_TYPE, [
                    `<ct:Override PartName="/word/_rels/document.xml.rels" ContentType="application/xml"/>`,
                    `<ct:Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                ]),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The Override is rewritten in place, not answered with a second
            // one, and the read model resolves the created part through it.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(resolveContentType(parseContentTypes(dom), "/word/_rels/document.xml.rels")).toEqual({
                contentType: RELATIONSHIPS_TYPE,
                source: "override",
            });
            expect(namespacedOverridesNaming(dom, "/word/_rels/document.xml.rels")).toEqual([
                [CONTENT_TYPES_NAMESPACE, "ct", "/word/_rels/document.xml.rels", RELATIONSHIPS_TYPE],
            ]);
        });
    });

    it("adds no Override when a prefixed Override already types word/comments.xml (#923)", async () => {
        await withTempDir(async (dir) => {
            const commentsType = COMMENT_PART_TYPES[0]![1];
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                prefixedContentTypes(RELATIONSHIPS_TYPE, [
                    `<ct:Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    `<ct:Override PartName="/word/comments.xml" ContentType="${commentsType}"/>`,
                ]),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The prefixed Override is read as typing the part, so the stream
            // keeps exactly one Override naming each part it named: none is
            // added, none rewritten, and the created relationships part is
            // covered by the Default alone.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(namespacedOverridesNaming(dom, "/word/comments.xml")).toEqual([
                [CONTENT_TYPES_NAMESPACE, "ct", "/word/comments.xml", commentsType],
            ]);
            expect(namespacedOverridesNaming(dom, "/word/document.xml")).toEqual([
                [CONTENT_TYPES_NAMESPACE, "ct", "/word/document.xml", MAIN_DOCUMENT_CONTENT_TYPE],
            ]);
            // Counted across every namespace, so an Override added in no
            // namespace would show up here too.
            expect(getElementsByTagNameNSAll(dom, "*", "Override")).toHaveLength(2);
        });
    });

    it("does not read an Override in another namespace as typing the comments parts (#923)", async () => {
        await withTempDir(async (dir) => {
            // An unprefixed stream whose Override naming word/comments.xml
            // binds another namespace types nothing: the four Overrides are
            // added in the content-types namespace beside it, serialized
            // exactly as an unprefixed stream is today — no new declaration.
            const commentsType = COMMENT_PART_TYPES[0]![1];
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    `<x:Override xmlns:x="urn:example:other" PartName="/word/comments.xml" ContentType="${commentsType}"/>`,
                ]),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            const ctAfter = await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8");
            expect(ctAfter).toContain(`<Override PartName="/word/comments.xml" ContentType="${commentsType}"/>`);
            const dom = parseXml(ctAfter);
            for (const [partName, contentType] of COMMENT_PART_TYPES) {
                expect(namespacedOverridesNaming(dom, partName)).toEqual([[CONTENT_TYPES_NAMESPACE, "", partName, contentType]]);
            }
            expect(resolveContentType(parseContentTypes(dom), "/word/comments.xml")).toEqual({
                contentType: commentsType,
                source: "override",
            });
        });
    });

    it("relates the comments parts through a renamed main document's own relationships part (#906)", { timeout: 30000 }, async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The relationships live in the sidecar of the main document that
            // exists, not the conventional one that does not.
            const relsPath = path.join(dir, "word", "_rels", "main.xml.rels");
            const relationships = parseXml(await fs.readFile(relsPath, "utf-8")).getElementsByTagName("Relationship");
            const targets: string[] = [];
            for (let i = 0; i < relationships.length; i += 1) {
                targets.push(relationships.item(i)!.getAttribute("Target") ?? "");
            }
            expect(targets).toEqual(["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"]);
            await expect(fs.access(path.join(dir, "word", "_rels", "document.xml.rels"))).rejects.toThrow();

            // A second comment adds no second set of relationships.
            const second = await addComment({
                unpackedDir: dir,
                commentId: 1,
                text: "Reply text",
                author: "Reviewer",
                initials: "R",
                date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
            });
            expect(second.message).toMatch(/^Added comment 1 \(para_id=[0-9A-F]{8}\)$/);
            expect(parseXml(await fs.readFile(relsPath, "utf-8")).getElementsByTagName("Relationship").length).toBe(4);

            // The marked, packed tree validates: the real main document reads
            // its comments through its own relationships part.
            await fs.writeFile(path.join(dir, "word", "main.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "renamed-main.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "word-valid" });
            expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
            expect(validation.valid).toBe(true);
        });
    });

    it(
        "relates the comments parts from a main document outside word/, with folder-relative targets (#906)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("main/doc.xml"), "main/doc.xml", { relsDefault: false });

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                const relationships = parseXml(
                    await fs.readFile(path.join(dir, "main", "_rels", "doc.xml.rels"), "utf-8"),
                ).getElementsByTagName("Relationship");
                const targets: string[] = [];
                for (let i = 0; i < relationships.length; i += 1) {
                    targets.push(relationships.item(i)!.getAttribute("Target") ?? "");
                }
                expect(targets).toEqual([
                    "../word/comments.xml",
                    "../word/commentsExtended.xml",
                    "../word/commentsIds.xml",
                    "../word/commentsExtensible.xml",
                ]);

                // The relationships part a first comment creates outside the
                // conventional place is typed under its own name (#895).
                const ctAfter = await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8");
                expect(ctAfter).toContain(`<Override PartName="/main/_rels/doc.xml.rels" ContentType="${RELATIONSHIPS_TYPE}"/>`);

                await fs.writeFile(path.join(dir, "main", "doc.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "moved-main.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code?.startsWith("comment-"))).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it("keeps writing word/_rels/document.xml.rels when the package relationships name word/document.xml (#906)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"));

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            const relsXml = await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8");
            expect(relsXml).toContain('Target="comments.xml"');
            await expect(fs.access(path.join(dir, "word", "_rels", "main.xml.rels"))).rejects.toThrow();
        });
    });

    it.each(["/word/document.xml", "Word/Document.xml", "word/d%6Fcument.xml"])(
        "keeps the conventional wiring when the package relationships name %s, an equivalent of word/document.xml (#906)",
        async (target) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(target), "word/document.xml");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // OPC part names compare ASCII case-insensitively, with
                // unreserved escapes decoded, so the target names the
                // conventional main document.
                expect(await fs.readdir(path.join(dir, "word", "_rels"))).toEqual(["document.xml.rels"]);
                const relationships = parseXml(
                    await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8"),
                ).getElementsByTagName("Relationship");
                const targets: string[] = [];
                for (let i = 0; i < relationships.length; i += 1) {
                    targets.push(relationships.item(i)!.getAttribute("Target") ?? "");
                }
                expect(targets).toEqual(["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"]);
            });
        },
    );

    it("accepts the ISO Strict officeDocument relationship type when finding the main document (#906)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(
                dir,
                packageRels(`Type="${STRICT_OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="word/main.xml"`),
                "word/main.xml",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            expect(await fs.readFile(path.join(dir, "word", "_rels", "main.xml.rels"), "utf-8")).toContain('Target="comments.xml"');
        });
    });

    it.each(["/word/comments.xml", "./comments.xml", "Comments.xml", "comm%65nts.xml"])(
        "adds no second set of relationships when the main document already relates word/comments.xml as %s (#906)",
        async (target) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(
                    path.join(dir, "word", "_rels", "main.xml.rels"),
                    RELS_XML.replace("relationships/styles", "relationships/comments").replace('Target="styles.xml"', `Target="${target}"`),
                    "utf-8",
                );

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The resolved target is word/comments.xml under OPC
                // part-name equivalence, so nothing is appended.
                const relationships = parseXml(
                    await fs.readFile(path.join(dir, "word", "_rels", "main.xml.rels"), "utf-8"),
                ).getElementsByTagName("Relationship");
                expect(relationships.length).toBe(1);
                expect(relationships.item(0)!.getAttribute("Target")).toBe(target);
            });
        },
    );

    it.each([
        [
            "an External comments relationship (word/document.xml)",
            "word/document.xml",
            packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="/word/comments.xml" TargetMode="External"`),
            ["rId1", COMMENTS_RELATIONSHIP_TYPE, "/word/comments.xml", "External"],
        ],
        [
            "a hyperlink relationship (word/document.xml)",
            "word/document.xml",
            packageRels(`Type="${HYPERLINK_RELATIONSHIP_TYPE}" Target="comments.xml"`),
            ["rId1", HYPERLINK_RELATIONSHIP_TYPE, "comments.xml", ""],
        ],
        [
            "an External comments relationship (word/main.xml)",
            "word/main.xml",
            packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="/word/comments.xml" TargetMode="External"`),
            ["rId1", COMMENTS_RELATIONSHIP_TYPE, "/word/comments.xml", "External"],
        ],
        [
            "a hyperlink relationship (word/main.xml)",
            "word/main.xml",
            packageRels(`Type="${HYPERLINK_RELATIONSHIP_TYPE}" Target="comments.xml"`),
            ["rId1", HYPERLINK_RELATIONSHIP_TYPE, "comments.xml", ""],
        ],
    ])(
        "relates the comments parts when the main document's relationships part already holds %s (#917)",
        async (_label, mainPart, relsXml, existing) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(mainPart), mainPart);
                const relsPath = path.join(dir, "word", "_rels", `${path.posix.basename(mainPart)}.rels`);
                await fs.mkdir(path.dirname(relsPath), { recursive: true });
                await fs.writeFile(relsPath, relsXml, "utf-8");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // Neither relationship relates the comments part — not by
                // type, or not internally — so the four relationships are
                // added beside it, each with an id no other relationship in
                // the part uses, and the existing one is left as it was.
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                    existing,
                    ...firstCommentRelationships(2),
                ]);
            });
        },
    );

    it("validates once the markers are added over an External comments relationship (#917)", { timeout: 30000 }, async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(dir, "word", "_rels", "document.xml.rels"),
                packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="/word/comments.xml" TargetMode="External"`),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The comment is related through the relationship the first
            // comment added, so the markers name a comment Word reads.
            await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "external-comments-rel.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "word-valid" });
            expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
            expect(validation.valid).toBe(true);
        });
    });

    it("adds no second set of relationships when the main document already relates word/comments.xml with the Strict comments type (#917)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(dir, "word", "_rels", "main.xml.rels"),
                packageRels(`Type="${STRICT_COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"`),
                "utf-8",
            );

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The Strict comments type relates the part as the Transitional
            // one does, so nothing is appended.
            expect(relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", "main.xml.rels"), "utf-8")))).toEqual([
                ["rId1", STRICT_COMMENTS_RELATIONSHIP_TYPE, "comments.xml", ""],
            ]);
        });
    });

    it.each([
        [
            "the package relationships declare an encoding it cannot decode",
            officeDocumentRels("word/main.xml"),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} is stored in an encoding this tool cannot decode`,
            async (dir: string) => declareUndecodable(path.join(dir, "_rels", ".rels")),
        ],
        [
            "the package relationships are not well-formed XML",
            '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="' +
                OFFICE_DOCUMENT_RELATIONSHIP_TYPE +
                '" Target="word/main.xml">\n</Relationships>',
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} is not well-formed XML`,
            undefined,
        ],
        [
            "the package relationships name no officeDocument relationship",
            packageRels(`Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE.replace(/officeDocument$/, "styles")}" Target="word/main.xml"`),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} names no officeDocument relationship`,
            undefined,
        ],
        [
            "the officeDocument relationship is External",
            packageRels(`Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="https://example.com/doc.xml" TargetMode="External"`),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} names an External officeDocument relationship`,
            undefined,
        ],
        [
            "the officeDocument target escapes the package root",
            officeDocumentRels("../outside/main.xml"),
            undefined,
            (dir: string) => `${path.join(dir, "_rels", ".rels")} names a main document outside the package`,
            undefined,
        ],
        [
            "the named main document does not exist",
            officeDocumentRels("word/main.xml"),
            undefined,
            (dir: string) => `${path.join(dir, "word", "main.xml")} not found`,
            undefined,
        ],
    ])("refuses a first comment when %s, changing nothing (#906)", async (_label, relsXml, mainPart, expected, prepare) => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, relsXml, mainPart);
            await prepare?.(dir);

            const before = await readTree(dir);
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });

            expect(result.paraId).toBe("");
            expect(result.message).toMatch(/^Error: /);
            expect(result.message).toContain(expected(dir));
            expect(await readTree(dir)).toEqual(before);
        });
    });

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

    it.each([
        [
            "[Content_Types].xml is not well-formed XML",
            async (dir: string) => {
                await makeFixture(dir);
                await fs.writeFile(path.join(dir, "[Content_Types].xml"), malformed(CT_XML), "utf-8");
            },
            (dir: string) => path.join(dir, "[Content_Types].xml"),
        ],
        [
            "word/_rels/document.xml.rels is not well-formed XML",
            async (dir: string) => {
                await makeFixture(dir);
                await fs.writeFile(path.join(dir, "word", "_rels", "document.xml.rels"), malformed(RELS_XML), "utf-8");
            },
            (dir: string) => path.join(dir, "word", "_rels", "document.xml.rels"),
        ],
        [
            "a renamed main document's relationships part is not well-formed XML",
            async (dir: string) => {
                await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(path.join(dir, "word", "_rels", "main.xml.rels"), malformed(RELS_XML), "utf-8");
            },
            (dir: string) => path.join(dir, "word", "_rels", "main.xml.rels"),
        ],
    ])("refuses a first comment when %s, changing nothing (#905)", async (_label, prepare, partPath) => {
        await withTempDir(async (dir) => {
            await prepare(dir);

            const before = await readTree(dir);
            // Before #905 the call threw mid-write, after earlier parts changed.
            const call = addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            expect(result.message).toMatch(/^Error: /);
            expect(result.message).toContain(`${partPath(dir)} is not well-formed XML`);
            // The parser's own message follows the part name (#905).
            expect(result.message).not.toBe(`Error: ${partPath(dir)} is not well-formed XML`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it.each(["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"])(
        "refuses a later comment when word/%s is not well-formed XML, changing nothing (#905)",
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
                await fs.writeFile(part, malformed(await fs.readFile(part, "utf-8")), "utf-8");
                const before = await readTree(dir);
                // Before #905 the call threw mid-write, after earlier parts changed.
                const call = addComment({
                    unpackedDir: dir,
                    commentId: 1,
                    text: "second",
                    author: "Test",
                    initials: "T",
                });
                await expect(call).resolves.toMatchObject({ paraId: "" });
                const result = await call;

                expect(result.message).toMatch(/^Error: /);
                expect(result.message).toContain(`${part} is not well-formed XML`);
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );
});
