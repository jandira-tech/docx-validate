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
const STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const MARKUP_COMPATIBILITY_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const UNKNOWN_NAMESPACE = "urn:unknown";
const COMMENTS_EXTENDED_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_IDS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
const WORD_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const W16CEX_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml/cex";

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

/** The package relationships part with its namespace bound to the `p:` prefix, holding one relationship with `attributes` (#918). */
const prefixedPackageRels = (attributes: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:Relationships xmlns:p="${PACKAGE_RELATIONSHIPS_NAMESPACE}">\n  <p:Relationship Id="rId1" ${attributes}/>\n</p:Relationships>`;

/** A relationships stream whose root also binds `mc`, `u` and `w14`, holding `body` between its root tags (#932). */
const mcRelationships = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" xmlns:u="${UNKNOWN_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}">${body}</Relationships>`;

/** An `mc:AlternateContent` whose Fallback is selected: its Choice requires `u`, a namespace nothing understands (#932). */
const fallbackBranch = (content: string): string =>
    `<mc:AlternateContent><mc:Choice Requires="u"/><mc:Fallback>${content}</mc:Fallback></mc:AlternateContent>`;

/** An `mc:AlternateContent` whose Choice is selected: it requires `w14`, a namespace the checks understand (#932). */
const understoodChoiceBranch = (content: string): string =>
    `<mc:AlternateContent><mc:Choice Requires="w14">${content}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;

/** An `mc:AlternateContent` whose Choice is not selected: it requires `u`, a namespace nothing understands, and its Fallback is empty (#927). */
const unselectedChoiceBranch = (content: string): string =>
    `<mc:AlternateContent><mc:Choice Requires="u">${content}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;

/** The package relationships part whose only relationship, with `attributes`, sits inside the branch `branch` wraps it in (#932). */
const branchedPackageRels = (branch: (content: string) => string, attributes: string): string =>
    mcRelationships(branch(`<Relationship Id="rId1" ${attributes}/>`));

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

/** A content-types stream whose root binds the namespace both to the `ct:` prefix and as the default, with a `.rels` Default naming `relsType` and `overrides` spliced in (#937). */
const dualPrefixedContentTypes = (relsType: string, overrides: readonly string[] = []): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ct:Types xmlns:ct="${CONTENT_TYPES_NAMESPACE}" xmlns="${CONTENT_TYPES_NAMESPACE}"><ct:Default Extension="rels" ContentType="${relsType}"/>${overrides.join("")}</ct:Types>`;

/** The four comments parts a first comment types, each with the type it gets (#923). */
const COMMENT_PART_TYPES: readonly [string, string][] = [
    ["/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"],
    ["/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"],
    ["/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"],
    ["/word/commentsExtensible.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"],
];

/** The four comment-part Overrides, the comments part spelled `commentsSpelling`, each element's name carrying `prefix` (#937). */
const commentOverrides = (prefix: string, commentsSpelling: string): string[] =>
    COMMENT_PART_TYPES.map(
        ([partName, contentType], index) =>
            `<${prefix}Override PartName="${index === 0 ? commentsSpelling : partName}" ContentType="${contentType}"/>`,
    );

/** The minimal main document with a paired marker triple for comment 0. */
const markedMainDocument = (): string =>
    MINIMAL_PARTS["word/document.xml"]!.replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>',
    );

/** The minimal main document with a paired marker triple for each of comments 0 and 1. */
const twiceMarkedMainDocument = (): string =>
    MINIMAL_PARTS["word/document.xml"]!.replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>Hello</w:t></w:r>' +
            '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>' +
            '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>',
    );

/**
 * A part's XML with the namespace `from` prefixes rebound to `to`: the
 * declaration, the element prefixes and the prefixed attributes alike, so
 * the part keeps its content and names it under another prefix (#930).
 */
function rebindPrefix(text: string, from: string, to: string): string {
    return text
        .replaceAll(`xmlns:${from}=`, `xmlns:${to}=`)
        .replace(new RegExp(`<(\\/?)${from}:`, "g"), `<$1${to}:`)
        .replace(new RegExp(` ${from}:([A-Za-z]+)=`, "g"), ` ${to}:$1=`);
}

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

/**
 * Every file under `dir`, keyed by its relative path, with its bytes. A
 * symbolic link is recorded by the path it names, not read through, so the
 * snapshot never reads a file outside the tree (#939).
 */
async function readTree(dir: string): Promise<Map<string, Buffer>> {
    const files = new Map<string, Buffer>();
    const walk = async (current: string): Promise<void> => {
        for (const entry of await fs.readdir(current, { withFileTypes: true })) {
            const full = path.join(current, entry.name);
            if (entry.isSymbolicLink()) {
                files.set(path.relative(dir, full), Buffer.from(await fs.readlink(full), "utf-8"));
            } else if (entry.isDirectory()) {
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

/** Every [localName, prefix, namespaceURI, Id, Type, Target] row of a relationships root's element children, whatever their namespace (#918). */
function relationshipElementRows(dom: ReturnType<typeof parseXml>): string[][] {
    const rows: string[][] = [];
    for (let child = dom.documentElement.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) {
            continue;
        }
        const rel = child as Element;
        rows.push([
            rel.localName ?? "",
            rel.prefix ?? "",
            rel.namespaceURI ?? "",
            rel.getAttribute("Id") ?? "",
            rel.getAttribute("Type") ?? "",
            rel.getAttribute("Target") ?? "",
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

    it("adds no second Override when a prefixed Override already types word/comments.xml (#923)", async () => {
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

            // The prefixed Override is read as typing the part, so it is not
            // answered with a second Override and is not rewritten, and the
            // created relationships part is covered by the Default alone.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(namespacedOverridesNaming(dom, "/word/comments.xml")).toEqual([
                [CONTENT_TYPES_NAMESPACE, "ct", "/word/comments.xml", commentsType],
            ]);
            expect(namespacedOverridesNaming(dom, "/word/document.xml")).toEqual([
                [CONTENT_TYPES_NAMESPACE, "ct", "/word/document.xml", MAIN_DOCUMENT_CONTENT_TYPE],
            ]);
            // The other three comment parts are untyped, so each gains its
            // one Override beside it, also under the root's prefix (#937).
            for (const [partName, contentType] of COMMENT_PART_TYPES.slice(1)) {
                expect(namespacedOverridesNaming(dom, partName)).toEqual([[CONTENT_TYPES_NAMESPACE, "ct", partName, contentType]]);
            }
            // Counted across every namespace, so an Override added in no
            // namespace would show up here too.
            expect(getElementsByTagNameNSAll(dom, "*", "Override")).toHaveLength(5);
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

    it.each([
        ["an unprefixed stream", "", contentTypes],
        ["a `ct:`-prefixed stream", "ct:", prefixedContentTypes],
    ] as const)(
        "declares each comment part once when %s already types all four, the comments part under another spelling (#937)",
        { timeout: 30000 },
        async (_streamKind, prefix, buildStream) => {
            await withTempDir(async (dir) => {
                // The stream types all four comment parts, but the comments
                // part's Override is spelled /WORD/comments.xml: under OPC
                // equivalence that names the same part, so nothing may be
                // added — before, the literal comparison missed it and all
                // four parts were declared a second time.
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    buildStream(RELATIONSHIPS_TYPE, [
                        `<${prefix}Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                        ...commentOverrides(prefix, "/WORD/comments.xml"),
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

                // Counted by canonical name across every namespace, each
                // comment part is declared exactly once.
                const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
                const counts = new Map<string, number>();
                for (const override of getElementsByTagNameNSAll(dom, "*", "Override")) {
                    const key = canonicalPartName(override.getAttribute("PartName") ?? "");
                    counts.set(key, (counts.get(key) ?? 0) + 1);
                }
                for (const [partName] of COMMENT_PART_TYPES) {
                    expect(counts.get(canonicalPartName(partName))).toBe(1);
                }

                // The marked, packed tree validates: no part is declared
                // twice.
                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "respelled-comments.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "ct-duplicate-declaration")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it("reads an escaped spelling of the comments part as naming it, adding nothing for it (#937)", async () => {
        await withTempDir(async (dir) => {
            // %65 is `e`, so /word/comm%65nts.xml names word/comments.xml
            // under OPC equivalence: the Override already types the part and
            // keeps its place, its spelling and its type.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    ...commentOverrides("", "/word/comm%65nts.xml"),
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

            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            for (const [partName, contentType] of COMMENT_PART_TYPES) {
                const spelling = partName === "/word/comments.xml" ? "/word/comm%65nts.xml" : partName;
                expect(overridesNaming(dom, partName)).toEqual([[spelling, contentType]]);
            }
        });
    });

    it("decides each comment part on its own, typing only the ones no Override names (#937)", async () => {
        await withTempDir(async (dir) => {
            // The comments part is untyped but commentsIds.xml is not: before,
            // one gate on the comments part added all four Overrides and
            // declared commentsIds.xml twice.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    `<Override PartName="/word/commentsIds.xml" ContentType="${COMMENT_PART_TYPES[2]![1]}"/>`,
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

            // The three no Override names are added; commentsIds.xml stays
            // declared exactly once, by its own Override.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            for (const [partName, contentType] of COMMENT_PART_TYPES) {
                expect(overridesNaming(dom, partName)).toEqual([[partName, contentType]]);
            }
        });
    });

    it("leaves the stream byte-identical when all four comment parts are typed with their exact spellings (#937)", async () => {
        await withTempDir(async (dir) => {
            // A control: nothing is added and no #901 rewrite applies, so
            // the stream is not written at all.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    ...commentOverrides("", "/word/comments.xml"),
                ]),
                "utf-8",
            );
            const before = await fs.readFile(path.join(dir, "[Content_Types].xml"));

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            expect(await fs.readFile(path.join(dir, "[Content_Types].xml"))).toEqual(before);
        });
    });

    it("leaves a single-quoted stream byte-identical too, when every comment part is typed (#937)", async () => {
        await withTempDir(async (dir) => {
            // The serializer normalizes single-quoted attributes to double
            // quotes, so a needless write would show: with nothing to add
            // and no #901 rewrite, the stream is not written at all.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            const overrides = commentOverrides("", "/word/comments.xml").map((row) => row.replaceAll('"', "'"));
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns='${CONTENT_TYPES_NAMESPACE}'><Default Extension='rels' ContentType='${RELATIONSHIPS_TYPE}'/><Override PartName='/word/document.xml' ContentType='${MAIN_DOCUMENT_CONTENT_TYPE}'/>${overrides.join("")}</Types>`,
                "utf-8",
            );
            const before = await fs.readFile(path.join(dir, "[Content_Types].xml"));

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            expect(await fs.readFile(path.join(dir, "[Content_Types].xml"))).toEqual(before);
        });
    });

    it("carries the root's prefix on new Overrides when the root also binds the namespace as the default (#937)", async () => {
        await withTempDir(async (dir) => {
            // A root binding the content-types namespace both to `ct:` and
            // as the default: the root's own prefix is `ct`, so the new
            // Overrides carry it — an unnamespaced one would serialize
            // under the default binding instead.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                dualPrefixedContentTypes(RELATIONSHIPS_TYPE, [
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

            // Each new Override is in the content-types namespace under the
            // root's own prefix, read back after a reparse.
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            for (const [partName, contentType] of COMMENT_PART_TYPES) {
                expect(namespacedOverridesNaming(dom, partName)).toEqual([[CONTENT_TYPES_NAMESPACE, "ct", partName, contentType]]);
            }
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
            // The named conventional main document is stored; before #919 a
            // missing one was still written to instead of refused.
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");

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

    it.each([
        ["word/main.xml?x=1", "word/main.xml", "main.xml.rels"],
        ["word/main.xml#anchor", "word/main.xml", "main.xml.rels"],
        ["word/document.xml?x=1", "word/document.xml", "document.xml.rels"],
        ["word/m%61in.xml", "word/main.xml", "main.xml.rels"],
    ])(
        "relates the comments parts through the stored main document when the package relationships name %s (#919)",
        async (target, mainPart, relsName) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(target), mainPart);

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The target is resolved as the read model resolves it — the
                // query or fragment dropped, the escaped spelling decoded —
                // and the comments are wired through the stored part's own
                // name, never through the reference as written.
                const relsPath = path.join(dir, "word", "_rels", relsName);
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual(firstCommentRelationships(1));
                const other = relsName === "document.xml.rels" ? "main.xml.rels" : "document.xml.rels";
                await expect(fs.access(path.join(dir, "word", "_rels", other))).rejects.toThrow();
            });
        },
    );

    it("validates once the markers are added over a main document named with a query (#919)", { timeout: 30000 }, async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml?x=1"), "word/main.xml");

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The read model drops the query too, so Word reads the main
            // document the relationship names and the comment the markers
            // reference through that document's own relationships part.
            await fs.writeFile(path.join(dir, "word", "main.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "query-main.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "word-valid" });
            expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
            expect(validation.valid).toBe(true);
        });
    });

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
        ["word/document.xml", "document.xml.rels"],
        ["word/main.xml", "main.xml.rels"],
    ])("relates the comments parts through a _rels/.rels bound to a prefix naming %s (#918)", async (mainPart, relsName) => {
        await withTempDir(async (dir) => {
            await makePackageFixture(
                dir,
                prefixedPackageRels(`Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="${mainPart}"`),
                mainPart,
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

            // The prefixed p:Relationship names the main document, so the
            // comments are related from that document's own sidecar.
            const relsPath = path.join(dir, "word", "_rels", relsName);
            expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual(firstCommentRelationships(1));
        });
    });

    it(
        "writes the new relationships in the package-relationships namespace, under the root's prefix (#918)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
                await fs.writeFile(
                    relsPath,
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                        `<p:Relationships xmlns:p="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                        `<p:Relationship Id="rId1" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>` +
                        `</p:Relationships>`,
                    "utf-8",
                );
                // The styles relationship must name a stored, typed part for the
                // packed tree to validate, as the minimal package's does.
                await fs.writeFile(path.join(dir, "word", "styles.xml"), MINIMAL_PARTS["word/styles.xml"]!, "utf-8");
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    contentTypes(RELATIONSHIPS_TYPE, [
                        `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                        `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
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

                // Five relationships: the prefixed one that was there, unchanged,
                // and the four a first comment appends beside it — every one in
                // the package-relationships namespace under the root's `p:`
                // prefix, with ids rId1–rId5, all distinct.
                const rows = relationshipElementRows(parseXml(await fs.readFile(relsPath, "utf-8")));
                expect(rows).toEqual([
                    ["Relationship", "p", PACKAGE_RELATIONSHIPS_NAMESPACE, "rId1", STYLES_RELATIONSHIP_TYPE, "styles.xml"],
                    ...firstCommentRelationships(2).map(([id, type, target]) => [
                        "Relationship",
                        "p",
                        PACKAGE_RELATIONSHIPS_NAMESPACE,
                        id,
                        type,
                        target,
                    ]),
                ]);
                expect(new Set(rows.map((row) => row[3])).size).toBe(5);

                // The marked, packed tree validates: the relationships Word reads
                // are the ones the root's prefix binds, duplicated ids and all.
                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "prefixed-rels.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "opc-relationships-structure")).toEqual([]);
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it("adds no second set of relationships when a prefixed relationships part already relates word/comments.xml (#918)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            const relsPath = path.join(dir, "word", "_rels", "main.xml.rels");
            await fs.writeFile(
                relsPath,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<p:Relationships xmlns:p="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                    `<p:Relationship Id="rId1" Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"/>` +
                    `</p:Relationships>`,
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

            // The prefixed p:Relationship relates the comments part under OPC
            // part-name equivalence, so nothing is appended and it is left
            // as it was, prefix and all.
            expect(relationshipElementRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                ["Relationship", "p", PACKAGE_RELATIONSHIPS_NAMESPACE, "rId1", COMMENTS_RELATIONSHIP_TYPE, "comments.xml"],
            ]);
        });
    });

    it("writes the new relationships under the root's prefix when it binds the namespace both ways (#918)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            const relsPath = path.join(dir, "word", "_rels", "main.xml.rels");
            await fs.writeFile(
                relsPath,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<p:Relationships xmlns:p="${PACKAGE_RELATIONSHIPS_NAMESPACE}" xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                    `<p:Relationship Id="rId1" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>` +
                    `</p:Relationships>`,
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

            // The root has a prefix, so the new relationships carry it — not
            // the default binding the root also declares.
            expect(relationshipElementRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                ["Relationship", "p", PACKAGE_RELATIONSHIPS_NAMESPACE, "rId1", STYLES_RELATIONSHIP_TYPE, "styles.xml"],
                ...firstCommentRelationships(2).map(([id, type, target]) => [
                    "Relationship",
                    "p",
                    PACKAGE_RELATIONSHIPS_NAMESPACE,
                    id,
                    type,
                    target,
                ]),
            ]);
        });
    });

    it("numbers new ids past every Relationship in the part, in any namespace, and reads only namespaced relationships (#918)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
            const relsXml =
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                `<Relationship Id="rId1" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>` +
                `<x:Relationship xmlns:x="urn:example:other" Id="rId7" Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"/>` +
                `</Relationships>`;
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

            // The comments-typed relationship in another namespace does not
            // relate the comments part, so the four relationships are added;
            // and their ids start past the rId7 it holds, whatever its
            // namespace, so no id the part already uses is repeated.
            expect(relationshipElementRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                ["Relationship", "", PACKAGE_RELATIONSHIPS_NAMESPACE, "rId1", STYLES_RELATIONSHIP_TYPE, "styles.xml"],
                ["Relationship", "x", "urn:example:other", "rId7", COMMENTS_RELATIONSHIP_TYPE, "comments.xml"],
                ...firstCommentRelationships(8).map(([id, type, target]) => [
                    "Relationship",
                    "",
                    PACKAGE_RELATIONSHIPS_NAMESPACE,
                    id,
                    type,
                    target,
                ]),
            ]);

            // The appended relationships serialize exactly as the unprefixed
            // part always has: no xmlns declaration of their own.
            const appended = firstCommentRelationships(8)
                .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
                .join("");
            expect(await fs.readFile(relsPath, "utf-8")).toBe(relsXml.replace("</Relationships>", `${appended}</Relationships>`));
        });
    });

    it.each([
        ["rId9007199254740992", ["rId9007199254740993", "rId9007199254740994", "rId9007199254740995", "rId9007199254740996"]],
        [
            "rId99999999999999999999",
            ["rId100000000000000000000", "rId100000000000000000001", "rId100000000000000000002", "rId100000000000000000003"],
        ],
        [
            "rId123456789012345678901234567890",
            [
                "rId123456789012345678901234567891",
                "rId123456789012345678901234567892",
                "rId123456789012345678901234567893",
                "rId123456789012345678901234567894",
            ],
        ],
    ])("numbers new ids one by one past a held %s, and the package validates (#940)", { timeout: 30000 }, async (heldId, newIds) => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
            await fs.writeFile(
                relsPath,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                    `<Relationship Id="${heldId}" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>` +
                    `</Relationships>`,
                "utf-8",
            );
            await fs.writeFile(path.join(dir, "word", "styles.xml"), MINIMAL_PARTS["word/styles.xml"]!, "utf-8");
            await fs.writeFile(
                path.join(dir, "[Content_Types].xml"),
                contentTypes(RELATIONSHIPS_TYPE, [
                    `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
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

            // The suffix is counted exactly, however many digits it runs to, so
            // each new id is one past the one before it and none repeats an id
            // the part holds.
            expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                [heldId, STYLES_RELATIONSHIP_TYPE, "styles.xml", ""],
                ...firstCommentRelationships(1).map(([, type, target, mode], i) => [newIds[i]!, type, target, mode]),
            ]);

            await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "long-rid.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "word-valid" });
            expect(validation.issues.filter((issue) => issue.code === "rels-id-duplicate")).toEqual([]);
            expect(validation.valid).toBe(true);
        });
    });

    it.each([
        ["rIdX", 1],
        ["R1", 1],
        ["rId007", 8],
        ["rId12x", 13],
        ["rId+5", 1],
    ])("counts only the digits right after rId in a held %s (#940)", async (heldId, firstRid) => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
            await fs.writeFile(
                relsPath,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                    `<Relationship Id="${heldId}" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>` +
                    `</Relationships>`,
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

            // An id that is not `rId` followed by digits is not counted, and an
            // `rId` suffix counts by its leading digits: `rId007` as 7 and
            // `rId12x` as 12. `rId+5`, which `Number.parseInt` read as 5, is
            // not an id a new `rId<N>` can repeat, and no longer counts.
            expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                [heldId, STYLES_RELATIONSHIP_TYPE, "styles.xml", ""],
                ...firstCommentRelationships(firstRid),
            ]);
        });
    });

    it("refuses a first comment when _rels/.rels's only officeDocument-typed element is in another namespace, changing nothing (#918)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(
                dir,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">` +
                    `<x:Relationship xmlns:x="urn:example:other" Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="word/document.xml"/>` +
                    `</Relationships>`,
            );

            const before = await readTree(dir);
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });

            expect(result.paraId).toBe("");
            expect(result.message).toBe(`Error: ${path.join(dir, "_rels", ".rels")} names no officeDocument relationship`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it.each([
        ["a selected Fallback", fallbackBranch],
        ["an understood Choice", understoodChoiceBranch],
    ])(
        "relates a first comment through an officeDocument relationship %s of _rels/.rels holds (#932)",
        { timeout: 30000 },
        async (_label, branch) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(
                    dir,
                    branchedPackageRels(branch, `Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="word/document.xml"`),
                    "word/document.xml",
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

                // The checks follow a relationship a selected branch holds as
                // they follow a direct child, so the wiring resolves through it
                // and the four relationships land in the conventional sidecar.
                const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual(firstCommentRelationships(1));

                // The marked, packed tree validates: the branch-held relationship
                // is the one the checks read.
                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "branched-package-rels.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it("refuses a first comment when _rels/.rels's only officeDocument relationship sits in an unselected Choice, changing nothing (#932)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(
                dir,
                branchedPackageRels(unselectedChoiceBranch, `Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="word/document.xml"`),
            );

            const before = await readTree(dir);
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });

            // Nothing selects the Choice, so the checks read no officeDocument
            // relationship, and neither does addComment.
            expect(result.paraId).toBe("");
            expect(result.message).toBe(`Error: ${path.join(dir, "_rels", ".rels")} names no officeDocument relationship`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it.each([
        ["a selected Fallback", fallbackBranch],
        ["an understood Choice", understoodChoiceBranch],
    ])(
        "adds no second set of relationships when %s of the main document's relationships part already relates word/comments.xml (#932)",
        { timeout: 30000 },
        async (_label, branch) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
                await fs.mkdir(path.dirname(relsPath), { recursive: true });
                await fs.writeFile(
                    relsPath,
                    mcRelationships(branch(`<Relationship Id="rId1" Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"/>`)),
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

                // The branch-held relationship relates the comments part as a
                // direct child does, so nothing is appended: the part keeps
                // exactly the one comments relationship it stored.
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                    ["rId1", COMMENTS_RELATIONSHIP_TYPE, "comments.xml", ""],
                ]);

                // The marked, packed tree validates: the branch-held
                // relationship is the one the checks read the comments part
                // through.
                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "branched-sidecar-rels.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it(
        "adds the four relationships beside a comments relationship only an unselected Choice holds (#927)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const relsPath = path.join(dir, "word", "_rels", "document.xml.rels");
                await fs.mkdir(path.dirname(relsPath), { recursive: true });
                await fs.writeFile(
                    relsPath,
                    mcRelationships(
                        unselectedChoiceBranch(`<Relationship Id="rId1" Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"/>`),
                    ),
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

                // The checks read no relationship the unselected Choice holds, so
                // the comments part is related to nothing and the four relationships
                // are added on the stored root, beside the untouched branch: the
                // part holds two comments relationships, and the new ids skip the
                // rId1 the stored one uses, unselected as it is.
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                    ["rId1", COMMENTS_RELATIONSHIP_TYPE, "comments.xml", ""],
                    ...firstCommentRelationships(2),
                ]);

                // The marked, packed tree validates: the markers name a comment
                // the added relationship relates.
                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "unselected-choice-rels.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it.each([
        ["comments.xml?q=1", "word/document.xml"],
        ["comments.xml#frag", "word/document.xml"],
        ["..\\word\\comments.xml", "main/doc.xml"],
        ["../word/comments.xml", "main/doc.xml"],
    ])(
        "adds no second set of relationships when the main document already relates word/comments.xml as %s (#928)",
        async (target, mainPart) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(mainPart), mainPart);
                const relsPath = path.join(
                    dir,
                    ...path.posix.dirname(mainPart).split("/"),
                    "_rels",
                    `${path.posix.basename(mainPart)}.rels`,
                );
                await fs.mkdir(path.dirname(relsPath), { recursive: true });
                await fs.writeFile(relsPath, packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="${target}"`), "utf-8");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The target is resolved as the read model resolves it — the
                // query or fragment dropped, `\` read as `/` — so it names
                // word/comments.xml and the part keeps exactly the one
                // comments relationship it stored.
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                    ["rId1", COMMENTS_RELATIONSHIP_TYPE, target, ""],
                ]);
            });
        },
    );

    it("validates once the markers are added over a comments relationship named with a query (#928)", { timeout: 30000 }, async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(dir, "word", "_rels", "document.xml.rels"),
                packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml?q=1"`),
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

            // The read model drops the query too, so Word reads the comments
            // part through the stored relationship and the comment's markers
            // name a comment Word reads.
            await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
            const file = path.join(dir, "query-comments-rel.docx");
            await packTree(dir, file);
            const validation = await validate(file, { profile: "word-valid" });
            expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
            expect(validation.valid).toBe(true);
        });
    });

    it.each(["http://example.com/comments.xml", "../../word/comments.xml"])(
        "adds the four relationships beside a comments relationship whose target %s names no part (#928)",
        async (target) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(
                    path.join(dir, "word", "_rels", "document.xml.rels"),
                    packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="${target}"`),
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

                // A URI names no package part, whatever its TargetMode says,
                // and a target that climbs past the package root keeps its
                // `../`, as the read model leaves it unclamped: the stored
                // relationship relates nothing, so the four relationships are
                // added beside it.
                expect(
                    relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8"))),
                ).toEqual([["rId1", COMMENTS_RELATIONSHIP_TYPE, target, ""], ...firstCommentRelationships(2)]);
            });
        },
    );

    it.each(["/word/../word/main.xml", "/word/./main.xml"])(
        "adds a comment through word/_rels/main.xml.rels when the package relationships name %s (#947)",
        async (target) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(target), "word/main.xml");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The package-absolute target's dot segments resolve as a
                // relative target's do, so it names the stored word/main.xml
                // and the four relationships are wired through that part's
                // own relationships part.
                const relsPath = path.join(dir, "word", "_rels", "main.xml.rels");
                expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual(firstCommentRelationships(1));
            });
        },
    );

    it("still refuses a first comment when the package relationships name /../word/document.xml, changing nothing (#947)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("/../word/document.xml"), "word/main.xml");

            const before = await readTree(dir);
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });

            // The read model leaves an escaping target unclamped, so the
            // climb past the package root is kept and the target still names
            // no part, exactly as before #947.
            expect(result.paraId).toBe("");
            expect(result.message).toBe(`Error: ${path.join(dir, "_rels", ".rels")} names a main document outside the package`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it("adds no second set of relationships when the main document relates word/comments.xml as /word/../word/comments.xml (#947)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/main.xml"), "word/main.xml");
            const relsPath = path.join(dir, "word", "_rels", "main.xml.rels");
            await fs.mkdir(path.dirname(relsPath), { recursive: true });
            await fs.writeFile(relsPath, packageRels(`Type="${COMMENTS_RELATIONSHIP_TYPE}" Target="/word/../word/comments.xml"`), "utf-8");

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The package-absolute target's dot segments resolve, so the
            // stored relationship already names word/comments.xml and the
            // part keeps exactly that one comments relationship.
            expect(relationshipRows(parseXml(await fs.readFile(relsPath, "utf-8")))).toEqual([
                ["rId1", COMMENTS_RELATIONSHIP_TYPE, "/word/../word/comments.xml", ""],
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

    it.each([
        [
            "word/document.xml, which is not stored",
            officeDocumentRels("word/document.xml"),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "word", "document.xml")} not found`,
        ],
        [
            "an absolute URI",
            packageRels(`Type="${OFFICE_DOCUMENT_RELATIONSHIP_TYPE}" Target="http://example.com/document.xml"`),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} names a main document outside the package`,
        ],
        [
            "/../word/document.xml, which climbs past the package root",
            officeDocumentRels("/../word/document.xml"),
            "word/main.xml",
            (dir: string) => `${path.join(dir, "_rels", ".rels")} names a main document outside the package`,
        ],
    ])(
        "refuses a first comment when the package relationships name %s, naming no stored part, changing nothing (#919)",
        async (_label, relsXml, mainPart, expected) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, relsXml, mainPart);

                const before = await readTree(dir);
                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "x",
                    author: "TestAuthor",
                    initials: "TA",
                });

                // A named main document that is not stored is refused like
                // any other, before any write — the conventional name included.
                expect(result.paraId).toBe("");
                expect(result.message).toBe(`Error: ${expected(dir)}`);
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );

    it.each([
        ["word/Document.xml", "word/Document.xml?x=1", "Document.xml.rels"],
        ["word/docum%65nt.xml", "word/docum%65nt.xml", "docum%65nt.xml.rels"],
        ["word/main.xml", "word/main.xml", "m%61in.xml.rels"],
        ["word/document.xml", "word/document.xml", "docum%65nt.xml.rels"],
    ])(
        "writes a first comment's relationships through the stored sidecar beside a main document stored as %s (#946)",
        async (mainPart, relsTarget, relsName) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(relsTarget), mainPart);
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(path.join(dir, "word", "_rels", relsName), RELS_XML, "utf-8");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The relationships join the styles relationship in the
                // stored sidecar, found under OPC equivalence like the main
                // document itself, whatever spelling either has; before, a
                // differently spelled word/document.xml was wired through
                // word/_rels/document.xml.rels — a second, OPC-equivalent
                // part beside its own sidecar — and a sidecar stored under
                // another spelling got that conventional twin beside it.
                expect(await fs.readdir(path.join(dir, "word", "_rels"))).toEqual([relsName]);
                expect(relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", relsName), "utf-8")))).toEqual([
                    ["rId1", STYLES_RELATIONSHIP_TYPE, "styles.xml", ""],
                    ...firstCommentRelationships(2),
                ]);
            });
        },
    );

    it.each(["docum%65nt.xml.rels", "Document.xml.rels"])(
        "writes a first comment's relationships through the stored sidecar word/_rels/%s when the package has no _rels/.rels (#946)",
        async (relsName) => {
            await withTempDir(async (dir) => {
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(path.join(dir, "word", "_rels", relsName), RELS_XML, "utf-8");
                await fs.writeFile(path.join(dir, "[Content_Types].xml"), CT_XML, "utf-8");

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // With no package relationships the main document is the
                // conventional word/document.xml, and its sidecar is still
                // found under OPC equivalence. The case variant folds into
                // the stored file on a case-insensitive file system, so only
                // the escaped spelling fails there without the lookup.
                expect(await fs.readdir(path.join(dir, "word", "_rels"))).toEqual([relsName]);
                expect(relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", relsName), "utf-8")))).toEqual([
                    ["rId1", STYLES_RELATIONSHIP_TYPE, "styles.xml", ""],
                    ...firstCommentRelationships(2),
                ]);
            });
        },
    );

    it.each([
        ["word/docum%65nt.xml", "docum%65nt.xml.rels"],
        ["word/Document.xml", "Document.xml.rels"],
    ])(
        "creates the sidecar under the main document's stored spelling when %s, a differently spelled word/document.xml, has none (#946)",
        async (mainName, sidecar) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels(mainName), mainName);

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "Hello world",
                    author: "Test",
                    initials: "C",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The created part keeps the main document's own spelling, the
                // one the read model's equivalence lookup finds again; before,
                // the canonical comparison derived
                // word/_rels/document.xml.rels from the conventional name. With
                // no sidecar stored, nothing folds the case variant either, so
                // both rows fail before #946 on any file system.
                expect(await fs.readdir(path.join(dir, "word", "_rels"))).toEqual([sidecar]);
                expect(relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", sidecar), "utf-8")))).toEqual(
                    firstCommentRelationships(1),
                );
            });
        },
    );

    it("keeps creating and typing word/_rels/document.xml.rels when the conventional main document has no sidecar (#895, #901, #946)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml", { relsDefault: false });

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            expect(await fs.readdir(path.join(dir, "word", "_rels"))).toEqual(["document.xml.rels"]);
            expect(relationshipRows(parseXml(await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8")))).toEqual(
                firstCommentRelationships(1),
            );

            // The created part is typed under its own name when nothing
            // else covers it (#895, #901).
            const dom = parseXml(await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8"));
            expect(overridesNaming(dom, "/word/_rels/document.xml.rels")).toEqual([["/word/_rels/document.xml.rels", RELATIONSHIPS_TYPE]]);
        });
    });

    it(
        "validates once the markers are added over a main document stored as word/Document.xml, with no OPC-equivalent part names (#946)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/Document.xml?x=1"), "word/Document.xml");
                await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
                await fs.writeFile(path.join(dir, "word", "_rels", "Document.xml.rels"), RELS_XML, "utf-8");
                await fs.writeFile(path.join(dir, "word", "styles.xml"), MINIMAL_PARTS["word/styles.xml"]!, "utf-8");
                await fs.writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    contentTypes(RELATIONSHIPS_TYPE, [
                        `<Override PartName="/word/Document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPE}"/>`,
                        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
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

                // With the relationships in the stored sidecar, the packed
                // tree holds one spelling of it; the conventional twin the
                // old wiring created on a case-sensitive file system was a
                // second, OPC-equivalent spelling of the same part.
                await fs.writeFile(path.join(dir, "word", "Document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "case-main.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "opc-equivalent-part-names")).toEqual([]);
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
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

    it(
        "appends a second comment to a comments part that binds the WordprocessingML namespace to x:, and the packed tree validates (#930)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const first = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "first",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(first.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                const commentsPath = path.join(dir, "word", "comments.xml");
                const rebound = rebindPrefix(await fs.readFile(commentsPath, "utf-8"), "w", "x");
                expect(rebound).not.toContain("w:");
                await fs.writeFile(commentsPath, rebound, "utf-8");

                // Before #930 the call threw: appendXml looked for a literal
                // <w:comments> root and found none under the x: binding.
                const call = addComment({
                    unpackedDir: dir,
                    commentId: 1,
                    text: "second",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
                });
                await expect(call).resolves.toMatchObject({ paraId: expect.stringMatching(/^[0-9A-F]{8}$/) });
                const second = await call;
                expect(second.message).toBe(`Added comment 1 (para_id=${second.paraId})`);

                // The part holds both comments in the WordprocessingML
                // namespace, whatever prefix each carries: the appended
                // subtree declares its own w: beside the root's x:.
                const dom = parseXml(await fs.readFile(commentsPath, "utf-8"));
                const ids = getElementsByTagNameNSAll(dom, WORD_MAIN_NAMESPACE, "comment").map((comment) =>
                    comment.getAttributeNS(WORD_MAIN_NAMESPACE, "id"),
                );
                expect(ids).toEqual(["0", "1"]);

                // The read model agrees: with markers for both comments in
                // the main document, the packed tree validates under word-valid.
                await fs.writeFile(path.join(dir, "word", "document.xml"), twiceMarkedMainDocument(), "utf-8");
                const file = path.join(dir, "x-prefix-comments.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.issues.filter((issue) => issue.code === "comment-marker-missing")).toEqual([]);
                expect(validation.valid).toBe(true);
            });
        },
    );

    it("refuses a duplicate id on a comments part bound to x:, changing nothing (#930)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "first",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            const commentsPath = path.join(dir, "word", "comments.xml");
            const rebound = rebindPrefix(await fs.readFile(commentsPath, "utf-8"), "w", "x");
            expect(rebound).not.toContain("w:");
            await fs.writeFile(commentsPath, rebound, "utf-8");

            // Before #930 the id check read no comment by its literal w:
            // name, and only the missing literal root made the call throw.
            const before = await readTree(dir);
            const call = addComment({ unpackedDir: dir, commentId: 0, text: "duplicate", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;
            expect(result.message).toBe("Error: Comment id 0 already exists");
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it("answers a reply on a comments part bound to x:, linking the parent by paragraph id (#930)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            const first = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "first",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(first.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            const commentsPath = path.join(dir, "word", "comments.xml");
            const rebound = rebindPrefix(await fs.readFile(commentsPath, "utf-8"), "w", "x");
            expect(rebound).not.toContain("w:");
            await fs.writeFile(commentsPath, rebound, "utf-8");

            // Before #930 the reply could not find its parent: findParaId
            // read no comment by its literal w: name, so it failed with
            // `Error: Parent comment 0 not found` instead.
            const reply = await addComment({
                unpackedDir: dir,
                commentId: 1,
                parent: 0,
                text: "reply",
                author: "Reviewer",
                initials: "R",
                date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
            });
            expect(reply.message).toMatch(/^Added reply 1 \(para_id=[0-9A-F]{8}\)$/);

            // The new w15:commentEx names the parent's paragraph id, which
            // the reply found in the x:-bound comments part.
            const extDom = parseXml(await fs.readFile(path.join(dir, "word", "commentsExtended.xml"), "utf-8"));
            const entries = getElementsByTagNameNSAll(extDom, W15_NAMESPACE, "commentEx");
            expect(entries).toHaveLength(2);
            expect(entries[1]!.getAttributeNS(W15_NAMESPACE, "paraId")).toBe(reply.paraId);
            expect(entries[1]!.getAttributeNS(W15_NAMESPACE, "paraIdParent")).toBe(first.paraId);
        });
    });

    it.each([
        ["commentsExtended.xml", "w15", W15_NAMESPACE, "commentEx"],
        ["commentsIds.xml", "w16cid", W16CID_NAMESPACE, "commentId"],
        ["commentsExtensible.xml", "w16cex", W16CEX_NAMESPACE, "commentExtensible"],
    ])(
        "appends a second comment when word/%s binds its namespace to x:, whatever prefix it carries (#930)",
        async (name, prefix, namespace, entry) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "first",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });

                const partPath = path.join(dir, "word", name);
                const rebound = rebindPrefix(await fs.readFile(partPath, "utf-8"), prefix, "x");
                expect(rebound).not.toContain(`${prefix}:`);
                await fs.writeFile(partPath, rebound, "utf-8");

                // Before #930 the call threw: appendXml looked for the
                // literal <w15:commentsEx> root, or its w16cid / w16cex kin.
                const call = addComment({
                    unpackedDir: dir,
                    commentId: 1,
                    text: "second",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
                });
                await expect(call).resolves.toMatchObject({ paraId: expect.stringMatching(/^[0-9A-F]{8}$/) });

                // The part holds both entries in its namespace, whatever
                // prefix each carries.
                const dom = parseXml(await fs.readFile(partPath, "utf-8"));
                expect(getElementsByTagNameNSAll(dom, namespace, entry)).toHaveLength(2);
            });
        },
    );

    it("refuses a later comment whose comments part holds a root in another namespace, changing nothing (#930)", async () => {
        await withTempDir(async (dir) => {
            await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
            await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "first",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            const commentsPath = path.join(dir, "word", "comments.xml");
            await fs.writeFile(
                commentsPath,
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:comments xmlns:x="urn:example:other"/>',
                "utf-8",
            );

            // Before #930 the call threw instead of returning an error.
            const before = await readTree(dir);
            const call = addComment({ unpackedDir: dir, commentId: 1, text: "second", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;
            expect(result.message).toBe(`Error: ${commentsPath} has no comments root element in the ${WORD_MAIN_NAMESPACE} namespace`);
            expect(await readTree(dir)).toEqual(before);
        });
    });

    it.each([
        ["commentsExtended.xml", "commentsEx", W15_NAMESPACE],
        ["commentsIds.xml", "commentsIds", W16CID_NAMESPACE],
        ["commentsExtensible.xml", "commentsExtensible", W16CEX_NAMESPACE],
    ])(
        "refuses a first comment when the stored word/%s holds a root in another namespace, changing nothing (#930)",
        async (name, localName, namespace) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const partPath = path.join(dir, "word", name);
                await fs.writeFile(
                    partPath,
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:${localName} xmlns:x="urn:example:other"/>`,
                    "utf-8",
                );

                // A first comment appends to a Microsoft part already stored,
                // so its root is checked before comments.xml, the
                // relationships and the content types are written; before,
                // the call threw after writing them.
                const before = await readTree(dir);
                const call = addComment({ unpackedDir: dir, commentId: 0, text: "first", author: "Test", initials: "T" });
                await expect(call).resolves.toMatchObject({ paraId: "" });
                const result = await call;
                expect(result.message).toBe(`Error: ${partPath} has no ${localName} root element in the ${namespace} namespace`);
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );

    it(
        "adds a first comment beside a stored commentsIds.xml that binds its namespace to x:, and the packed tree validates (#930)",
        { timeout: 30000 },
        async () => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const idsPath = path.join(dir, "word", "commentsIds.xml");
                await fs.writeFile(
                    idsPath,
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:commentsIds xmlns:x="${W16CID_NAMESPACE}"/>`,
                    "utf-8",
                );

                const result = await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "first",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // The stored part keeps its x: root and gains the one entry,
                // in its namespace, naming the new comment's paragraph.
                const dom = parseXml(await fs.readFile(idsPath, "utf-8"));
                expect(dom.documentElement!.prefix).toBe("x");
                const entries = getElementsByTagNameNSAll(dom, W16CID_NAMESPACE, "commentId");
                expect(entries.map((entry) => entry.getAttributeNS(W16CID_NAMESPACE, "paraId"))).toEqual([result.paraId]);

                await fs.writeFile(path.join(dir, "word", "document.xml"), markedMainDocument(), "utf-8");
                const file = path.join(dir, "x-prefix-ids.docx");
                await packTree(dir, file);
                const validation = await validate(file, { profile: "word-valid" });
                expect(validation.valid).toBe(true);
            });
        },
    );

    it.each(
        [
            ["comments.xml", "comments", WORD_MAIN_NAMESPACE],
            ["commentsExtended.xml", "commentsEx", W15_NAMESPACE],
            ["commentsIds.xml", "commentsIds", W16CID_NAMESPACE],
            ["commentsExtensible.xml", "commentsExtensible", W16CEX_NAMESPACE],
        ].flatMap(([name, localName, namespace]) => [
            [name, "in another namespace", localName, namespace, "urn:example:other"],
            [name, "of another local name", localName, namespace, namespace],
        ]),
    )(
        "refuses a later comment when the stored word/%s nests its root under a document element %s, changing nothing (#953)",
        async (name, _kind, localName, namespace, wrapperNamespace) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                await addComment({
                    unpackedDir: dir,
                    commentId: 0,
                    text: "first",
                    author: "Test",
                    initials: "T",
                    date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
                });
                const partPath = path.join(dir, "word", name);
                const stored = await fs.readFile(partPath, "utf-8");
                const declaration = stored.startsWith("<?xml") ? stored.slice(0, stored.indexOf("?>") + 2) : "";
                await fs.writeFile(
                    partPath,
                    `${declaration}<x:wrap xmlns:x="${wrapperNamespace}">${stored.slice(declaration.length)}</x:wrap>`,
                    "utf-8",
                );

                // The part still holds its root element, one level down; only
                // the document element counts, so the call is refused before
                // any write instead of appending inside the nested element.
                const before = await readTree(dir);
                const result = await addComment({ unpackedDir: dir, commentId: 1, text: "second", author: "Test", initials: "T" });
                expect(result).toEqual({
                    paraId: "",
                    message: `Error: ${partPath} has no ${localName} root element in the ${namespace} namespace`,
                });
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );

    it.each([
        ["commentsExtended.xml", "commentsEx", W15_NAMESPACE],
        ["commentsIds.xml", "commentsIds", W16CID_NAMESPACE],
        ["commentsExtensible.xml", "commentsExtensible", W16CEX_NAMESPACE],
    ])(
        "refuses a first comment when the stored word/%s nests its root under a foreign document element, changing nothing (#953)",
        async (name, localName, namespace) => {
            await withTempDir(async (dir) => {
                await makePackageFixture(dir, officeDocumentRels("word/document.xml"), "word/document.xml");
                const partPath = path.join(dir, "word", name);
                await fs.writeFile(
                    partPath,
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:wrap xmlns:x="urn:example:other"><m:${localName} xmlns:m="${namespace}"/></x:wrap>`,
                    "utf-8",
                );

                const before = await readTree(dir);
                const result = await addComment({ unpackedDir: dir, commentId: 0, text: "first", author: "Test", initials: "T" });
                expect(result).toEqual({
                    paraId: "",
                    message: `Error: ${partPath} has no ${localName} root element in the ${namespace} namespace`,
                });
                expect(await readTree(dir)).toEqual(before);
            });
        },
    );

    // #939: a part addComment would write must not run through, or end at, a
    // symbolic link: the write would follow it, out of the unpacked tree
    // when the link points outside. `unpack` refuses the same before its
    // first write (#675), and the validators' write-back never follows a
    // link (#425).

    it("refuses a first comment whose relationships part runs through a symbolic link, writing nothing (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            const outside = path.join(dir, "outside");
            await fs.mkdir(tree, { recursive: true });
            await fs.mkdir(outside, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
            // The package holds no relationships part of its own, so the one
            // addComment creates would run through the link.
            await fs.symlink(outside, path.join(tree, "word", "_rels"), "dir");

            const before = await readTree(tree);
            const call = addComment({ unpackedDir: tree, commentId: 0, text: "x", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            // Before #939 the call reported success and wrote the
            // relationships part through the link, out of the tree.
            expect(result.message).toBe(
                `Error: ${path.join(tree, "word", "_rels")} is a symbolic link; addComment does not write through one`,
            );
            expect(await readTree(tree)).toEqual(before);
            expect(await fs.readdir(outside)).toEqual([]);
        });
    });

    it("names the link as the caller spells the unpacked directory (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            const outside = path.join(dir, "outside");
            await fs.mkdir(tree, { recursive: true });
            await fs.mkdir(outside, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.symlink(outside, path.join(tree, "word", "_rels"), "dir");

            // The other refusals name a part under the caller's spelling of
            // the tree, a relative one included; the link is named the same way.
            const relativeTree = path.relative(process.cwd(), tree);
            expect(path.isAbsolute(relativeTree)).toBe(false);
            const result = await addComment({ unpackedDir: relativeTree, commentId: 0, text: "x", author: "Test", initials: "T" });
            expect(result).toEqual({
                paraId: "",
                message: `Error: ${path.join(relativeTree, "word", "_rels")} is a symbolic link; addComment does not write through one`,
            });
            expect(await fs.readdir(outside)).toEqual([]);
        });
    });

    it("refuses a first comment whose relationships part runs through a symbolic link, under a renamed main document (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            const outside = path.join(dir, "outside");
            await fs.mkdir(tree, { recursive: true });
            await fs.mkdir(outside, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/m%61in.xml"), "word/main.xml");
            await fs.symlink(outside, path.join(tree, "word", "_rels"), "dir");

            const before = await readTree(tree);
            const call = addComment({ unpackedDir: tree, commentId: 0, text: "x", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            // The wiring resolves the stored `word/main.xml`, so the part
            // that would run through the link is `word/_rels/main.xml.rels`.
            expect(result.message).toBe(
                `Error: ${path.join(tree, "word", "_rels")} is a symbolic link; addComment does not write through one`,
            );
            expect(await readTree(tree)).toEqual(before);
            expect(await fs.readdir(outside)).toEqual([]);
        });
    });

    it("refuses a first comment whose relationships part runs through a symbolic link inside the tree (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            await fs.mkdir(tree, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.mkdir(path.join(tree, "word", "rels-real"), { recursive: true });
            await fs.symlink(path.join(tree, "word", "rels-real"), path.join(tree, "word", "_rels"), "dir");

            const before = await readTree(tree);
            const call = addComment({ unpackedDir: tree, commentId: 0, text: "x", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            // A link that stays inside the tree is refused too: the write
            // would follow it wherever it names.
            expect(result.message).toBe(
                `Error: ${path.join(tree, "word", "_rels")} is a symbolic link; addComment does not write through one`,
            );
            expect(await readTree(tree)).toEqual(before);
        });
    });

    it("refuses a later comment whose comments part is a symbolic link, changing nothing (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            await fs.mkdir(tree, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
            const first = await addComment({
                unpackedDir: tree,
                commentId: 0,
                text: "first",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(first.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The stored comments part becomes a link to a file elsewhere in
            // the temporary directory.
            const target = path.join(dir, "comments-elsewhere.xml");
            const commentsPath = path.join(tree, "word", "comments.xml");
            await fs.rename(commentsPath, target);
            await fs.symlink(target, commentsPath, "file");
            const targetBefore = await fs.readFile(target);

            const before = await readTree(tree);
            const call = addComment({ unpackedDir: tree, commentId: 1, text: "second", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            // Before #939 the call appended the comment to the link's
            // target, outside the tree.
            expect(result.message).toBe(`Error: ${commentsPath} is a symbolic link; addComment does not write through one`);
            expect(await readTree(tree)).toEqual(before);
            expect(await fs.readFile(target)).toEqual(targetBefore);
        });
    });

    it.each(["commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"])(
        "refuses a later comment whose %s is a symbolic link, changing nothing (#939)",
        async (name) => {
            await withTempDir(async (dir) => {
                const tree = path.join(dir, "unpacked");
                await fs.mkdir(tree, { recursive: true });
                await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
                const first = await addComment({ unpackedDir: tree, commentId: 0, text: "first", author: "Test", initials: "T" });
                expect(first.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

                // Every part a later comment appends to is checked, not only
                // the comments part.
                const target = path.join(dir, `${name}-elsewhere`);
                const partPath = path.join(tree, "word", name);
                await fs.rename(partPath, target);
                await fs.symlink(target, partPath, "file");
                const targetBefore = await fs.readFile(target);

                const before = await readTree(tree);
                const result = await addComment({ unpackedDir: tree, commentId: 1, text: "second", author: "Test", initials: "T" });
                expect(result).toEqual({
                    paraId: "",
                    message: `Error: ${partPath} is a symbolic link; addComment does not write through one`,
                });
                expect(await readTree(tree)).toEqual(before);
                expect(await fs.readFile(target)).toEqual(targetBefore);
            });
        },
    );

    it("refuses a first comment whose [Content_Types].xml is a symbolic link, changing nothing (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            await fs.mkdir(tree, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");

            // A first comment also writes the content-types stream.
            const target = path.join(dir, "content-types-elsewhere.xml");
            const ctPath = path.join(tree, "[Content_Types].xml");
            await fs.rename(ctPath, target);
            await fs.symlink(target, ctPath, "file");
            const targetBefore = await fs.readFile(target);

            const before = await readTree(tree);
            const result = await addComment({ unpackedDir: tree, commentId: 0, text: "first", author: "Test", initials: "T" });
            expect(result).toEqual({
                paraId: "",
                message: `Error: ${ctPath} is a symbolic link; addComment does not write through one`,
            });
            expect(await readTree(tree)).toEqual(before);
            expect(await fs.readFile(target)).toEqual(targetBefore);
        });
    });

    it("refuses a later comment whose comments part is a link to a part it cannot decode, naming the link (#939)", async () => {
        await withTempDir(async (dir) => {
            const tree = path.join(dir, "unpacked");
            await fs.mkdir(tree, { recursive: true });
            await makePackageFixture(tree, officeDocumentRels("word/document.xml"), "word/document.xml");
            const first = await addComment({
                unpackedDir: tree,
                commentId: 0,
                text: "first",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(first.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            const target = path.join(dir, "comments-elsewhere.xml");
            const commentsPath = path.join(tree, "word", "comments.xml");
            await fs.rename(commentsPath, target);
            await declareUndecodable(target);
            await fs.symlink(target, commentsPath, "file");

            const before = await readTree(tree);
            const call = addComment({ unpackedDir: tree, commentId: 1, text: "second", author: "Test", initials: "T" });
            await expect(call).resolves.toMatchObject({ paraId: "" });
            const result = await call;

            // The link is refused before the preflight reads the target
            // through it, so the link is named and not the encoding.
            expect(result.message).toBe(`Error: ${commentsPath} is a symbolic link; addComment does not write through one`);
            expect(await readTree(tree)).toEqual(before);
        });
    });

    it("adds a comment to an unpacked directory reached through a symbolic link: the caller chose it (#939)", async () => {
        await withTempDir(async (dir) => {
            const real = path.join(dir, "real");
            const link = path.join(dir, "link");
            await fs.mkdir(real, { recursive: true });
            await makePackageFixture(real, officeDocumentRels("word/document.xml"), "word/document.xml");
            await fs.symlink(real, link, "dir");

            const result = await addComment({
                unpackedDir: link,
                commentId: 0,
                text: "x",
                author: "Test",
                initials: "T",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);

            // The parts land in the directory the link names.
            const dom = parseXml(await fs.readFile(path.join(real, "word", "comments.xml"), "utf-8"));
            expect(
                getElementsByTagNameNSAll(dom, WORD_MAIN_NAMESPACE, "comment").map((comment) =>
                    comment.getAttributeNS(WORD_MAIN_NAMESPACE, "id"),
                ),
            ).toEqual(["0"]);
        });
    });
});
