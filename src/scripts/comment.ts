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

/**
 * Add comments to DOCX documents.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/comment.py` (task #8).
 *
 * NOTE: `src/html-to-docx/src/tracking.ts` already implements a higher-level
 * comment / suggestion token system that knows about ranges, references, and
 * threading. This file is intentionally a verbatim port of the Python helper
 * — it operates directly on the unpacked OOXML parts (comments.xml,
 * commentsExtended.xml, commentsIds.xml, commentsExtensible.xml, plus
 * `_rels/document.xml.rels` and `[Content_Types].xml`) and is meant for
 * scripts / agents that need to mutate the on-disk DOCX bundle. Reach for
 * `tracking.ts` if you want a structured API instead.
 *
 * Usage:
 *   pnpm exec tsx scripts/comment.ts unpacked/ 0 "Comment text"
 *   pnpm exec tsx scripts/comment.ts unpacked/ 1 "Reply text" --parent 0
 *
 * Text should be pre-escaped XML (e.g., &amp; for &, &#x2019; for smart
 * quotes).
 *
 * After running, add markers to document.xml:
 *   <w:commentRangeStart w:id="0"/>
 *   ... commented content ...
 *   <w:commentRangeEnd w:id="0"/>
 *   <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import {
    canonicalPartName,
    contentTypeMediaType,
    CONTENT_TYPES_NAMESPACE,
    findEquivalentPart,
    parseContentTypes,
    relationshipTargetPartName,
    resolveContentType,
} from "../lib/content-types";
import { applyMarkupCompatibility, RELATIONSHIPS_UNDERSTOOD_NAMESPACES } from "../lib/markup-compatibility";
import { commanderExitCode, runCli } from "../lib/run-cli";
import { nextSecureLongHexNumber } from "../lib/secure-id";
import { NS as XML_NS } from "../lib/types";
import { readXmlPart, writeXmlPart, writeXmlText, type DecodedXmlPart } from "../lib/xml-part-io";
import { getElementsByTagNameNSAll, getRelationshipChildren, parseXml, serializeXml, xmlDeclarationEncoding } from "../lib/xml-helpers";

const TEMPLATE_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "templates");

const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
/** The relationships part addComment creates for a main document that has none (#895). */
const EMPTY_RELATIONSHIPS_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

/** The package relationships part, as it is stored at the package root. */
const PACKAGE_RELS_PART = "_rels/.rels";
/** The relationship types that name a WordprocessingML main document, in either conformance class. */
const OFFICE_DOCUMENT_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
]);
/** The relationship types that relate the comments part, in either conformance class. */
const COMMENTS_PART_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments",
]);
/** The conventional main document, whose relationships part is `word/_rels/document.xml.rels` (#906). */
const CONVENTIONAL_MAIN_DOCUMENT = "word/document.xml";
/** The comments parts, in the order their relationships are written. */
const COMMENT_PART_NAMES: readonly string[] = ["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"];
/** The relationship type reaching each comments part, in `COMMENT_PART_NAMES` order. */
const COMMENT_RELATIONSHIP_TYPES: readonly string[] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
];

const NS: Record<string, string> = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
};

const COMMENT_XML = `<w:comment w:id="{id}" w:author="{author}" w:date="{date}" w:initials="{initials}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r>
  </w:p>
</w:comment>`;

const COMMENT_MARKER_TEMPLATE = `
Add to document.xml (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>`;

const REPLY_MARKER_TEMPLATE = `
Nest markers inside parent {pid}'s markers (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{pid}"/><w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/><w:commentRangeEnd w:id="{pid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{pid}"/></w:r>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>`;

/**
 * Generate an 8-digit hex id for `w14:paraId` / `w16cid:durableId`.
 *
 * [MS-OI29500] 2.6.2.3 types both as `ST_LongHexNumber` whose value MUST be
 * greater than 0 and less than 0x80000000. The `1 +` lower bound mirrors the
 * repair path in `validators/docx.ts` (which seeds durableIds the same way) so
 * we never write a zero or an over-cap id of our own. Range: [1, 0x7FFFFFFE],
 * which also clears the tighter durableId cap of 0x7FFFFFFF.
 */
export function generateHexId(): string {
    return nextSecureLongHexNumber().toString(16).toUpperCase().padStart(8, "0");
}

const SMART_QUOTE_ENTITIES: Record<string, string> = {
    "“": "&#x201C;",
    "”": "&#x201D;",
    "‘": "&#x2018;",
    "’": "&#x2019;",
};

function encodeSmartQuotes(text: string): string {
    let out = text;
    for (const [char, entity] of Object.entries(SMART_QUOTE_ENTITIES)) {
        out = out.split(char).join(entity);
    }
    return out;
}

function escapeXml(text: string): string {
    return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function formatTemplate(tpl: string, vars: Record<string, string | number>): string {
    return tpl.replace(/\{(\w+)\}/g, (_, key: string) => {
        if (key in vars) {
            return String(vars[key]);
        }
        return `{${key}}`;
    });
}

async function pathExists(p: string): Promise<boolean> {
    try {
        await fs.access(p);
        return true;
    } catch {
        return false;
    }
}

/** Append `content` under the part's root element, in the part's own encoding. */
async function appendXml(xmlPath: string, rootTag: string, content: string, decoded?: DecodedXmlPart): Promise<boolean> {
    const part = decoded ?? (await readXmlPart(xmlPath));
    if (part === undefined) {
        // An encoding we cannot decode: leave the part's bytes untouched.
        return false;
    }
    const dom = parseXml(part.text);
    const roots = dom.getElementsByTagName(rootTag);
    if (roots.length === 0) {
        throw new Error(`Root tag <${rootTag}> not found in ${xmlPath}`);
    }
    const root = roots.item(0)!;
    const nsAttrs = Object.entries(NS)
        .map(([k, v]) => `xmlns:${k}="${v}"`)
        .join(" ");
    const wrapperDom = parseXml(`<root ${nsAttrs}>${content}</root>`);
    const children = wrapperDom.documentElement.childNodes;
    for (let i = 0; i < children.length; i += 1) {
        const child = children.item(i);
        if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
            root.appendChild(dom.importNode(child, true));
        }
    }
    const output = encodeSmartQuotes(serializeXml(dom, xmlDeclarationEncoding(part)));
    await writeXmlText(xmlPath, output, part);
    return true;
}

function findParaId(commentsPart: DecodedXmlPart, commentId: number): string | null {
    const dom = parseXml(commentsPart.text);
    const comments = dom.getElementsByTagName("w:comment");
    for (let i = 0; i < comments.length; i += 1) {
        const c = comments.item(i)!;
        if (c.getAttribute("w:id") === String(commentId)) {
            const ps = c.getElementsByTagName("w:p");
            for (let j = 0; j < ps.length; j += 1) {
                const p = ps.item(j)!;
                const pid = p.getAttribute("w14:paraId");
                if (pid) {
                    return pid;
                }
            }
        }
    }
    return null;
}

/** Return true if the comments part already contains a `<w:comment>` with the
 * given `w:id`. Used to preflight `addComment()` before mutating any part. */
function commentIdExists(commentsPart: DecodedXmlPart, commentId: number): boolean {
    const dom = parseXml(commentsPart.text);
    const comments = dom.getElementsByTagName("w:comment");
    for (let i = 0; i < comments.length; i += 1) {
        const c = comments.item(i)!;
        if (c.getAttribute("w:id") === String(commentId)) {
            return true;
        }
    }
    return false;
}

/**
 * The next free relationship id: one past the `rId<N>` id of every element
 * whose local name is `Relationship`, whatever its namespace or prefix, so a
 * new id never repeats one the part already holds — in a part whose
 * relationships are prefixed, or one holding a `Relationship` element in
 * another namespace, alike (#918).
 */
function getNextRid(dom: Document): number {
    let maxRid = 0;
    const rels = getElementsByTagNameNSAll(dom, "*", "Relationship");
    for (let i = 0; i < rels.length; i += 1) {
        const rel = rels[i]!;
        const rid = rel.getAttribute("Id");
        if (rid && rid.startsWith("rId")) {
            const n = Number.parseInt(rid.slice(3), 10);
            if (Number.isFinite(n)) {
                maxRid = Math.max(maxRid, n);
            }
        }
    }
    return maxRid + 1;
}

/**
 * The relationships a stored relationships part holds as the checks read
 * them: markup compatibility selects a branch first, with the namespaces the
 * read model understands in a relationships part, so a relationship a
 * selected `mc:Fallback` or an understood `mc:Choice` holds counts, and one
 * only an unselected branch holds does not (#927); before, the stored root's
 * direct children were read as stored, so a selected branch's relationship
 * was read as absent — a package whose `officeDocument` relationship sat in
 * one was refused, and a `comments` relationship in one got a second set
 * beside it (#932). The issues the processor reports (`mce-malformed`, for
 * example) are not addComment's to judge: validation reports them.
 * `getNextRid` and every write keep the stored DOM (#918).
 */
function effectiveRelationships(dom: Document): Element[] {
    return getRelationshipChildren(applyMarkupCompatibility(dom, { understood: RELATIONSHIPS_UNDERSTOOD_NAMESPACES }).document);
}

/**
 * Whether a relationship in `dom` relates `partName`: an internal
 * relationship of the comments type, Transitional or Strict, whose target
 * resolves to `partName` under OPC part-name equivalence (#906). The
 * relationships are read on the effective view `effectiveRelationships`
 * returns — the stored DOM after markup compatibility selects a branch —
 * so a relationship a selected `mc:Fallback` or an understood `mc:Choice`
 * holds counts, and one only an unselected branch holds does not (#932,
 * #927); a part whose root binds the package-relationships namespace to a
 * prefix is read too, and an element in another namespace is not read as a
 * relationship, even one named `Relationship` (#918). An External
 * relationship, or one of another type, does not count; before, it stopped
 * addComment adding the comments relationships, so it reported success with
 * the comments parts related to nothing (#917).
 */
function relatesCommentsPart(dom: Document, baseDir: string, partName: string): boolean {
    const key = canonicalPartName(partName);
    const rels = effectiveRelationships(dom);
    for (let i = 0; i < rels.length; i += 1) {
        const rel = rels[i]!;
        if (!COMMENTS_PART_RELATIONSHIP_TYPES.has(rel.getAttribute("Type") ?? "")) {
            continue;
        }
        if (rel.getAttribute("TargetMode") === "External") {
            continue;
        }
        const resolved = resolveTargetFrom(baseDir, rel.getAttribute("Target") ?? "");
        if (resolved !== undefined && canonicalPartName(resolved) === key) {
            return true;
        }
    }
    return false;
}

/**
 * Resolve a relationship target against `baseDir`, the folder of the part
 * whose relationships part holds it: a leading `/` is absolute from the
 * package root, anything else is relative to the folder, and `.` and `..`
 * segments normalize. `undefined` when the target is empty or climbs past
 * the package root.
 */
function resolveTargetFrom(baseDir: string, target: string): string | undefined {
    if (target === "") {
        return undefined;
    }
    const reference = target.startsWith("/") ? target.slice(1) : path.posix.join(baseDir, target);
    const normalized = path.posix.normalize(reference);
    if (normalized === ".." || normalized.startsWith("../")) {
        return undefined;
    }
    return normalized;
}

/**
 * Whether a part reference climbs past the package root, from the root
 * (`/../x.xml`) or relative (`../x.xml`), once its dot segments resolve.
 */
function climbsPastPackageRoot(reference: string): boolean {
    const normalized = path.posix.normalize(reference.replace(/^\//, ""));
    return normalized === ".." || normalized.startsWith("../");
}

/**
 * Every file under `unpackedDir`, as a forward-slash part name relative to
 * it — the stored part names a relationship target is looked up among.
 */
async function storedPartNames(unpackedDir: string): Promise<string[]> {
    const names: string[] = [];
    const walk = async (dir: string, prefix: string): Promise<void> => {
        for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
            if (entry.isFile()) {
                names.push(prefix === "" ? entry.name : `${prefix}/${entry.name}`);
            } else if (entry.isDirectory()) {
                await walk(path.join(dir, entry.name), prefix === "" ? entry.name : `${prefix}/${entry.name}`);
            }
        }
    };
    await walk(unpackedDir, "");
    return names;
}

/**
 * Where a first comment's relationships are wired through (#906): the main
 * document `_rels/.rels` names, resolved as the read model resolves it
 * (`\` as `/`, the query and fragment dropped) and looked up among the
 * stored parts under OPC part-name equivalence, so `word/m%61in.xml` names
 * a stored `word/main.xml`; before, a query, a fragment or an escaped
 * spelling was refused (#919). A missing `_rels/.rels`, or one whose stored
 * part is the conventional `word/document.xml`, keeps the wiring #895
 * established; a package naming another main document wires through that
 * document's own relationships part, under its stored name. The
 * relationships are read on the effective view `effectiveRelationships`
 * returns — the stored DOM after markup compatibility selects a branch —
 * so a relationship a selected `mc:Fallback` or an understood `mc:Choice`
 * holds counts, and one only an unselected branch holds does not (#932); a
 * part whose root binds the package-relationships namespace to a prefix is
 * read too, and an element in another namespace is not read as a
 * relationship (#918).
 * What cannot be resolved — a URI scheme, an empty reference, a target that
 * climbs past the package root, or a named part that is not stored, the
 * conventional `word/document.xml` included — is refused, naming `_rels/.rels`
 * or the missing part, before any write (#919).
 */
type CommentWiring = { kind: "conventional" } | { kind: "wired"; mainDocument: string } | { kind: "refused"; message: string };

async function resolveCommentWiring(unpackedDir: string): Promise<CommentWiring> {
    const packageRelsPath = path.join(unpackedDir, ...PACKAGE_RELS_PART.split("/"));
    if (!(await pathExists(packageRelsPath))) {
        return { kind: "conventional" };
    }
    const part = await readXmlPart(packageRelsPath);
    if (part === undefined) {
        return { kind: "refused", message: `Error: ${packageRelsPath} is stored in an encoding this tool cannot decode` };
    }
    let dom: Document;
    try {
        dom = parseXml(part.text);
    } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        return { kind: "refused", message: `Error: ${packageRelsPath} is not well-formed XML: ${detail}` };
    }
    const relationships = effectiveRelationships(dom);
    let mainRel: Element | undefined;
    let sawOfficeDocument = false;
    for (let i = 0; i < relationships.length; i += 1) {
        const rel = relationships[i]!;
        if (!OFFICE_DOCUMENT_RELATIONSHIP_TYPES.has(rel.getAttribute("Type") ?? "")) {
            continue;
        }
        sawOfficeDocument = true;
        if (rel.getAttribute("TargetMode") === "External") {
            continue;
        }
        mainRel = rel;
        break;
    }
    if (mainRel === undefined) {
        return {
            kind: "refused",
            message: sawOfficeDocument
                ? `Error: ${packageRelsPath} names an External officeDocument relationship`
                : `Error: ${packageRelsPath} names no officeDocument relationship`,
        };
    }
    // The reference is resolved as the read model resolves it (a URI scheme
    // or an empty target names no part, the query and fragment are dropped),
    // and a reference that climbs past the package root names none either.
    const reference = relationshipTargetPartName(".", mainRel.getAttribute("Target") ?? "");
    if (reference === undefined || climbsPastPackageRoot(reference)) {
        return { kind: "refused", message: `Error: ${packageRelsPath} names a main document outside the package` };
    }
    // The reference names the stored part the read model would find, under
    // OPC equivalence, and the wiring goes through that part's own name; a
    // named part that is not stored is refused before any write, the
    // conventional `word/document.xml` included (#919).
    const storedName = findEquivalentPart(await storedPartNames(unpackedDir), reference);
    if (storedName === undefined) {
        return { kind: "refused", message: `Error: ${path.join(unpackedDir, ...reference.replace(/^\//, "").split("/"))} not found` };
    }
    // OPC part names compare ASCII case-insensitively, with unreserved
    // escapes decoded, so `Word/Document.xml` names the conventional part.
    if (canonicalPartName(storedName) === canonicalPartName(CONVENTIONAL_MAIN_DOCUMENT)) {
        return { kind: "conventional" };
    }
    return { kind: "wired", mainDocument: storedName };
}

/**
 * Whether the content-types stream holds an Override for `partName`, read
 * in the content-types namespace so a stream that binds the namespace to a
 * prefix is seen, and one in another namespace is not (#923). The `PartName`
 * comparison stays exact: OPC equivalence there is a separate question.
 */
function hasContentType(dom: Document, partName: string): boolean {
    for (const override of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override")) {
        if (override.getAttribute("PartName") === partName) {
            return true;
        }
    }
    return false;
}

/**
 * Relate the comments parts from the main document's relationships part,
 * stored at `relsPath` with each part's `target` relative to the main
 * document's folder `baseDir` (#906). New relationships are created in the
 * package-relationships namespace under the root's prefix, so a prefixed
 * part keeps its prefix on every element and an unprefixed one serializes
 * as it always has, with no declaration of the new elements' own (#918).
 * Returns whether it created that part: a main document with no
 * relationships part is legal, but the comments parts are reached only
 * through relationships, so it creates one rather than leave them related
 * to nothing (#895).
 */
async function ensureCommentRelationships(
    relsPath: string,
    baseDir: string,
    targets: readonly string[],
    decoded?: DecodedXmlPart,
): Promise<boolean> {
    const created = !(await pathExists(relsPath));
    if (created) {
        await fs.mkdir(path.dirname(relsPath), { recursive: true });
        await fs.writeFile(relsPath, EMPTY_RELATIONSHIPS_XML, "utf-8");
    }

    const part = decoded ?? (await readXmlPart(relsPath));
    if (part === undefined) {
        // addComment's preflight decodes this part before any write, so this is
        // reachable only when the part changed mid-run: refuse, never report a
        // write that did not happen (#879).
        throw new Error(`${relsPath} is stored in an encoding this tool cannot decode`);
    }
    const dom = parseXml(part.text);
    if (relatesCommentsPart(dom, baseDir, path.posix.join("word", "comments.xml"))) {
        return created;
    }
    const root = dom.documentElement;
    // The root may bind the package-relationships namespace to a prefix; the
    // new relationships carry that prefix, and no declaration of their own.
    const qualifiedName = root.prefix ? `${root.prefix}:Relationship` : "Relationship";
    let nextRid = getNextRid(dom);

    for (let i = 0; i < COMMENT_RELATIONSHIP_TYPES.length; i += 1) {
        const rel = dom.createElementNS(XML_NS.PR, qualifiedName);
        rel.setAttribute("Id", `rId${nextRid}`);
        rel.setAttribute("Type", COMMENT_RELATIONSHIP_TYPES[i]!);
        rel.setAttribute("Target", targets[i]!);
        root.appendChild(rel);
        nextRid += 1;
    }

    await writeXmlPart(relsPath, dom, part);
    return created;
}

/**
 * Type the comments parts in `[Content_Types].xml`, and the main document's
 * relationships part — named `relsPartName`, the part's own name — when
 * `createdRels` says addComment created it and nothing resolves it to the
 * relationships media type (#895, #906, #901): a Default of another type is
 * left alone and answered with an Override under the part's own name, and
 * an Override that names the part under OPC equivalence with another type
 * is rewritten in place, never answered with a second Override. A
 * relationships part that already existed is never retyped. The stream is
 * read and written in its namespace, whatever prefix it binds it to: an
 * Override is read only in the content-types namespace, and each new one is
 * created there carrying the prefix the root `Types` element uses, so an
 * unprefixed stream serializes exactly as before (#923).
 */
async function ensureCommentContentTypes(
    unpackedDir: string,
    createdRels: boolean,
    relsPartName: string,
    decoded?: DecodedXmlPart,
): Promise<void> {
    const ctPath = path.join(unpackedDir, "[Content_Types].xml");
    if (!(await pathExists(ctPath))) {
        return;
    }

    const part = decoded ?? (await readXmlPart(ctPath));
    if (part === undefined) {
        // addComment's preflight decodes this part before any write, so this is
        // reachable only when the part changed mid-run: refuse, never report a
        // write that did not happen (#879).
        throw new Error(`${ctPath} is stored in an encoding this tool cannot decode`);
    }
    const dom = parseXml(part.text);
    const overrides: Array<[string, string]> = [];
    if (!hasContentType(dom, "/word/comments.xml")) {
        overrides.push(
            ["/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"],
            ["/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"],
            ["/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"],
            ["/word/commentsExtensible.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"],
        );
    }
    // Coverage alone is not correctness: a `.rels` Default of another type
    // covered the created part with the wrong type while addComment reported
    // success (#901). The media type comparison ignores case and parameters.
    let mistypedRelsOverride = false;
    if (createdRels) {
        const resolved = resolveContentType(parseContentTypes(dom), relsPartName);
        if (
            resolved === null ||
            (resolved.source === "default" &&
                contentTypeMediaType(resolved.contentType) !== contentTypeMediaType(RELATIONSHIPS_CONTENT_TYPE))
        ) {
            overrides.push([relsPartName, RELATIONSHIPS_CONTENT_TYPE]);
        } else if (
            resolved.source === "override" &&
            contentTypeMediaType(resolved.contentType) !== contentTypeMediaType(RELATIONSHIPS_CONTENT_TYPE)
        ) {
            mistypedRelsOverride = true;
        }
    }
    if (overrides.length === 0 && !mistypedRelsOverride) {
        return;
    }
    const root = dom.documentElement;
    const overridePrefix = root.prefix ? `${root.prefix}:` : "";

    for (const [partName, contentType] of overrides) {
        const override = dom.createElementNS(CONTENT_TYPES_NAMESPACE, `${overridePrefix}Override`);
        override.setAttribute("PartName", partName);
        override.setAttribute("ContentType", contentType);
        root.appendChild(override);
    }
    if (mistypedRelsOverride) {
        const key = canonicalPartName(relsPartName);
        for (const override of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override")) {
            if (canonicalPartName(override.getAttribute("PartName") ?? "") === key) {
                override.setAttribute("ContentType", RELATIONSHIPS_CONTENT_TYPE);
            }
        }
    }

    await writeXmlPart(ctPath, dom, part);
}

export interface AddCommentOptions {
    unpackedDir: string;
    commentId: number;
    text: string;
    parent?: number;
    /** Comment author. Required. */
    author: string;
    /** Author initials. Required (Word renders this in the comment thumbnail). */
    initials: string;
    /** Override the timestamp; defaults to `new Date()`. Format: `YYYY-MM-DDTHH:MM:SSZ`. */
    date?: Date;
}

export interface AddCommentResult {
    paraId: string;
    message: string;
}

/**
 * Add a comment (or reply) to an unpacked DOCX.
 *
 * On success, returns the generated `paraId` and a human-readable status
 * message. On failure (missing `word/` directory, missing parent comment, a
 * part this edit would write that is stored in an encoding this tool cannot
 * decode (#879) or is not well-formed XML (#905), or, for a first comment,
 * missing `[Content_Types].xml`, an `_rels/.rels` that cannot be decoded or
 * parsed, an `_rels/.rels` naming no non-External `officeDocument`
 * relationship, or an `officeDocument` target that, resolved as the read
 * model resolves it (`\` as `/`, the query and fragment dropped, the
 * stored part matched under OPC part-name equivalence), is a URI, escapes
 * the package root or names no stored part — the conventional
 * `word/document.xml` included (#906, #919)), returns `paraId = ""`
 * and a message starting with `Error:`.
 */
export async function addComment(opts: AddCommentOptions): Promise<AddCommentResult> {
    const { unpackedDir, commentId, text, parent, author, initials, date } = opts;

    const word = path.join(unpackedDir, "word");
    if (!(await pathExists(word))) {
        return { paraId: "", message: `Error: ${word} not found` };
    }

    const paraId = generateHexId();
    const durableId = generateHexId();
    const ts = formatUtcTimestamp(date ?? new Date());

    const comments = path.join(word, "comments.xml");
    const ext = path.join(word, "commentsExtended.xml");
    const ids = path.join(word, "commentsIds.xml");
    const extensible = path.join(word, "commentsExtensible.xml");
    const ctPath = path.join(unpackedDir, "[Content_Types].xml");
    const firstComment = !(await pathExists(comments));
    if (firstComment && !(await pathExists(ctPath))) {
        // A first comment's parts must be typed in the content-types stream.
        // A package without one cannot be made valid here, so refuse before
        // anything is written rather than report a success (#895).
        return { paraId: "", message: `Error: ${ctPath} not found` };
    }

    // A first comment relates the comments parts from the main document's
    // relationships part, which belongs to the main document `_rels/.rels`
    // names — not necessarily `word/document.xml` (#906). A later comment
    // never touches relationships, so it reads no `_rels/.rels` at all.
    let mainDocument = CONVENTIONAL_MAIN_DOCUMENT;
    if (firstComment) {
        const wiring = await resolveCommentWiring(unpackedDir);
        if (wiring.kind === "refused") {
            return { paraId: "", message: wiring.message };
        }
        if (wiring.kind === "wired") {
            mainDocument = wiring.mainDocument;
        }
    }
    const mainFolder = path.posix.dirname(mainDocument);
    const mainRelsPart = path.posix.join(mainFolder, "_rels", `${path.posix.basename(mainDocument)}.rels`);
    const mainRelsPath = path.join(unpackedDir, ...mainRelsPart.split("/"));
    const commentTargets = COMMENT_PART_NAMES.map((name) => path.posix.relative(mainFolder, path.posix.join("word", name)));

    // Preflight: decode and parse every existing part this edit may write,
    // before any file is created, copied or changed. The main document's
    // relationships part and `[Content_Types].xml` are written only for the
    // first comment. A part whose encoding cannot be decoded is left
    // byte-for-byte untouched (#866), so an undecodable part found mid-write
    // used to leave earlier parts already changed, or report success for a
    // wiring it skipped (#879). A part that decodes but does not parse used
    // to throw mid-write the same way, so it is refused here too (#905).
    const writtenParts = [comments, ext, ids, extensible];
    if (firstComment) {
        writtenParts.push(mainRelsPath, ctPath);
    }
    const decoded = new Map<string, DecodedXmlPart>();
    for (const part of writtenParts) {
        if (!(await pathExists(part))) {
            continue;
        }
        const xml = await readXmlPart(part);
        if (xml === undefined) {
            return { paraId: "", message: `Error: ${part} is stored in an encoding this tool cannot decode` };
        }
        try {
            parseXml(xml.text);
        } catch (err) {
            const detail = err instanceof Error ? err.message : String(err);
            return { paraId: "", message: `Error: ${part} is not well-formed XML: ${detail}` };
        }
        decoded.set(part, xml);
    }

    // Preflight validation: check duplicate commentId and resolve `parent`
    // BEFORE we mutate any part. Without this, an existing comments.xml could
    // be mutated and then the parent-not-found branch leaves an orphan comment
    // behind on a failed call.
    if (!firstComment) {
        if (commentIdExists(decoded.get(comments)!, commentId)) {
            return {
                paraId: "",
                message: `Error: Comment id ${commentId} already exists`,
            };
        }
        if (parent !== undefined) {
            const parentPara = findParaId(decoded.get(comments)!, parent);
            if (!parentPara) {
                return {
                    paraId: "",
                    message: `Error: Parent comment ${parent} not found`,
                };
            }
        }
    } else if (parent !== undefined) {
        // First comment in the document but a parent was requested — the
        // parent cannot exist yet, so reject before creating template files.
        return { paraId: "", message: `Error: Parent comment ${parent} not found` };
    }

    if (firstComment) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "comments.xml"), comments);
        const createdRels = await ensureCommentRelationships(mainRelsPath, mainFolder, commentTargets, decoded.get(mainRelsPath));
        await ensureCommentContentTypes(unpackedDir, createdRels, `/${mainRelsPart}`, decoded.get(ctPath));
    }
    if (
        !(await appendXml(
            comments,
            "w:comments",
            formatTemplate(COMMENT_XML, {
                id: commentId,
                author: escapeXml(author),
                date: ts,
                initials: escapeXml(initials),
                para_id: paraId,
                text: escapeXml(text),
            }),
            decoded.get(comments),
        ))
    ) {
        return { paraId: "", message: `Error: ${comments} is stored in an encoding this tool cannot decode` };
    }

    if (!(await pathExists(ext))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsExtended.xml"), ext);
    }

    let commentEx = `<w15:commentEx w15:paraId="${paraId}" w15:done="0"/>`;
    if (parent !== undefined) {
        // Already preflighted above; resolve the actual paraId here for the link.
        const parentPara = findParaId(decoded.get(comments)!, parent);
        if (!parentPara) {
            // Defensive: should be unreachable given preflight, but keep a
            // structured failure rather than a silent bad ref.
            return {
                paraId: "",
                message: `Error: Parent comment ${parent} not found`,
            };
        }
        commentEx = `<w15:commentEx w15:paraId="${paraId}" w15:paraIdParent="${parentPara}" w15:done="0"/>`;
    }

    if (!(await pathExists(ids))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsIds.xml"), ids);
    }

    if (!(await pathExists(extensible))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsExtensible.xml"), extensible);
    }

    if (!(await appendXml(ext, "w15:commentsEx", commentEx, decoded.get(ext)))) {
        return { paraId: "", message: `Error: ${ext} is stored in an encoding this tool cannot decode` };
    }
    if (
        !(await appendXml(
            ids,
            "w16cid:commentsIds",
            `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/>`,
            decoded.get(ids),
        ))
    ) {
        return { paraId: "", message: `Error: ${ids} is stored in an encoding this tool cannot decode` };
    }
    if (
        !(await appendXml(
            extensible,
            "w16cex:commentsExtensible",
            `<w16cex:commentExtensible w16cex:durableId="${durableId}" w16cex:dateUtc="${ts}"/>`,
            decoded.get(extensible),
        ))
    ) {
        return { paraId: "", message: `Error: ${extensible} is stored in an encoding this tool cannot decode` };
    }

    const action = parent !== undefined ? "reply" : "comment";
    return {
        paraId,
        message: `Added ${action} ${commentId} (para_id=${paraId})`,
    };
}

function formatUtcTimestamp(d: Date): string {
    const pad = (n: number) => String(n).padStart(2, "0");
    return (
        `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` +
        `T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`
    );
}

export function buildCommentCommand(): Command {
    const cmd = new Command();
    cmd.name("comment")
        .description("Add comments to DOCX documents")
        .argument("<unpackedDir>", "Unpacked DOCX directory")
        .argument("<commentId>", "Comment ID (must be unique)")
        .argument("<text>", "Comment text")
        .requiredOption("--author <author>", "Author name (required)")
        .requiredOption("--initials <initials>", "Author initials (required; Word renders these in the comment thumbnail)")
        .option("--parent <parent>", "Parent comment ID (for replies)");
    return cmd;
}

interface CliOptions {
    author: string;
    initials: string;
    parent?: string;
}

export async function runCommentFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildCommentCommand();
    cmd.exitOverride();
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [unpackedDir, commentIdRaw, text] = cmd.args;

    const commentId = Number.parseInt(commentIdRaw, 10);
    if (!Number.isFinite(commentId)) {
        process.stderr.write(`Error: commentId must be an integer (got ${commentIdRaw})\n`);
        return 1;
    }
    const parent = opts.parent !== undefined ? Number.parseInt(opts.parent, 10) : undefined;
    if (parent !== undefined && !Number.isFinite(parent)) {
        process.stderr.write(`Error: --parent must be an integer (got ${opts.parent})\n`);
        return 1;
    }

    const result = await addComment({
        unpackedDir,
        commentId,
        text,
        author: opts.author,
        initials: opts.initials,
        parent,
    });
    process.stdout.write(`${result.message}\n`);
    if (result.message.includes("Error")) {
        return 1;
    }
    if (parent !== undefined) {
        process.stdout.write(`${formatTemplate(REPLY_MARKER_TEMPLATE, { pid: parent, cid: commentId })}\n`);
    } else {
        process.stdout.write(`${formatTemplate(COMMENT_MARKER_TEMPLATE, { cid: commentId })}\n`);
    }
    return 0;
}

runCli(import.meta.url, () => runCommentFromArgv(process.argv.slice(2)));
