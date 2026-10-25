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
 * Shared XML utilities. Every other module imports from here so the
 * `@xmldom/xmldom` + `xpath` dependency stays swappable.
 *
 * Convention for porters: do not call `DOMParser` / `XMLSerializer` directly.
 * Use `parseXml` / `serializeXml` / `prettyXml` here.
 */

import { DOMParser, XMLSerializer, type Node as XmldomNode } from "@xmldom/xmldom";
import * as xpath from "xpath";

import { NS, XPATH_NS } from "./types";

/**
 * Parse an XML string into a DOM `Document`. Mirrors the Python
 * `defusedxml.minidom.parseString(text)` shape.
 *
 * Throws on malformed input. The error message is preserved (xmldom emits
 * detailed line/column info in the message body).
 */
export const parseXml = (text: string): Document => {
    const errors: string[] = [];
    const parser = new DOMParser({
        onError: (level, msg) => {
            if (level === "error" || level === "fatalError") {
                errors.push(msg);
            }
        },
    });
    // XML 1.0 §2.8 permits a UTF-8 BOM (U+FEFF) at the start of an entity;
    // @xmldom does not strip it and surfaces it as a "processing instruction
    // at position 1" error. Strip defensively so callers don't see spurious
    // failures on files produced by Microsoft Office (which routinely emits
    // BOM-prefixed parts). The strict profile detects BOMs at the byte level
    // before they reach the parser, so this strip never silently masks an
    // intentional check.
    const stripped = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
    const doc = parser.parseFromString(stripped, "text/xml") as unknown as Document;
    if (errors.length > 0) {
        throw new Error(`parseXml failed: ${errors.join("; ")}`);
    }
    return doc;
};

/**
 * Byte-level decoding for the DOM layer. `@xmldom` takes a JavaScript string,
 * so a part stored as UTF-16 must be decoded from its ORIGINAL BYTES first —
 * reading it as UTF-8 (`fs.readFile(file, "utf-8")`) yields NUL-interleaved
 * garbage that no parser can rescue.
 *
 * Encoding selection follows XML 1.0 §4.3.3: the byte order mark wins, then
 * the `encoding` pseudo-attribute of the XML declaration. Only the three
 * encodings OOXML parts actually use are supported — UTF-8, UTF-16LE,
 * UTF-16BE — and anything else falls back to UTF-8 rather than guessing.
 * The result carries no BOM and no leading whitespace, so callers can hand it
 * straight to {@link parseXml}.
 */
export const decodeXmlBytes = (bytes: Uint8Array): string => {
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
        // UTF-8 BOM. `TextDecoder` drops it for "utf-8" only when told not to
        // be fatal about it, so decode the remainder explicitly.
        return new TextDecoder("utf-8").decode(bytes.subarray(3));
    }
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
        return new TextDecoder("utf-16le").decode(bytes.subarray(2));
    }
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
        return new TextDecoder("utf-16be").decode(bytes.subarray(2));
    }
    // No BOM. A 16-bit entity puts a NUL beside its first (ASCII) character,
    // so the byte order shows in the first two bytes (XML 1.0 Appendix F).
    // Read the declaration with that decoder, in full: it may carry any
    // amount of whitespace, so no fixed byte window is safe. Otherwise the
    // declaration is ASCII and the first 200 bytes are plenty to name UTF-16.
    const sixteenBit = bytes.length >= 2 && (bytes[0] === 0) !== (bytes[1] === 0);
    const head = sixteenBit
        ? new TextDecoder(bytes[0] === 0 ? "utf-16be" : "utf-16le").decode(bytes)
        : new TextDecoder("latin1").decode(bytes.subarray(0, Math.min(bytes.length, 200)));
    const declared = /^\s*<\?xml[^>]*encoding\s*=\s*["']([^"']+)["']/i
        .exec(head)?.[1]
        ?.toLowerCase()
        .replace(/[^a-z0-9]/g, "");
    if (declared === "utf16" || declared === "utf16le" || declared === "utf16be") {
        // XML 1.0 §4.3.3 requires a BOM for UTF-16 entities; without one, the
        // first character discriminates the endianness ('<' as 3C 00 or 00 3C).
        const littleEndian = declared === "utf16le" || (declared === "utf16" && !(bytes[0] === 0 && bytes[1] === 0x3c));
        return new TextDecoder(littleEndian ? "utf-16le" : "utf-16be").decode(bytes);
    }
    return new TextDecoder("utf-8").decode(bytes);
};

/**
 * Parse ORIGINAL XML part bytes into a DOM `Document`.
 *
 * The decoding step lives in {@link decodeXmlBytes} and the DOM construction
 * stays here, so no caller ever calls `DOMParser` directly.
 */
export const parseXmlBytes = (bytes: Uint8Array): Document => parseXml(decodeXmlBytes(bytes));

/**
 * Serialize a DOM `Document` (or any `Node`) back to a string.
 *
 * `encoding` is accepted for parity with Python's `doc.toxml(encoding=...)`
 * and is currently advisory: `@xmldom` always emits UTF-16 JS strings and
 * leaves byte-encoding to the caller (write the result with `fs.writeFile(..., 'utf8')`
 * unless you have a reason to do otherwise). When `encoding` is supplied
 * we prepend an XML declaration that names it, like Python does — and when
 * the document already carries a declaration, its encoding pseudo-attribute
 * is rewritten to name the requested encoding, since the returned string is
 * what the caller will store or re-parse as that encoding.
 */
export const serializeXml = (node: Node, encoding?: string): string => {
    const xml = new XMLSerializer().serializeToString(node as unknown as XmldomNode);
    if (!encoding) {
        return xml;
    }
    // `<?xml-stylesheet?>` and friends are PIs, not the XML declaration.
    const declMatch = /^<\?xml[\s?]/.exec(xml);
    if (!declMatch) {
        return `<?xml version="1.0" encoding="${encoding}" standalone="yes"?>\n${xml}`;
    }
    // The serialized string will be stored or re-parsed as `encoding`, so a
    // declaration the source document carried must not keep naming a
    // different one: a retained encoding="UTF-16" on a JS string (handed to
    // libxml2 as UTF-8 bytes, or written to disk as UTF-8) makes the next
    // reader mis-decode every byte (#290). Rewrite the declaration's encoding
    // pseudo-attribute to the requested encoding; a declaration without one
    // stays as it is (absent means UTF-8/UTF-16-by-BOM detection anyway).
    return xml.replace(/^(<\?xml\b(?:(?!\?>)[\s\S])*?encoding\s*=\s*)(["'])[^"']*\2/i, `$1$2${encoding}$2`);
};

// Match the former /<[^/!?][^>]*>$/ suffix check in one backward scan.
// Retrying the regex at each '<' made comments with repeated '< ' quadratic.
const hasOpeningTagSuffix = (line: string): boolean => {
    const end = line.length - 1;
    if (line[end] !== ">") {
        return false;
    }
    for (let i = end - 1; i >= 0; i -= 1) {
        // An earlier '>' can only be consumed as the tag's second character.
        if (line[i] === ">") {
            return i > 0 && line[i - 1] === "<";
        }
        if (i < end - 1 && line[i] === "<" && !"/!?".includes(line[i + 1])) {
            return true;
        }
    }
    return false;
};

/**
 * Pretty-print a DOM `Document` with the given indent (default two spaces).
 * Mirrors Python's `doc.toprettyxml(indent="  ")`.
 *
 * Implementation note: `@xmldom` does not have a built-in pretty-printer, so
 * we walk the serialized output and re-indent. This is good enough for
 * human-readable diffs; do NOT use the result as the canonical form for
 * roundtrip comparison (whitespace inside `xml:space="preserve"` runs
 * survives but surrounding indentation is reflowed).
 */
export const prettyXml = (node: Node, indent = "  "): string => {
    const raw = serializeXml(node);
    // Match either: an XML comment (<!-- ... -->, may contain '>'), any other
    // tag/decl/PI, or a run of text. The comment branch must come first so
    // comments containing '>' don't terminate the [^>]+ branch prematurely.
    const tokens = raw.match(/<(?:!--[\s\S]*?--|[^>]+)>|[^<]+/g) ?? [];
    let depth = 0;
    const lines: string[] = [];
    for (const token of tokens) {
        if (token.startsWith("<?") || token.startsWith("<!")) {
            lines.push(indent.repeat(depth) + token);
            continue;
        }
        if (token.startsWith("</")) {
            depth = Math.max(0, depth - 1);
            lines.push(indent.repeat(depth) + token);
            continue;
        }
        if (token.startsWith("<") && !token.endsWith("/>")) {
            lines.push(indent.repeat(depth) + token);
            depth += 1;
            continue;
        }
        if (token.startsWith("<") && token.endsWith("/>")) {
            lines.push(indent.repeat(depth) + token);
            continue;
        }
        const trimmed = token.replace(/\s+/g, " ").trim();
        if (trimmed === "") {
            continue;
        }
        if (lines.length > 0 && !lines[lines.length - 1].endsWith(">")) {
            lines[lines.length - 1] += trimmed;
        } else if (lines.length > 0 && hasOpeningTagSuffix(lines[lines.length - 1])) {
            lines[lines.length - 1] += trimmed;
        } else {
            lines.push(indent.repeat(depth) + trimmed);
        }
    }
    return lines.join("\n");
};

/**
 * One canonical helper matching the convention used throughout the Python
 * source where `getElementsByTagName('w:p')` is called with a *prefixed*
 * name. The Python code relies on the parser keeping the prefix attached;
 * `@xmldom` follows the DOM spec and exposes nodes via namespace URI +
 * localName instead.
 *
 * Pass the namespace URI (use `NS` from `./types`) and the local name; you
 * get a flat array, not a live `NodeList`.
 *
 * Use `"*"` as the namespace to match across all namespaces (mirrors
 * `Document.getElementsByTagName('*')`).
 *
 * A query rooted at a Document reads an index of its elements once a second
 * query finds the document unchanged since the first: a full DOCX run asks
 * the same main document for one element name after another, hundreds of
 * walks that dominated a 2 MB document's validation time. A caller that
 * changes the document between queries keeps the plain walk instead of
 * rebuilding the index for each one.
 */
export const getElementsByTagNameNSAll = (root: Document | Element, namespaceURI: string, localName: string): Element[] => {
    if (root.nodeType === 9 /* DOCUMENT_NODE */) {
        const index = documentElementIndex(root as Document);
        if (index !== undefined) {
            if (namespaceURI === "*" && localName === "*") return index.all.slice();
            if (namespaceURI !== "*" && localName !== "*") return index.byName.get(namespaceURI)?.get(localName)?.slice() ?? [];
            return index.all.filter(
                (elem) =>
                    (namespaceURI === "*" || elem.namespaceURI === namespaceURI) && (localName === "*" || elem.localName === localName),
            );
        }
    }
    return walkElements(root, namespaceURI, localName);
};

/** Every descendant element of a document in document order, and the same grouped by namespace and local name. */
interface DocumentElementIndex {
    /** The document's mutation counter when the index was built. */
    readonly inc: number;
    readonly all: Element[];
    readonly byName: Map<string | null, Map<string, Element[]>>;
}

const documentElementIndexes = new WeakMap<Document, DocumentElementIndex>();

/** The mutation counter each document had at its last query that walked instead of reading an index. */
const documentWalkedAt = new WeakMap<Document, number>();

/**
 * The element index of `document`, current as of its last change, or
 * undefined when the caller should walk. @xmldom increments the private
 * `_inc` counter on every child or attribute change of a node the document
 * owns, which is what its own live NodeLists key their caches on; without
 * that counter there is no index and callers walk. The index is built on the
 * second query that sees the same counter, so a caller that changes the
 * document between queries pays one walk per query, as without an index.
 */
const documentElementIndex = (document: Document): DocumentElementIndex | undefined => {
    const inc = (document as Document & { _inc?: unknown })._inc;
    if (typeof inc !== "number") return undefined;
    const cached = documentElementIndexes.get(document);
    if (cached?.inc === inc) return cached;
    if (documentWalkedAt.get(document) !== inc) {
        documentWalkedAt.set(document, inc);
        return undefined;
    }
    const all = walkElements(document, "*", "*");
    const byName = new Map<string | null, Map<string, Element[]>>();
    for (const elem of all) {
        let byLocal = byName.get(elem.namespaceURI);
        if (byLocal === undefined) {
            byLocal = new Map();
            byName.set(elem.namespaceURI, byLocal);
        }
        const name = elem.localName ?? "";
        const list = byLocal.get(name);
        if (list === undefined) byLocal.set(name, [elem]);
        else list.push(elem);
    }
    const index = { inc, all, byName };
    documentElementIndexes.set(document, index);
    return index;
};

/** The stack-based DFS behind {@link getElementsByTagNameNSAll}: matching descendants of `root` in document order. */
const walkElements = (root: Document | Element, namespaceURI: string, localName: string): Element[] => {
    const out: Element[] = [];
    const stack: Node[] = [root];

    while (stack.length > 0) {
        const node = stack.pop()!;
        if (node !== root && node.nodeType === 1) {
            const elem = node as Element;
            if ((namespaceURI === "*" || elem.namespaceURI === namespaceURI) && (localName === "*" || elem.localName === localName)) {
                out.push(elem);
            }
        }

        let child = node.lastChild;
        while (child) {
            if (child.nodeType === 1) {
                stack.push(child);
            }
            child = child.previousSibling;
        }
    }
    return out;
};

/**
 * Stack-based DFS replacement for `getElementsByTagName('*')`. `@xmldom`'s
 * live `NodeList` re-traverses the tree after every DOM change, so a loop
 * that edits while it iterates is quadratic; this returns a flat array of
 * every descendant element (root excluded) in document order.
 *
 * Thin wrapper over `getElementsByTagNameNSAll(root, "*", "*")` so the two
 * entry points share one traversal implementation — the wildcard form and
 * the NS-filtered form must never drift apart.
 */
export const getElementsByTagNameAll = (root: Document | Element): Element[] => getElementsByTagNameNSAll(root, "*", "*");

/** OPC relationships are direct children of the package Relationships root (#434). */
export const getRelationshipChildren = (document: Document): Element[] => {
    const root = document.documentElement;
    if (!root || root.localName !== "Relationships" || root.namespaceURI !== NS.PR) {
        return [];
    }
    const relationships: Element[] = [];
    for (let child = root.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) {
            continue;
        }
        const element = child as Element;
        if (element.localName === "Relationship" && element.namespaceURI === NS.PR) {
            relationships.push(element);
        }
    }
    return relationships;
};

/**
 * Build a namespace-aware xpath selector pre-bound to the OOXML prefixes in
 * `XPATH_NS`. Use the returned function exactly like the `xpath.select`
 * default export — it just spares every caller from re-declaring the
 * namespace map.
 *
 * Example:
 *   const $$ = makeSelect();
 *   const paragraphs = $$("//w:p", doc) as Element[];
 */
export const makeSelect = (extraNamespaces: Record<string, string> = {}): xpath.XPathSelect =>
    xpath.useNamespaces({ ...XPATH_NS, ...extraNamespaces });

/**
 * 1-based source line number for a parsed node, or `0` if the parser did not
 * record one. `@xmldom` exposes `lineNumber` directly on Element/Attr nodes
 * (zero-based-internally, but exposed 1-based) when the document was parsed
 * from a string — matches Python's `elem.sourceline`.
 *
 * Use this everywhere instead of poking `(node as any).lineNumber` so callers
 * cope cleanly with constructed-in-memory nodes (which have no source).
 */
export const getLineNumber = (node: Node | null | undefined): number => {
    if (!node) {
        return 0;
    }
    const value = (node as unknown as { lineNumber?: number }).lineNumber;
    return typeof value === "number" ? value : 0;
};

/**
 * Local name of a DOM element (the part after the `:` prefix). For nodes
 * built by `@xmldom` this is `localName`; the helper exists so call sites
 * read self-documenting and so we have one place to fall back to `nodeName`
 * if a future parser swap returns `null`.
 *
 * Returns lower-cased name when `lowercase` is true (matches the Python
 * source's `tag.split('}')[-1].lower()` idiom used in unique-ID checks).
 */
export const getLocalName = (elem: Element, lowercase = false): string => {
    const name = elem.localName ?? elem.nodeName ?? "";
    return lowercase ? name.toLowerCase() : name;
};
