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

import { XPATH_NS } from "./types.ts";

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
 * Serialize a DOM `Document` (or any `Node`) back to a string.
 *
 * `encoding` is accepted for parity with Python's `doc.toxml(encoding=...)`
 * and is currently advisory: `@xmldom` always emits UTF-16 JS strings and
 * leaves byte-encoding to the caller (write the result with `fs.writeFile(..., 'utf8')`
 * unless you have a reason to do otherwise). When `encoding` is supplied
 * we prepend an XML declaration that names it, like Python does.
 */
export const serializeXml = (node: Node, encoding?: string): string => {
    const xml = new XMLSerializer().serializeToString(node as unknown as XmldomNode);
    if (encoding) {
        const decl = `<?xml version="1.0" encoding="${encoding}" standalone="yes"?>`;
        return xml.startsWith('<?xml') ? xml : `${decl}\n${xml}`;
    }
    return xml;
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
        } else if (lines.length > 0 && /<[^/!?][^>]*>$/.test(lines[lines.length - 1])) {
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
 */
export const getElementsByTagNameNSAll = (root: Document | Element, namespaceURI: string, localName: string): Element[] => {
    const list = root.getElementsByTagNameNS(namespaceURI, localName);
    const out: Element[] = [];
    for (let i = 0; i < list.length; i += 1) {
        const item = list.item(i);
        if (item) {
            out.push(item);
        }
    }
    return out;
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
export const makeSelect = (extraNamespaces: Record<string, string> = {}): xpath.XPathSelect => xpath.useNamespaces({ ...XPATH_NS, ...extraNamespaces });

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
