// cspell:words Xmldom DTDLOAD NOENT
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
 * Built-in XML syntax checking with original-byte diagnostics.
 *
 * `@xmldom/xmldom` is the DOM layer (see `src/lib/xml-helpers.ts`): it is
 * lenient, recovers from some malformed input, takes a JavaScript string (so
 * a UTF-16 part decoded as UTF-8 is garbage), and reports failures as one
 * opaque `Error` with no structured position. libxml2 parses *bytes*, honours
 * the BOM and the XML-declaration encoding, and hands back one structured
 * diagnostic per defect with 1-based line/column.
 *
 * This module owns the byte-level half of that story:
 *
 *  - {@link checkXmlSyntax} parses raw part bytes with libxml2 WITHOUT
 *    `XML_PARSE_RECOVER` (no silent acceptance of malformed content) and WITH
 *    `XML_PARSE_NONET | XML_PARSE_NO_XXE` and WITHOUT `XML_PARSE_NOENT` /
 *    `XML_PARSE_DTDLOAD`, so external entities are never fetched or expanded.
 *  - {@link mapXmlDiagnostics} is the pure `diagnostic -> ValidationIssue`
 *    mapping, so the stable codes are pinned by unit tests that never touch
 *    the wasm engine.
 *
 * Codes are stable and independent of libxml2's message wording:
 *
 *  - `xml-syntax` — not well-formed.
 *  - `xml-undeclared-prefix` — an element/attribute uses an undeclared
 *    namespace prefix.
 *  - `xml-encoding` — the declared encoding is unsupported, or the bytes do
 *    not decode in the declared encoding.
 *
 * Parsing never writes: bytes go in, issues come out.
 */

import { DEFAULT_RESOURCE_CEILINGS, type ResourceCeilings } from "./resource-ceilings";
import type { Severity, ValidationIssue } from "./types";
import { decodeXmlBytes } from "./xml-helpers";

/** Stable codes emitted by the built-in XML syntax checker. */
export type XmlSyntaxCode = "xml-syntax" | "xml-undeclared-prefix" | "xml-encoding";

/** Stable codes emitted when a resource ceiling refuses a part before a DOM is built. */
export type XmlCeilingCode = "xml-attribute-ceiling" | "xml-entity-declaration-ceiling" | "xml-depth-ceiling";

/**
 * Structural view of one libxml2 diagnostic (`ErrorDetail` in
 * `libxml2-wasm/lib/libxml2.mts`). Declared structurally — rather than
 * importing the wasm type — so the mapping stays testable as a pure function
 * and the libxml2 version is free to add fields.
 */
export type XmlParseDiagnostic = {
    readonly message: string;
    /** libxml2 severity: 1 = warning, 2 = error, 3 = fatal. */
    readonly level: number;
    /** 1-based line, 0 when libxml2 has no position for the diagnostic. */
    readonly line: number;
    /** 1-based column, 0 when libxml2 has no position for the diagnostic. */
    readonly col: number;
    readonly file?: string;
    readonly xpath?: string;
};

/** Diagnostics below this level are advisories, not failures. */
const DIAGNOSTIC_ERROR_LEVEL = 2;

/** libxml2 prefixes every diagnostic message with the reporting routine. */
const UNDECLARED_PREFIX_PATTERNS: readonly RegExp[] = [
    /^Namespace prefix \S+ on \S+ is not defined/,
    /^Namespace prefix \S+ for \S+ on \S+ is not defined/,
];

const UNSUPPORTED_ENCODING_PATTERNS: readonly RegExp[] = [/^Unsupported encoding/, /^Invalid bytes in character encoding/];

/**
 * Classify one diagnostic into a stable code. Message matching is
 * deliberately narrow and anchored: these strings are libxml2's own
 * diagnostics for structural facts (a prefix that is not declared, an
 * encoding libxml2 cannot decode), not free-form prose. Everything else —
 * every well-formedness defect — is `xml-syntax`.
 */
export const classifyXmlDiagnostic = (message: string): XmlSyntaxCode => {
    const text = message.trim();
    if (UNDECLARED_PREFIX_PATTERNS.some((pattern) => pattern.test(text))) {
        return "xml-undeclared-prefix";
    }
    if (UNSUPPORTED_ENCODING_PATTERNS.some((pattern) => pattern.test(text))) {
        return "xml-encoding";
    }
    return "xml-syntax";
};

/**
 * Pure `diagnostic[] -> ValidationIssue[]` mapping.
 *
 * Every issue carries the ORIGINAL part path (`partName`), never the path
 * libxml2 echoes back in `detail.file` — the caller owns the mapping from a
 * package-relative name to a location, and the diagnostic position is the only
 * thing worth taking from the engine. Line/column are omitted (rather than
 * reported as 0) when libxml2 has no position for the diagnostic.
 */
export const mapXmlDiagnostics = (details: readonly XmlParseDiagnostic[], partName: string): ValidationIssue[] =>
    details.map((detail) => {
        const severity: Severity = detail.level >= DIAGNOSTIC_ERROR_LEVEL ? "error" : "warning";
        const issue: ValidationIssue = {
            severity,
            code: classifyXmlDiagnostic(detail.message),
            message: detail.message.trim(),
            path: partName,
        };
        return detail.line > 0 ? { ...issue, line: detail.line, ...(detail.col > 0 ? { column: detail.col } : {}) } : issue;
    });

/**
 * Parse one XML part's original bytes and report syntax diagnostics.
 *
 * Returns an empty array for well-formed input. Returns one issue per
 * libxml2 diagnostic otherwise (plus any non-fatal parse warnings, at
 * `warning` severity). Never throws and never writes: the caller keeps
 * ownership of the bytes.
 *
 * External entities are never resolved — `XML_PARSE_NO_XXE` is set and neither
 * `XML_PARSE_NOENT` nor `XML_PARSE_DTDLOAD` is, so a `<!ENTITY e SYSTEM
 * "file:///…">` stays an unresolved entity reference node.
 */
export const checkXmlSyntax = async (
    bytes: Uint8Array,
    partName: string,
    ceilings: ResourceCeilings = DEFAULT_RESOURCE_CEILINGS,
): Promise<ValidationIssue[]> => {
    // One linear pre-scan before the engine parse: an attribute, entity or
    // (caller-lowered) depth bomb is refused with its named ceiling and the
    // engine never parses the bytes. A depth breach at libxml2's own limit is
    // left to libxml2, whose "Excessive depth" refusal carries the position
    // (#1159).
    const ceilingIssues = xmlCeilingIssues(decodeXmlBytes(bytes), partName, ceilings);
    if (ceilingIssues.length > 0) {
        return ceilingIssues;
    }
    const { XmlDocument, XmlParseError, ParseOption } = await import("libxml2-wasm");
    // No XML_PARSE_RECOVER: malformed content must not be silently accepted.
    // NONET + NO_XXE, and deliberately no NOENT/DTDLOAD so external entities
    // are neither fetched nor substituted. libxml2's ParseOption is a bit-flag
    // field, so combining flags requires a bitwise OR.
    // oxlint-disable-next-line no-bitwise
    const option = ParseOption.XML_PARSE_NONET | ParseOption.XML_PARSE_NO_XXE;
    let doc;
    try {
        doc = XmlDocument.fromBuffer(bytes, { option, url: partName });
    } catch (err) {
        if (err instanceof XmlParseError && err.details.length > 0) {
            return mapXmlDiagnostics(err.details, partName);
        }
        // No structured detail (e.g. a non-Error throw): still a hard failure,
        // still attributed to the original part.
        const message = err instanceof Error ? err.message : String(err);
        return [{ severity: "error", code: "xml-syntax", message, path: partName }];
    }
    const warnings = mapXmlDiagnostics(doc.warnings, partName);
    doc.dispose();
    return warnings;
};

/**
 * libxml2's element-depth limit without `XML_PARSE_HUGE`: `checkXmlSyntax`
 * accepts a document nested 256 elements deep, the root counting as one, and
 * refuses 257 with "Excessive depth in document" (#1159).
 */
export const XML_MAX_ELEMENT_DEPTH = 256;

/** The index just past `terminator` at or after `from`, or -1 when the text ends first. */
const indexPast = (text: string, terminator: string, from: number): number => {
    const at = text.indexOf(terminator, from);
    return at === -1 ? -1 : at + terminator.length;
};

/**
 * The index just past the `>` that ends the tag or declaration starting
 * before `from`, reading quoted values and a declaration's `[…]` internal
 * subset as opaque, or -1 when the text ends first.
 */
const indexPastTag = (text: string, from: number): number => {
    let quote = "";
    let brackets = 0;
    for (let i = from; i < text.length; i += 1) {
        const c = text[i];
        if (quote !== "") {
            if (c === quote) {
                quote = "";
            }
        } else if (brackets > 0 && text.startsWith("<!--", i)) {
            // A comment in the internal subset is opaque: its brackets and
            // quotes do not nest the declaration.
            const end = text.indexOf("-->", i + 4);
            if (end === -1) {
                return -1;
            }
            i = end + 2;
        } else if (c === '"' || c === "'") {
            quote = c;
        } else if (c === "[") {
            brackets += 1;
        } else if (c === "]") {
            brackets -= 1;
        } else if (c === ">" && brackets <= 0) {
            return i + 1;
        }
    }
    return -1;
};

/**
 * Whether `text` nests elements deeper than `limit`, read in one linear pass
 * without building a DOM (#1159). Comments, CDATA sections, processing
 * instructions, declarations and quoted attribute values are skipped, and an
 * empty-element tag counts at the depth it would open. Malformed markup ends
 * the scan with `false`: `checkXmlSyntax` reports it, and this reading is only
 * for keeping depth-quadratic work away from a part libxml2 refuses.
 */
export const exceedsXmlElementDepth = (text: string, limit: number = XML_MAX_ELEMENT_DEPTH): boolean => {
    let depth = 0;
    let i = text.indexOf("<");
    while (i !== -1) {
        let next: number;
        if (text.startsWith("<!--", i)) {
            next = indexPast(text, "-->", i + 4);
        } else if (text.startsWith("<![CDATA[", i)) {
            next = indexPast(text, "]]>", i + 9);
        } else if (text.startsWith("<?", i)) {
            next = indexPast(text, "?>", i + 2);
        } else if (text.startsWith("<!", i)) {
            next = indexPastTag(text, i + 2);
        } else if (text.startsWith("</", i)) {
            depth -= 1;
            next = indexPast(text, ">", i + 2);
        } else {
            next = indexPastTag(text, i + 1);
            if (next === -1) {
                return false;
            }
            depth += 1;
            if (depth > limit) {
                return true;
            }
            if (text[next - 2] === "/") {
                depth -= 1;
            }
        }
        if (next === -1) {
            return false;
        }
        i = text.indexOf("<", next);
    }
    return false;
};

/** Which resource ceiling a pre-scan breach names. */
export type XmlCeilingKind = "depth" | "attributes" | "entity-declarations";

/** One pre-scan breach: the ceiling it broke, the observed count, and where in the text. */
export type XmlCeilingBreach = {
    readonly kind: XmlCeilingKind;
    readonly observed: number;
    readonly limit: number;
    /** Index in `text` of the tag or declaration that crossed the ceiling. */
    readonly index: number;
};

/**
 * One linear pass over `text` that reads element depth, attributes per start
 * tag, and `<!ENTITY` declarations against the ceilings — the
 * {@link exceedsXmlElementDepth} scan generalized — so a hostile part is
 * refused before a DOM is built over it. Comments, CDATA sections, processing
 * instructions, declarations and quoted attribute values are skipped, an
 * empty-element tag counts at the depth it would open, and attributes are
 * counted as quoted values inside the start tag (one per attribute in
 * well-formed markup). Malformed markup ends the scan: `checkXmlSyntax`
 * reports it, and this reading only keeps hostile shapes away from parsers.
 * At most one breach per kind is reported.
 */
export const scanXmlCeilings = (
    text: string,
    ceilings: Pick<ResourceCeilings, "maxDepth" | "maxAttributes" | "maxEntityDeclarations">,
): XmlCeilingBreach[] => {
    const breaches: XmlCeilingBreach[] = [];
    const found = new Set<XmlCeilingKind>();
    let depth = 0;
    let entityDeclarations = 0;
    let i = text.indexOf("<");
    while (i !== -1 && found.size < 3) {
        let next: number;
        if (text.startsWith("<!--", i)) {
            next = indexPast(text, "-->", i + 4);
        } else if (text.startsWith("<![CDATA[", i)) {
            next = indexPast(text, "]]>", i + 9);
        } else if (text.startsWith("<?", i)) {
            next = indexPast(text, "?>", i + 2);
        } else if (text.startsWith("<!ENTITY", i)) {
            entityDeclarations += 1;
            if (entityDeclarations > ceilings.maxEntityDeclarations && !found.has("entity-declarations")) {
                breaches.push({
                    kind: "entity-declarations",
                    observed: entityDeclarations,
                    limit: ceilings.maxEntityDeclarations,
                    index: i,
                });
                found.add("entity-declarations");
            }
            next = indexPastTag(text, i + 2);
        } else if (text.startsWith("<!DOCTYPE", i)) {
            next = indexPastTag(text, i + 2);
            // The declaration's internal subset `[...]` is where the ENTITY
            // declarations live; indexPastTag skips it opaquely, so count
            // them inside the span here.
            if (next !== -1) {
                // Searched within the declaration's own span only, so a run
                // of declarations stays linear (r4651164156).
                const span = text.slice(i, next);
                const open = span.indexOf("[");
                if (open !== -1) {
                    let at = span.indexOf("<!ENTITY", open);
                    while (at !== -1) {
                        const j = i + at;
                        entityDeclarations += 1;
                        if (entityDeclarations > ceilings.maxEntityDeclarations && !found.has("entity-declarations")) {
                            breaches.push({
                                kind: "entity-declarations",
                                observed: entityDeclarations,
                                limit: ceilings.maxEntityDeclarations,
                                index: j,
                            });
                            found.add("entity-declarations");
                        }
                        at = span.indexOf("<!ENTITY", at + 8);
                    }
                }
            }
        } else if (text.startsWith("<!", i)) {
            next = indexPastTag(text, i + 2);
        } else if (text.startsWith("</", i)) {
            depth -= 1;
            next = indexPast(text, ">", i + 2);
        } else {
            // A `<` that opens no name is malformed text, not a tag: end the
            // scan (checkXmlSyntax reports it) rather than count it as depth.
            if (!/[A-Za-z_:\u00C0-\uFFFF]/.test(text[i + 1] ?? "")) {
                break;
            }
            next = indexPastTag(text, i + 1);
            if (next === -1) {
                break;
            }
            depth += 1;
            if (depth > ceilings.maxDepth && !found.has("depth")) {
                breaches.push({ kind: "depth", observed: depth, limit: ceilings.maxDepth, index: i });
                found.add("depth");
            }
            let attributes = 0;
            let quote = "";
            for (let j = i + 1; j < next; j += 1) {
                const c = text[j]!;
                if (quote !== "") {
                    if (c === quote) {
                        quote = "";
                        attributes += 1;
                    }
                } else if (c === '"' || c === "'") {
                    quote = c;
                }
            }
            if (attributes > ceilings.maxAttributes && !found.has("attributes")) {
                breaches.push({ kind: "attributes", observed: attributes, limit: ceilings.maxAttributes, index: i });
                found.add("attributes");
            }
            if (text[next - 2] === "/") {
                depth -= 1;
            }
        }
        if (next === -1) {
            break;
        }
        i = text.indexOf("<", next);
    }
    return breaches;
};

const CEILING_CODES: Record<XmlCeilingKind, XmlCeilingCode> = {
    attributes: "xml-attribute-ceiling",
    "entity-declarations": "xml-entity-declaration-ceiling",
    depth: "xml-depth-ceiling",
};

const CEILING_LABELS: Record<XmlCeilingKind, { noun: string; option: string }> = {
    attributes: { noun: "attribute", option: "maxAttributes" },
    "entity-declarations": { noun: "entity-declaration", option: "maxEntityDeclarations" },
    depth: { noun: "element-depth", option: "maxDepth" },
};

/**
 * The error findings for the pre-scan breaches a part's own engine parse does
 * not already report. Attribute and entity-declaration breaches always
 * become findings; a depth breach does only when the caller set `maxDepth`
 * below libxml2's own limit of {@link XML_MAX_ELEMENT_DEPTH}, since the
 * engine refuses deeper input on its own (#1159) and would otherwise accept
 * what the lowered ceiling refuses. Each names the ceiling it enforced.
 */
export const xmlCeilingIssues = (text: string, partName: string, ceilings: ResourceCeilings): ValidationIssue[] =>
    scanXmlCeilings(text, ceilings)
        .filter((breach) => breach.kind !== "depth" || ceilings.maxDepth < XML_MAX_ELEMENT_DEPTH)
        .map((breach) => ({
            severity: "error" as const,
            code: CEILING_CODES[breach.kind],
            message: `Part ${JSON.stringify(partName)} crosses the ${CEILING_LABELS[breach.kind].noun} ceiling: ${breach.observed} > ${breach.limit} (${CEILING_LABELS[breach.kind].option}). Refused before a DOM is built over it.`,
            path: partName,
        }));
