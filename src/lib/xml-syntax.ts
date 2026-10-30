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

import type { Severity, ValidationIssue } from "./types";

/** Stable codes emitted by the built-in XML syntax checker. */
export type XmlSyntaxCode = "xml-syntax" | "xml-undeclared-prefix" | "xml-encoding";

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
export const checkXmlSyntax = async (bytes: Uint8Array, partName: string): Promise<ValidationIssue[]> => {
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
