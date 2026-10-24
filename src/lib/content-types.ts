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
 * OPC content-type resolution (ECMA-376 Part 2 §10.1.2.4).
 *
 * A part's content type is resolved by an exact `<Override PartName>` match
 * first, then by a `<Default Extension>` matching the part name's extension.
 * PartName comparison canonicalises both sides the same way (see
 * {@link canonicalPartName}): percent-escapes of unreserved characters
 * decode, other escapes stay, characters outside `pchar`/`/` encode, the
 * escape hex is uppercased, and the result compares ASCII
 * case-insensitively — so `%40` and `@` remain distinct part names while
 * `my%20img` matches `my img`. Duplicate declarations keep their first entry
 * deterministically. Pure functions only — no filesystem access; callers
 * hand in the parsed [Content_Types].xml DOM.
 */

import path from "node:path";

import { getElementsByTagNameNSAll } from "./xml-helpers";
import type { ValidationIssue } from "./types";

export const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
export const CT_DUPLICATE_DECLARATION = "ct-duplicate-declaration";

/**
 * The content-types stream is a ZIP item, not an OPC part. Match its ASCII
 * spelling without percent-decoding, as System.IO.Packaging does (#396).
 */
export const isContentTypesStreamName = (name: string): boolean =>
    /^\[content_types\]\.xml$/i.test(name.replace(/\\/g, "/").replace(/^\//, ""));

/**
 * The stored content-types stream: `[Content_Types].xml` when it is stored,
 * otherwise the first name {@link isContentTypesStreamName} accepts. Of two
 * stored streams, System.IO.Packaging keeps the last matching ZIP item in
 * central-directory order; this follows the rule the package view applies to
 * equivalent entries (#366), which a directory input, having no ZIP order,
 * can apply too. The opc-equivalent-part-names check reports duplicates (#351).
 */
export const findContentTypesStream = (names: Iterable<string>): string | undefined => {
    let found: string | undefined;
    for (const name of names) {
        if (name === "[Content_Types].xml") {
            return name;
        }
        if (found === undefined && isContentTypesStreamName(name)) {
            found = name;
        }
    }
    return found;
};

/** Normalized declaration table extracted from a [Content_Types].xml DOM. */
export type ContentTypeTable = {
    /** Normalized PartName → ContentType for every `<Override>`. */
    readonly overrides: ReadonlyMap<string, string>;
    /** ASCII-lowercased Extension → ContentType for every `<Default>`. */
    readonly defaults: ReadonlyMap<string, string>;
};

/** Result of resolving one part name against a {@link ContentTypeTable}. */
export type ResolvedContentType = {
    readonly contentType: string;
    readonly source: "override" | "default";
};

/**
 * ASCII-only lowercase. OPC part-name equivalence and `Default@Extension`
 * matching are ASCII case-insensitive; full Unicode `toLowerCase()` would
 * fold non-ASCII letters (e.g. `É` → `é`) that OPC compares verbatim.
 */
export const asciiLowercase = (value: string): string => {
    let out = "";
    for (let i = 0; i < value.length; i += 1) {
        const code = value.charCodeAt(i);
        out += code >= 0x41 && code <= 0x5a ? String.fromCharCode(code + 0x20) : value[i];
    }
    return out;
};

/**
 * Decode percent-escapes safely. Malformed sequences (e.g. `%zz`) keep
 * the raw text — the lookup then simply misses and the caller's residual
 * pass reports the part, instead of a URIError taking down validation.
 */
const decodePercentEscapes = (text: string): string => {
    try {
        return decodeURIComponent(text);
    } catch {
        return text;
    }
};

/** RFC 3986 unreserved characters: ALPHA / DIGIT / "-" / "." / "_" / "~". */
const isUnreservedAscii = (code: number): boolean =>
    (code >= 0x41 && code <= 0x5a) ||
    (code >= 0x61 && code <= 0x7a) ||
    (code >= 0x30 && code <= 0x39) ||
    code === 0x2d ||
    code === 0x2e ||
    code === 0x5f ||
    code === 0x7e;

/** OPC [M1.8]: stored name segments must not percent-encode RFC 3986 unreserved characters. */
export const hasEscapedUnreservedCharacter = (name: string): boolean =>
    Array.from(name.matchAll(/%([0-9a-f]{2})/gi)).some((match) => isUnreservedAscii(Number.parseInt(match[1]!, 16)));

/**
 * OPC [M1.7]: a part-name segment must not percent-encode `/` or `\`. The
 * escape is three literal characters in the stored spelling, so a contiguous
 * match cannot fire inside `%252F` (an encoded `%` before `2F`).
 */
export const hasEscapedSeparatorCharacter = (name: string): boolean => /%2f|%5c/i.test(name);

/**
 * `pchar` (RFC 3986 §3.3: unreserved / pct-encoded / sub-delims / ":" / "@")
 * plus the "/" separator of the hierarchical part name.
 */
const isPcharOrSlash = (code: number): boolean =>
    isUnreservedAscii(code) ||
    code === 0x21 /* ! */ ||
    code === 0x24 /* $ */ ||
    code === 0x26 /* & */ ||
    code === 0x27 /* ' */ ||
    code === 0x28 /* ( */ ||
    code === 0x29 /* ) */ ||
    code === 0x2a /* * */ ||
    code === 0x2b /* + */ ||
    code === 0x2c /* , */ ||
    code === 0x2f /* / */ ||
    code === 0x3a /* : */ ||
    code === 0x3b /* ; */ ||
    code === 0x3d /* = */ ||
    code === 0x40; /* @ */

/**
 * RFC 3987 §2.2 `ucschar`:
 *
 *     ucschar = %xA0-D7FF / %xF900-FDCF / %xFDF0-FFEF
 *               / %x10000-1FFFD / %x20000-2FFFD / %x30000-3FFFD / %x40000-4FFFD
 *               / %x50000-5FFFD / %x60000-6FFFD / %x70000-7FFFD / %x80000-8FFFD
 *               / %x90000-9FFFD / %xA0000-AFFFD / %xB0000-BFFFD / %xC0000-CFFFD
 *               / %xD0000-DFFFD / %xE1000-EFFFD
 *
 * The last block starts at E1000, so `U+E0000..U+E0FFF` is no `ucschar`
 * (#522). C1 controls, the surrogate blocks, the noncharacters and the BMP
 * private-use area stay out.
 */
const isUcschar = (codePoint: number): boolean =>
    (codePoint >= 0xa0 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfdcf) ||
    (codePoint >= 0xfdf0 && codePoint <= 0xffef) ||
    // Each astral block is n0000–nFFFD for n = 1…D: within 0x10000–0xDFFFD,
    // that is exactly a low half of at most 0xFFFD. The last block is
    // E1000–EFFFD, so it starts 0x1000 code points higher than the others.
    (codePoint >= 0x10000 && codePoint <= 0xdfffd && codePoint % 0x10000 <= 0xfffd) ||
    (codePoint >= 0xe1000 && codePoint <= 0xefffd);

/** RFC 3987 `iunreserved`: ALPHA / DIGIT / "-" / "." / "_" / "~" / ucschar. */
const isIunreserved = (codePoint: number): boolean => isUnreservedAscii(codePoint) || isUcschar(codePoint);

const UPPER_HEX = "0123456789ABCDEF";
const encodeOctet = (byte: number): string => `%${UPPER_HEX[Math.floor(byte / 16)]}${UPPER_HEX[byte % 16]}`;

/** Hex digit value, or -1 for anything that is not one. */
const hexValue = (code: number): number => {
    if (code >= 0x30 && code <= 0x39) {
        return code - 0x30;
    }
    if (code >= 0x41 && code <= 0x46) {
        return code - 0x37;
    }
    if (code >= 0x61 && code <= 0x66) {
        return code - 0x57;
    }
    return -1;
};

const utf8 = new TextEncoder();

/**
 * Canonical comparison form of a part name (issue #297): both an Override
 * PartName and a package part path reduce to the same key, so two distinct
 * names can never collapse into one. Escapes of unreserved characters
 * decode (`%7e` and `~` are the same name); escapes of everything else stay
 * escaped (`%40` and `@` stay distinct, because RFC 3986 only equates an
 * escaped character with its literal when the character is unreserved);
 * characters outside `pchar`/`/` — spaces, a lone `%`, non-ASCII — encode as
 * UTF-8 percent-escapes; the escape hex is uppercased; and the whole key
 * compares ASCII case-insensitively, as OPC part-name equivalence demands.
 */
export const canonicalPartName = (partName: string): string => {
    const text = partName.replace(/^\/+/, "");
    let out = "";
    let i = 0;
    while (i < text.length) {
        const codePoint = text.codePointAt(i)!;
        if (codePoint === 0x25 /* % */) {
            const high = hexValue(text.charCodeAt(i + 1));
            const low = hexValue(text.charCodeAt(i + 2));
            if (high >= 0 && low >= 0) {
                const octet = high * 16 + low;
                out += isUnreservedAscii(octet) ? String.fromCharCode(octet) : encodeOctet(octet);
                i += 3;
                continue;
            }
            out += "%25"; // a % not followed by two hex digits is data, not an escape
            i += 1;
            continue;
        }
        const ch = String.fromCodePoint(codePoint);
        if (isPcharOrSlash(codePoint)) {
            out += ch;
        } else {
            for (const byte of utf8.encode(ch)) {
                out += encodeOctet(byte);
            }
        }
        i += ch.length;
    }
    return asciiLowercase(out);
};

/**
 * The entry in `names` that `partName` names under OPC part-name
 * equivalence ({@link canonicalPartName}), preferring an exact match and
 * otherwise the first equivalent entry in code-unit order, so the choice
 * does not depend on how `names` is ordered; `undefined` when no entry is
 * equivalent (#344, #349).
 */
export const findEquivalentPart = (names: readonly string[], partName: string): string | undefined => {
    if (names.includes(partName)) {
        return partName;
    }
    const key = canonicalPartName(partName);
    let found: string | undefined;
    for (const name of names) {
        if (canonicalPartName(name) === key && (found === undefined || name < found)) {
            found = name;
        }
    }
    return found;
};

/**
 * A lookup that makes {@link findEquivalentPart}'s choice over `names`, built
 * once, so each lookup costs one canonicalization instead of a scan of every
 * entry (#358).
 */
export const equivalentPartLookup = (names: readonly string[]): ((partName: string) => string | undefined) => {
    const exact = new Set(names);
    const byKey = new Map<string, string>();
    for (const name of names) {
        const key = canonicalPartName(name);
        const found = byKey.get(key);
        if (found === undefined || name < found) {
            byKey.set(key, name);
        }
    }
    return (partName) => (exact.has(partName) ? partName : byKey.get(canonicalPartName(partName)));
};

/**
 * The entry in `names` that a package-relative part reference names: a query
 * or fragment is dropped, since only the path names a part (System.IO.Packaging's
 * `ResolvePartUri` keeps only the path, #361), `\` reads as `/`, a leading `/`
 * is dropped and dot segments are resolved, then
 * the OPC-equivalent entry is taken ({@link findEquivalentPart}). Falls back
 * to the normalized name when no entry is equivalent, and is `undefined` for
 * a reference that is empty, names the package root or climbs above it
 * (#356). Every resolver of the main document part goes through here, so
 * they agree on which entry it is (#349).
 */
export const resolvePartName = (names: readonly string[], partName: string): string | undefined => {
    const [reference = ""] = partName.split(/[?#]/, 1);
    const normalized = path.posix.normalize(reference.replace(/\\/g, "/").replace(/^\/+/, ""));
    if (normalized === "." || normalized === ".." || normalized.startsWith("../")) {
        return undefined;
    }
    return findEquivalentPart(names, normalized) ?? normalized;
};

/**
 * The fully percent-decoded form of a part name: leading `/` separators
 * stripped (after decoding, so an encoded leading `%2F` strips too) and
 * ASCII-lowercased. This is the decoded display form, not the comparison
 * key — Override lookups use {@link canonicalPartName}, which keeps escapes
 * of reserved characters distinct from their literals (issue #297).
 */
export const normalizePartName = (partName: string): string => asciiLowercase(decodePercentEscapes(partName).replace(/^\/+/, ""));

/**
 * Whether a part is a relationships part: its OPC-canonical name
 * ({@link canonicalPartName}: ASCII case folded, escapes of unreserved
 * characters decoded) ends in `.rels`, so `_RELS/.RELS` and
 * `word/_rels/document.xml.%72els` are relationships parts. Every
 * relationships pass and the XSD walk test parts with it, so they judge
 * the same parts (#346, #364, #398).
 */
export const isRelationshipsPartName = (partName: string): boolean => canonicalPartName(partName).endsWith(".rels");

/**
 * OPC extension of a part name: the substring after the last dot of the
 * final segment, ASCII-lowercased — `_rels/.rels` has extension "rels"
 * (unlike POSIX `path.extname`, which returns "" for dotfiles). Returns ""
 * when the final segment has no dot.
 */
export const partExtension = (partPath: string): string => {
    const base = partPath.slice(partPath.lastIndexOf("/") + 1);
    const dot = base.lastIndexOf(".");
    return dot < 0 ? "" : asciiLowercase(base.slice(dot + 1));
};

/**
 * The {@link partNameViolation} rule for a folder segment that ends with a
 * dot ([M1.9]). Word 16.115 opens `/word./x.xml`, `/word/sub./x.xml` and
 * `/x./x.xml` clean, and offers recovery when the last segment ends with one
 * (#450; session 7 in docs/plans/evidence/word-probes-2026-10-24.md).
 */
export const PART_NAME_FOLDER_SEGMENT_DOT = "must not have a folder segment ending with '.'";

/**
 * Why a part name is not a valid OPC part name (ECMA-376 Part 2, 4th
 * edition (2012), §9.1.1.1, [M1.1]–[M1.10]), or `undefined` when it names one.
 *
 * A part name shall be an IRI and shall be encoded as either a Part URI
 * (§9.1.1.1.2: `segment = 1*( pchar )`, RFC 3986) or a Part IRI
 * (§9.1.1.1.1: `isegment = 1*( ipchar )`, RFC 3987, which allows a literal
 * `ucschar`). A name is valid when it is either, so `/word/café.xml` (a
 * Part IRI) and `/word/caf%C3%A9.xml` (a Part URI) both pass, while a mixed
 * spelling like `/word/caf%C3%A9-é.xml` fails both: its literal `ucschar`
 * leaves only the Part IRI form, whose [M1.8] forbids percent-encoded
 * `iunreserved` characters. A literal non-ASCII character that is not
 * `ucschar` (a C1 control, private use, a noncharacter) is in neither
 * grammar. Strict validation follows this 4th edition, the edition of the
 * bundled `schemas/ecma/fourth-edition/` OPC schemas; the 5th edition
 * (§6.2.2.2) keeps only the IRI form and forbids `%C3%A9` instead.
 *
 * A part name is a path: it starts with `/`, its segments are neither empty
 * nor made only of dots ([M1.10]), it ends with no separator and no segment
 * ends with a dot, it carries no query, fragment or backslash, and every
 * character in it is an RFC 3986 `pchar`, an RFC 3987 `ipchar`, or a
 * percent-escape. A percent-escaped separator ([M1.7]) or unreserved
 * character ([M1.8]) is forbidden too.
 *
 * A folder segment ending with a dot ({@link PART_NAME_FOLDER_SEGMENT_DOT}) is
 * judged last, and a segment made only of dots never reaches it (#453). Word 16.115 opens such a name clean while it refuses every
 * other violation (#450), so a name that breaks another rule as well reports
 * that rule.
 *
 * These are the part-name rules, not the stored item-name rules: a part name
 * starts with a slash, where a stored item name must not.
 */
export const partNameViolation = (partName: string): string | undefined => {
    if (!partName.startsWith("/")) {
        return "must start with '/'";
    }
    const segments = partName.slice(1).split("/");
    if (partName.endsWith("/")) {
        return "must not end with '/'";
    }
    if (segments.some((segment) => segment === "." || segment === "..")) {
        return "must not have a '.' or '..' segment";
    }
    // `...` is no dot-segment, but it has no non-dot character either. Word
    // 16.115 offers recovery for it even as a folder segment, where it
    // tolerates `word./` (#453), so it must not reach the folder-dot rule.
    if (segments.some((segment) => /^\.+$/.test(segment))) {
        return "must not have a dot-only segment (OPC [M1.10])";
    }
    if (segments.some((segment) => segment === "")) {
        return "must not have an empty segment";
    }
    if (partName.endsWith(".")) {
        return "must not end with '.'";
    }
    if (/[?#\\]/.test(partName)) {
        return "must not contain '?', '#' or '\\'";
    }
    if (hasEscapedSeparatorCharacter(partName)) {
        return "must not percent-encode '/' or '\\' (OPC [M1.7])";
    }
    if (hasEscapedUnreservedCharacter(partName)) {
        return "must not percent-encode an unreserved character (OPC [M1.8])";
    }
    if (hasNonPcharCharacter(partName)) {
        // Not a Part URI (§9.1.1.1.2). The name can still be a Part IRI
        // (§9.1.1.1.1): every character an RFC 3987 `ipchar` allows — a
        // literal `ucschar` — with [M1.8] reading "iunreserved" there.
        if (hasNonIpcharCharacter(partName)) {
            return "must contain only RFC 3986 pchar characters";
        }
        if (hasEscapedIunreservedCharacter(partName)) {
            return "must not percent-encode an unreserved character (OPC [M1.8])";
        }
    }
    if (segments.slice(0, -1).some((segment) => segment.endsWith("."))) {
        return PART_NAME_FOLDER_SEGMENT_DOT;
    }
    return undefined;
};

/**
 * Whether a name holds a character outside the RFC 3986 `pchar` set. A
 * percent-escape counts as the escape it is; a `%` that opens no valid
 * escape is data, and data outside `pchar` is refused.
 */
const hasNonPcharCharacter = (name: string): boolean => {
    let index = 0;
    while (index < name.length) {
        const code = name.charCodeAt(index);
        if (code === 0x25 /* % */) {
            const high = hexValue(name.charCodeAt(index + 1));
            const low = hexValue(name.charCodeAt(index + 2));
            if (high >= 0 && low >= 0) {
                index += 3;
                continue;
            }
            return true;
        }
        if (!isPcharOrSlash(code)) {
            return true;
        }
        index += 1;
    }
    return false;
};

/**
 * Whether a name holds a code point outside the RFC 3987 `ipchar` set
 * (`pchar` widened by `ucschar`), the Part IRI grammar of ECMA-376 Part 2
 * §9.1.1.1.1. Called only for names already outside the Part URI grammar,
 * so the answer decides whether the name can be a Part IRI at all.
 */
const hasNonIpcharCharacter = (name: string): boolean => {
    let index = 0;
    while (index < name.length) {
        const codePoint = name.codePointAt(index)!;
        if (codePoint === 0x25 /* % */) {
            const high = hexValue(name.charCodeAt(index + 1));
            const low = hexValue(name.charCodeAt(index + 2));
            if (high >= 0 && low >= 0) {
                index += 3;
                continue;
            }
            return true;
        }
        if (!isPcharOrSlash(codePoint) && !isUcschar(codePoint)) {
            return true;
        }
        index += String.fromCodePoint(codePoint).length;
    }
    return false;
};

/**
 * Decode one UTF-8 code point from `bytes` at `offset`, with its byte
 * length, or `undefined` when the sequence is malformed (an ASCII byte in
 * a multi-byte position, a truncated or overlong sequence, a surrogate, or
 * a value above U+10FFFF). WHATWG replacement never applies: [M1.8] asks
 * what character the octets encode, not what a decoder would display.
 */
const decodeUtf8CodePoint = (bytes: readonly number[], offset: number): { codePoint: number; length: number } | undefined => {
    const lead = bytes[offset]!;
    if (lead < 0x80) {
        return { codePoint: lead, length: 1 };
    }
    let length = 0;
    let minimum = 0;
    let codePoint = 0;
    if (lead >= 0xc2 && lead <= 0xdf) {
        length = 2;
        minimum = 0x80;
        codePoint = lead - 0xc0; // the lead's low 5 bits
    } else if (lead >= 0xe0 && lead <= 0xef) {
        length = 3;
        minimum = 0x800;
        codePoint = lead - 0xe0; // the lead's low 4 bits
    } else if (lead >= 0xf0 && lead <= 0xf4) {
        length = 4;
        minimum = 0x10000;
        codePoint = lead - 0xf0; // the lead's low 3 bits
    } else {
        return undefined; // a continuation byte, or an overlong lead (C0/C1)
    }
    for (let index = 1; index < length; index += 1) {
        const continuation = bytes[offset + index];
        if (continuation === undefined || continuation < 0x80 || continuation > 0xbf) {
            return undefined;
        }
        codePoint = codePoint * 0x40 + (continuation - 0x80);
    }
    if (codePoint < minimum || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
        return undefined;
    }
    return { codePoint, length };
};

/**
 * OPC [M1.8] for the Part IRI form (§9.1.1.1.1): an isegment shall not
 * percent-encode an RFC 3987 `iunreserved` character. A percent-encoded
 * sequence counts when its octets decode, as UTF-8, to one: `%C3%A9` encodes
 * `é`. Malformed UTF-8 escapes (a lone `%E9`, an overlong `%C0%AF`) are
 * pct-encoded octets, not encoded `iunreserved` characters, and a run that
 * decodes past a non-`iunreserved` character (`%20%C3%A9`) still encodes the
 * `é`, so the scan decodes one code point at a time. A malformed sequence
 * consumes one octet and the scan resynchronizes at the next, so a valid
 * sequence behind it (`%E9%C3%A9`) is still seen (#523).
 */
const hasEscapedIunreservedCharacter = (name: string): boolean => {
    const bytes: number[] = [];
    const flush = (): boolean => {
        let offset = 0;
        while (offset < bytes.length) {
            const decoded = decodeUtf8CodePoint(bytes, offset);
            if (decoded === undefined) {
                // Pct-encoded octets, not an encoded character: consume one
                // octet and decode from the next.
                offset += 1;
                continue;
            }
            if (isIunreserved(decoded.codePoint)) {
                return true;
            }
            offset += decoded.length;
        }
        bytes.length = 0;
        return false;
    };
    for (let index = 0; index < name.length; index += 1) {
        if (name[index] === "%" && /^[0-9a-f]{2}$/i.test(name.slice(index + 1, index + 3))) {
            bytes.push(Number.parseInt(name.slice(index + 1, index + 3), 16));
            index += 2;
            continue;
        }
        if (flush()) {
            return true;
        }
    }
    return flush();
};

/** Extract the declarations from a parsed [Content_Types].xml document. */
export const parseContentTypes = (dom: Document): ContentTypeTable => {
    const overrides = new Map<string, string>();
    const defaults = new Map<string, string>();

    for (const elem of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override")) {
        const partName = elem.getAttribute("PartName");
        if (!partName) {
            continue;
        }
        // First declaration wins: duplicate PartNames are themselves a
        // package defect, and coverage only needs one authoritative entry.
        const key = canonicalPartName(partName);
        if (!overrides.has(key)) {
            overrides.set(key, elem.getAttribute("ContentType") ?? "");
        }
    }

    for (const elem of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Default")) {
        const ext = elem.getAttribute("Extension");
        if (!ext) {
            continue;
        }
        const key = asciiLowercase(ext);
        if (!defaults.has(key)) {
            defaults.set(key, elem.getAttribute("ContentType") ?? "");
        }
    }

    return { overrides, defaults };
};

/**
 * Diagnose each repeated declaration against its first spelling (#431).
 * This reads raw declarations without changing the first-entry resolution
 * table. Callers must supply a successfully parsed content-types stream.
 */
export const duplicateContentTypeDeclarations = (dom: Document, stream: string): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const [element, attribute, canonicalize] of [
        ["Default", "Extension", asciiLowercase],
        ["Override", "PartName", canonicalPartName],
    ] as const) {
        const first = new Map<string, string>();
        for (const declaration of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, element)) {
            const spelling = declaration.getAttribute(attribute);
            if (!spelling) {
                continue;
            }
            const key = canonicalize(spelling);
            const previous = first.get(key);
            if (previous === undefined) {
                first.set(key, spelling);
            } else {
                issues.push({
                    code: CT_DUPLICATE_DECLARATION,
                    severity: "error",
                    path: stream,
                    message: `Duplicate ${element} ${attribute} declaration: ${JSON.stringify(spelling)} repeats ${JSON.stringify(previous)} (key ${JSON.stringify(key)}).`,
                });
            }
        }
    }
    return issues;
};

/**
 * The media type of a content-type value: everything before the first `;`,
 * trimmed and ASCII-lowercased. RFC 7303 §9.1 registers parameters such as
 * `charset` as optional additions to a media type, so two values that differ
 * only in parameters share one media type (#540).
 */
export const contentTypeMediaType = (contentType: string): string => {
    const semicolon = contentType.indexOf(";");
    return asciiLowercase((semicolon === -1 ? contentType : contentType.slice(0, semicolon)).trim());
};

/**
 * Whether a content-type value carries a MIME parameter section — a `;`
 * after the media type (#540). Word 16.115 compares the declared string
 * exactly, so a parameter it does not honour is a defect in its own right,
 * not a different media type.
 */
export const hasContentTypeParameter = (contentType: string): boolean => contentType.includes(";");

/**
 * Resolve a part name against the table (ECMA-376 Part 2 §10.1.2.4): an
 * exact `<Override PartName>` match wins; otherwise the `<Default
 * Extension>` for the part's OPC extension applies; `null` when neither
 * covers the part. Both the PartName and the path are compared in their
 * {@link canonicalPartName} form.
 */
export const resolveContentType = (table: ContentTypeTable, partName: string): ResolvedContentType | null => {
    const key = canonicalPartName(partName);
    const override = table.overrides.get(key);
    if (override !== undefined) {
        return { contentType: override, source: "override" };
    }
    const ext = partExtension(key);
    const byDefault = ext === "" ? undefined : table.defaults.get(ext);
    if (byDefault !== undefined) {
        return { contentType: byDefault, source: "default" };
    }
    return null;
};
