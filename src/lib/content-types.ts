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

import { getElementsByTagNameNSAll } from "./xml-helpers";

export const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

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
const asciiLowercase = (value: string): string => {
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
 * The fully percent-decoded form of a part name: leading `/` separators
 * stripped (after decoding, so an encoded leading `%2F` strips too) and
 * ASCII-lowercased. This is the decoded display form, not the comparison
 * key — Override lookups use {@link canonicalPartName}, which keeps escapes
 * of reserved characters distinct from their literals (issue #297).
 */
export const normalizePartName = (partName: string): string => asciiLowercase(decodePercentEscapes(partName).replace(/^\/+/, ""));

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
