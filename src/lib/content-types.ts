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
 * PartName comparison follows OPC part-name equivalence: ASCII
 * case-insensitive, with percent-escapes decoded (a PartName is a URI; the
 * on-disk path is its decoded form). Pure functions only — no filesystem
 * access; callers hand in the parsed [Content_Types].xml DOM.
 */

import { getElementsByTagNameNSAll } from "./xml-helpers";

export const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** Normalized declaration table extracted from a [Content_Types].xml DOM. */
export interface ContentTypeTable {
    /** Normalized PartName → ContentType for every `<Override>`. */
    readonly overrides: ReadonlyMap<string, string>;
    /** ASCII-lowercased Extension → ContentType for every `<Default>`. */
    readonly defaults: ReadonlyMap<string, string>;
}

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
 * Normalize a part name (or package-relative file path) to its comparison
 * form: strip leading `/` separators, decode percent-escapes safely
 * (a malformed escape keeps its raw form rather than throwing),
 * ASCII-lowercase. Normalization is for comparison only — the resolved
 * `contentType` is the declared value, never the normalized key.
 */
export function normalizePartName(partName: string): string {
    return asciiLowercase(partName.replace(/^\/+/, ""));
}

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
export function parseContentTypes(dom: Document): ContentTypeTable {
    const overrides = new Map<string, string>();
    const defaults = new Map<string, string>();

    for (const elem of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override")) {
        const partName = elem.getAttribute("PartName");
        // OPC PartName is a package-relative URI; the on-disk path is its
        // decoded form, so decode before storing or an escaped PartName
        // (%20 et al.) never matches the walked file list. Fall back to
        // the raw value on malformed escapes — the lookup then simply
        // misses and the residual pass reports the part.
        if (!partName) continue;
        let decoded = partName.replace(/^\/+/, "");
        try {
            decoded = decodeURIComponent(decoded);
        } catch {
            // keep raw
        }
        overrides.set(asciiLowercase(decoded), elem.getAttribute("ContentType") ?? "");
    }

    for (const elem of getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Default")) {
        const ext = elem.getAttribute("Extension");
        if (!ext) continue;
        defaults.set(asciiLowercase(ext), elem.getAttribute("ContentType") ?? "");
    }

    return { overrides, defaults };
}

/**
 * Resolve a part name against the table (ECMA-376 Part 2 §10.1.2.4): an
 * exact `<Override PartName>` match wins; otherwise the `<Default
 * Extension>` for the part's OPC extension applies; `null` when neither
 * covers the part.
 */
export function resolveContentType(table: ContentTypeTable, partName: string): ResolvedContentType | null {
    const key = normalizePartName(partName);
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
}
