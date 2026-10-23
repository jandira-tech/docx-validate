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
 * One package, loaded once, shared by every named check of a run.
 *
 * The context holds the package's part names, the original archive bytes when
 * the input was a packed file, and lazy caches: each part is read at most once
 * and each XML part is parsed at most once per context. A context belongs to a
 * single validation run; build a new one for the next run so no document can
 * leak between validations.
 *
 * The module has no filesystem dependency. Callers that start from a directory
 * pass lazy loaders (see `scripts/office/package-context.ts`).
 */

import JSZip from "jszip";

import { applyMarkupCompatibility, WORD_UNDERSTOOD_NAMESPACES } from "./markup-compatibility";
import { decodeXmlBytes, parseXml } from "./xml-helpers";

/** Where the package came from. Archive-only checks need `"archive"`. */
export type PackageSource = "archive" | "directory" | "memory";

/** Part content, eager or loaded on first use. */
export type PartContent = string | Uint8Array | (() => Uint8Array);

/** Outcome of parsing one XML part. */
export type XmlPart = { readonly ok: true; readonly document: Document } | { readonly ok: false; readonly message: string };

export type ValidationContextOptions = {
    readonly source?: PackageSource;
    /** The packed file's bytes; present only when `source` is `"archive"`. */
    readonly archiveBytes?: Uint8Array;
};

// Matches by base name in any folder, like DOCXSchemaValidator.userTextXmlFiles.
const STORY_PART = /(?:^|\/)(?:document|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/i;

/** Part names are package-relative: forward slashes, no leading slash. */
export const normalizePartName = (name: string): string => name.replace(/\\/g, "/").replace(/^\/+/, "");

export class ValidationContext {
    public readonly source: PackageSource;
    public readonly archiveBytes: Uint8Array | undefined;

    private readonly parts: ReadonlyMap<string, PartContent>;
    private readonly names: readonly string[];
    private readonly byteCache = new Map<string, Uint8Array>();
    private readonly textCache = new Map<string, string>();
    private readonly xmlCache = new Map<string, XmlPart>();
    private readonly effectiveCache = new Map<string, XmlPart>();
    private readonly reads = new Map<string, number>();
    private readonly parses = new Map<string, number>();

    private constructor(parts: ReadonlyMap<string, PartContent>, opts: ValidationContextOptions) {
        this.parts = parts;
        this.names = [...parts.keys()].sort();
        this.source = opts.source ?? "memory";
        this.archiveBytes = opts.archiveBytes;
    }

    /** Build a context from part names (leading `/` optional) and their content. */
    public static fromEntries(
        entries: Readonly<Record<string, PartContent>> | Iterable<readonly [string, PartContent]>,
        opts: ValidationContextOptions = {},
    ): ValidationContext {
        const pairs: Iterable<readonly [string, PartContent]> = Symbol.iterator in entries ? entries : Object.entries(entries);
        const parts = new Map<string, PartContent>();
        for (const [name, content] of pairs) {
            parts.set(normalizePartName(name), content);
        }
        return new ValidationContext(parts, opts);
    }

    /** Load every entry of a packed OPC archive. Throws when the bytes are not a readable ZIP. */
    public static async fromArchive(bytes: Uint8Array): Promise<ValidationContext> {
        const zip = await JSZip.loadAsync(bytes);
        const parts = new Map<string, PartContent>();
        await Promise.all(
            Object.values(zip.files)
                .filter((entry) => !entry.dir)
                .map(async (entry) => {
                    parts.set(normalizePartName(entry.name), await entry.async("uint8array"));
                }),
        );
        return new ValidationContext(parts, { source: "archive", archiveBytes: bytes });
    }

    /** Part names in sorted order, without a leading `/`. */
    public get partNames(): readonly string[] {
        return this.names;
    }

    public has(part: string): boolean {
        return this.parts.has(normalizePartName(part));
    }

    /** The part's bytes, or `undefined` when the package has no such part. */
    public bytes(part: string): Uint8Array | undefined {
        const name = normalizePartName(part);
        const cached = this.byteCache.get(name);
        if (cached) {
            return cached;
        }
        const content = this.parts.get(name);
        if (content === undefined) {
            return undefined;
        }
        const bytes = typeof content === "string" ? new TextEncoder().encode(content) : typeof content === "function" ? content() : content;
        this.reads.set(name, (this.reads.get(name) ?? 0) + 1);
        this.byteCache.set(name, bytes);
        return bytes;
    }

    /** The part decoded as text: byte order mark first, then the XML declaration's encoding. */
    public text(part: string): string | undefined {
        const name = normalizePartName(part);
        const cached = this.textCache.get(name);
        if (cached !== undefined) {
            return cached;
        }
        const bytes = this.bytes(name);
        if (bytes === undefined) {
            return undefined;
        }
        const text = decodeXmlBytes(bytes);
        this.textCache.set(name, text);
        return text;
    }

    /** The parsed XML part, or `undefined` when the package has no such part. Parsed once. */
    public xml(part: string): XmlPart | undefined {
        const name = normalizePartName(part);
        const cached = this.xmlCache.get(name);
        if (cached) {
            return cached;
        }
        const text = this.text(name);
        if (text === undefined) {
            return undefined;
        }
        this.parses.set(name, (this.parses.get(name) ?? 0) + 1);
        let parsed: XmlPart;
        try {
            parsed = { ok: true, document: parseXml(text) };
        } catch (err) {
            parsed = { ok: false, message: err instanceof Error ? err.message : String(err) };
        }
        this.xmlCache.set(name, parsed);
        return parsed;
    }

    /**
     * The part as Microsoft Word consumes it: the parsed XML after markup
     * compatibility, with the namespaces Word understands.
     */
    public effectiveXml(part: string): XmlPart | undefined {
        const name = normalizePartName(part);
        const cached = this.effectiveCache.get(name);
        if (cached) {
            return cached;
        }
        const parsed = this.xml(name);
        if (!parsed?.ok) {
            return parsed;
        }
        const effective: XmlPart = {
            ok: true,
            document: applyMarkupCompatibility(parsed.document, { understood: WORD_UNDERSTOOD_NAMESPACES }).document,
        };
        this.effectiveCache.set(name, effective);
        return effective;
    }

    /**
     * WordprocessingML story parts that carry user text: document, comments,
     * footnotes, endnotes, headers, and footers.
     */
    public storyParts(): readonly string[] {
        return this.names.filter((name) => STORY_PART.test(name));
    }

    /** How many times `part` was read from its source. Diagnostics for cache tests. */
    public readCount(part: string): number {
        return this.reads.get(normalizePartName(part)) ?? 0;
    }

    /** How many times `part` was parsed as XML. Diagnostics for cache tests. */
    public parseCount(part: string): number {
        return this.parses.get(normalizePartName(part)) ?? 0;
    }
}
