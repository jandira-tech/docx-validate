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

import type JSZip from "jszip";

import { applyMarkupCompatibility, understoodNamespacesForPart } from "./markup-compatibility";
import { zipCentralDirectoryEntries } from "./zip-central-directory";
import type { ValidationIssue } from "./types";
import { decodeXmlBytes, parseXml } from "./xml-helpers";
import { loadCentralNamedArchive } from "./zip-central-names";

/** Where the package came from. Archive-only checks need `"archive"`. */
export type PackageSource = "archive" | "directory" | "memory";

/** Part content, eager or loaded on first use. */
export type PartContent = string | Uint8Array | (() => Uint8Array);

/** Outcome of parsing one XML part. */
export type XmlPart = { readonly ok: true; readonly document: Document } | { readonly ok: false; readonly message: string };

/**
 * The package's view could not be read: a stored name cannot be placed, or
 * an archive entry's bytes could not be read. Staging refuses the same input
 * for the same reason, so callers report an extraction failure rather than
 * an open failure (#413, #424).
 */
export class PackageViewError extends Error {
    /**
     * The archive's stored file-entry names, read before the view was
     * refused, so the checks that need only them can still run (#420).
     */
    public readonly storedNames: readonly string[];

    public constructor(message: string, options?: ErrorOptions & { readonly storedNames?: readonly string[] }) {
        super(message, options);
        this.name = "PackageViewError";
        this.storedNames = options?.storedNames ?? [];
    }
}

/**
 * The archive bytes are not a readable ZIP: encrypted OOXML, or a truncated
 * or corrupted archive. Distinct from {@link PackageViewError} — which names
 * an archive that opened but whose package view cannot be read — and from an
 * internal error, which propagates unwrapped (#427).
 */
export class PackageOpenError extends Error {
    public constructor(message: string, options?: ErrorOptions) {
        super(message, options);
        this.name = "PackageOpenError";
    }
}

export type ValidationContextOptions = {
    readonly source?: PackageSource;
    /** The packed file's bytes; present only when `source` is `"archive"`. */
    readonly archiveBytes?: Uint8Array;
};

// Matches by base name in any folder, like DOCXSchemaValidator.userTextXmlFiles.
const STORY_PART = /(?:^|\/)(?:document|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/i;

/** Part names are package-relative: forward slashes, no leading slash. */
export const normalizePartName = (name: string): string => name.replace(/\\/g, "/").replace(/^\/+/, "");

/**
 * Why a ZIP item name places no part in the package, or `undefined` when it
 * does. Reading `\` as `/`, as part names do, the name is refused when it
 * starts with a slash or names a drive, which the ZIP format forbids
 * (APPNOTE 4.4.17.1) and System.IO.Packaging ignores (#372). A first segment
 * that starts with a letter and `:` names a drive even without a slash
 * (`c:doc/x.xml`): Windows resolves it against that drive's working
 * directory, so it is refused on every platform. A name is also refused
 * when it has an empty or a dot-only segment, which OPC forbids in part
 * names ([M1.3]; [M1.9] and [M1.10] cover `...` too). JSZip and extraction
 * rewrite `.`, `..` and empty segments, so the two views would otherwise
 * judge a name the package does not hold (#378).
 */
export const entryNameRefusal = (name: string): string | undefined => {
    const posix = name.replace(/\\/g, "/");
    const reason = posix.startsWith("/")
        ? "it starts with a slash"
        : /^[A-Za-z]:/.test(posix)
          ? "it names a drive"
          : posix.split("/").some((segment) => segment === "" || /^\.+$/.test(segment))
            ? "it has an empty or dot-only segment"
            : undefined;
    return reason === undefined ? undefined : `ZIP item ${JSON.stringify(name)} names no package part: ${reason}`;
};

/**
 * Throw for the first stored name that places no part. Central-directory
 * names include entries the loaded view dropped or rewrote (#411), such as
 * `a/b/../x.xml` stored before the `a/x.xml` it resolves to. Distinct normalized part or
 * folder spellings that collide under ASCII case and NFC folding are refused
 * on every host, before extraction, rather than depending on its filesystem
 * (#441). Exact normalized aliases and identical stored names keep the
 * existing entry-selection policy.
 */
export const assertPlaceableNames = (names: Iterable<string>): void => {
    const placed = new Map<string, string>();
    for (const name of names) {
        const refusal = entryNameRefusal(name);
        if (refusal !== undefined) {
            throw new Error(refusal);
        }
        // Decide from names before touching disk, on every host (#441).
        // Include folders: differently spelled folders can merge even when
        // their leaf names differ (#376). Exact normalized aliases still
        // stage as one part, so the existing duplicate choice handles them.
        const segments = normalizePartName(name).split("/");
        for (let length = 1; length <= segments.length; length += 1) {
            const prefix = segments.slice(0, length).join("/");
            const key = prefix.normalize("NFC").replace(/[A-Z]/g, (letter) => letter.toLowerCase());
            const previous = placed.get(key);
            if (previous !== undefined && previous !== prefix) {
                throw new Error(
                    `Package paths ${JSON.stringify(previous)} and ${JSON.stringify(prefix)} collide under portable ASCII-case/NFC folding`,
                );
            }
            placed.set(key, prefix);
        }
    }
};

/**
 * Throw for the first entry whose stored name places no part. JSZip has
 * already rewritten some names (`..`, `.`, empty segments, a leading root),
 * so the name as stored in the archive is judged (#378).
 */
export const assertPlaceableEntries = (entries: Iterable<JSZip.JSZipObject>): void => {
    assertPlaceableNames([...entries].map((entry) => entry.unsafeOriginalName ?? entry.name));
};

/**
 * The archive's stored file-entry names: the central directory's, in order
 * and with duplicates kept, when it parses; the loaded view's
 * `unsafeOriginalName`s otherwise. A malformed envelope falls back because
 * the ZIP envelope checks already diagnose it (#411).
 */
export const archiveStoredNames = (bytes: Uint8Array, files: readonly JSZip.JSZipObject[]): readonly string[] =>
    zipCentralDirectoryEntries(bytes)
        ?.filter((entry) => !entry.directory)
        .map((entry) => entry.name) ?? files.map((entry) => entry.unsafeOriginalName ?? entry.name);

/**
 * Of two stored names for one part, the one to keep: the name that already is
 * the part name (no `\`, no leading `/`), else the smaller in code-unit
 * order. The choice never depends on entry order, so a packed archive and its
 * extracted directory keep the same entry (#366).
 */
const preferredStoredName = (a: string, b: string): string => {
    const aIsPartName = a === normalizePartName(a);
    if (aIsPartName !== (b === normalizePartName(b))) {
        return aIsPartName ? a : b;
    }
    return a < b ? a : b;
};

/** Index entries by part name, keeping for each part the entry {@link preferredStoredName} picks. */
export const indexByPartName = <T>(entries: Iterable<readonly [string, T]>): Map<string, T> => {
    const stored = new Map<string, string>();
    const parts = new Map<string, T>();
    for (const [name, value] of entries) {
        const part = normalizePartName(name);
        const kept = stored.get(part);
        if (kept === undefined || preferredStoredName(kept, name) === name) {
            stored.set(part, name);
            parts.set(part, value);
        }
    }
    return parts;
};

export class ValidationContext {
    public readonly source: PackageSource;
    public readonly archiveBytes: Uint8Array | undefined;
    /**
     * Original file item names before normalized indexing drops aliases — for
     * an archive, the central directory's stored names, including entries the
     * loaded view dropped or renamed (#411); sorted and immutable.
     */
    public readonly storedNames: readonly string[];

    private readonly parts: ReadonlyMap<string, PartContent>;
    private readonly names: readonly string[];
    private readonly byteCache: Map<string, Uint8Array>;
    private readonly textCache: Map<string, string>;
    private readonly xmlCache: Map<string, XmlPart>;
    private readonly effectiveCache: Map<string, { readonly part: XmlPart; readonly issues: readonly ValidationIssue[] }>;
    private readonly reads: Map<string, number>;
    private readonly parses: Map<string, number>;
    /** Parts this view reports as not parsing, with the reason. */
    private readonly unparseable: ReadonlyMap<string, string>;

    /** `base`, when given, is the context whose caches this view shares. */
    private constructor(
        parts: ReadonlyMap<string, PartContent>,
        opts: ValidationContextOptions,
        storedNames: readonly string[],
        base?: ValidationContext,
        unparseable: ReadonlyMap<string, string> = new Map(),
    ) {
        this.parts = parts;
        this.storedNames = Object.freeze([...storedNames].sort());
        this.names = [...parts.keys()].sort();
        this.source = opts.source ?? "memory";
        this.archiveBytes = opts.archiveBytes;
        this.byteCache = base?.byteCache ?? new Map();
        this.textCache = base?.textCache ?? new Map();
        this.xmlCache = base?.xmlCache ?? new Map();
        this.effectiveCache = base?.effectiveCache ?? new Map();
        this.reads = base?.reads ?? new Map();
        this.parses = base?.parses ?? new Map();
        this.unparseable = unparseable;
    }

    /** Build a context from part names (leading `/` optional) and their content. */
    public static fromEntries(
        entries: Readonly<Record<string, PartContent>> | Iterable<readonly [string, PartContent]>,
        opts: ValidationContextOptions = {},
    ): ValidationContext {
        const pairs = [...(Symbol.iterator in entries ? entries : Object.entries(entries))];
        return new ValidationContext(
            indexByPartName(pairs),
            opts,
            pairs.map(([name]) => name),
        );
    }

    /**
     * Load a packed OPC archive, one entry per part name. Throws a
     * {@link PackageOpenError} when the bytes are not a readable ZIP, or a
     * {@link PackageViewError} when the opened archive's package view cannot
     * be read: an item name places no part (#372, #378), or an entry cannot
     * be read (#369) — the same archives extraction refuses, so callers
     * report the same failure. Any other throw is an internal error and
     * propagates unwrapped rather than wearing a file-defect code (#427).
     */
    public static async fromArchive(bytes: Uint8Array): Promise<ValidationContext> {
        let zip: JSZip;
        try {
            zip = await loadCentralNamedArchive(bytes);
        } catch (err) {
            throw new PackageOpenError(err instanceof Error ? err.message : String(err), { cause: err });
        }
        // The stored names survive a refused package view, so the checks that
        // need only them can run on the thrown error's copy (#420).
        let stored: readonly string[] = [];
        const files = Object.values(zip.files).filter((entry) => !entry.dir);
        stored = archiveStoredNames(bytes, files);
        try {
            assertPlaceableNames(stored);
        } catch (err) {
            throw new PackageViewError(err instanceof Error ? err.message : String(err), { cause: err, storedNames: stored });
        }
        const kept = indexByPartName(files.map((entry) => [entry.name, entry] as const));
        const parts = new Map<string, PartContent>();
        // Every entry is read, as extraction reads it, so an unreadable one
        // rejects the archive even when another entry names its part (#369).
        // Only the two failures a real extraction reports — a refused item
        // name above and an unreadable entry here — are PackageViewError;
        // anything else is an internal error and propagates unwrapped (#427).
        try {
            await Promise.all(
                files.map(async (entry) => {
                    const content = await entry.async("uint8array");
                    const part = normalizePartName(entry.name);
                    if (kept.get(part) === entry) {
                        parts.set(part, content);
                    }
                }),
            );
        } catch (err) {
            throw new PackageViewError(err instanceof Error ? err.message : String(err), { cause: err, storedNames: stored });
        }
        return new ValidationContext(parts, { source: "archive", archiveBytes: bytes }, stored);
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
        const reason = this.unparseable.get(name);
        if (reason !== undefined && this.parts.has(name)) {
            return { ok: false, message: reason };
        }
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
        const parsed = this.xml(name);
        if (!parsed?.ok) {
            return parsed;
        }
        return this.effective(name, parsed.document).part;
    }

    /**
     * A view of this package in which `parts` (part name to reason) do not
     * parse, so no check judges a DOM the lenient parser recovered from
     * them. The view shares this context's caches; this context is unchanged.
     */
    public withUnparseable(parts: ReadonlyMap<string, string>): ValidationContext {
        const unparseable = new Map(this.unparseable);
        for (const [part, reason] of parts) {
            unparseable.set(normalizePartName(part), reason);
        }
        return new ValidationContext(
            this.parts,
            { source: this.source, archiveBytes: this.archiveBytes },
            this.storedNames,
            this,
            unparseable,
        );
    }

    /**
     * What markup compatibility found while building {@link effectiveXml}'s
     * view, with the part's path. Empty for a missing or unparseable part.
     */
    public compatibilityIssues(part: string): readonly ValidationIssue[] {
        const name = normalizePartName(part);
        const parsed = this.xml(name);
        if (!parsed?.ok) {
            return [];
        }
        return this.effective(name, parsed.document).issues;
    }

    private effective(name: string, document: Document): { readonly part: XmlPart; readonly issues: readonly ValidationIssue[] } {
        const cached = this.effectiveCache.get(name);
        if (cached) {
            return cached;
        }
        const view = applyMarkupCompatibility(document, { understood: understoodNamespacesForPart(name) });
        const effective = {
            part: { ok: true, document: view.document } as const,
            issues: view.issues.map((issue) => ({ ...issue, path: name })),
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

    /**
     * A context with no parts that still carries the stored entry names — for
     * an archive whose package view was refused, so the checks that need only
     * the names run on the same list a full run diagnoses (#420).
     */
    public static fromStoredNames(storedNames: readonly string[], opts: ValidationContextOptions = {}): ValidationContext {
        return new ValidationContext(new Map(), opts, storedNames);
    }
}
