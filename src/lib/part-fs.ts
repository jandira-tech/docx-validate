/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

// PartFS — a thin filesystem abstraction over an unpacked OOXML package.
//
// The validators/repairers historically read and wrote each part directly via
// `node:fs` against an `unpackedDir`. That couples them to disk (and to Node).
// PartFS keeps the exact same call shape — methods take ABSOLUTE paths rooted at
// `root`, so `path.join(root, rel)` and `path.relative(root, abs)` keep working —
// but lets the backing store be either real disk (`DiskPartFS`, unchanged Node
// behavior) or an in-memory map (`MemoryPartFS`, no `fs`, browser-safe). Repair
// uses only `@xmldom/xmldom` for parsing, so an in-memory repair has no native
// dependency and runs in the browser.

// Browser-safe: this module uses only pathlite (no `node:fs`/`node:path`), so it
// bundles for the browser. The disk-backed `DiskPartFS` lives in
// `part-fs.node.ts` (Node-only) to keep `node:fs` out of the browser graph.
import path from "./pathlite";

export interface PartFS {
    /** The (possibly virtual) unpacked-package root these paths are relative to. */
    readonly root: string;
    /** Absolute paths of every part, optionally filtered by extension; sorted. */
    list(extensions?: ReadonlyArray<string>): string[];
    /** Whether a part exists at this absolute path. */
    exists(absPath: string): boolean;
    readText(absPath: string): Promise<string>;
    readBytes(absPath: string): Promise<Uint8Array>;
    readTextSync(absPath: string): string;
    readBytesSync(absPath: string): Uint8Array;
    /** Write a part; parent "directories" are created implicitly. */
    write(absPath: string, content: string | Uint8Array): Promise<void>;
    /** Delete a part. Missing parts are a no-op. */
    remove(absPath: string): Promise<void>;
}

/**
 * In-memory PartFS — no disk, no native deps. Keys are normalized to the part's
 * path relative to `root` (POSIX separators), so it round-trips with the same
 * `path.join(root, rel)` / `path.relative(root, abs)` arithmetic the validators
 * use. Construct from a `Map<relPath, content>` (e.g. a JSZip unpack).
 */
export class MemoryPartFS implements PartFS {
    readonly root: string;
    private readonly parts = new Map<string, Uint8Array>();
    private static readonly _encoder = new TextEncoder();
    private static readonly _decoder = new TextDecoder("utf-8");

    constructor(parts?: Iterable<[string, string | Uint8Array]>, root = "/__mem__") {
        this.root = root;
        if (parts) {
            for (const [rel, content] of parts) {
                this.parts.set(this.normalizeRel(rel), this.toBytes(content));
            }
        }
    }

    // Browser-safe: no Node `Buffer`. Strings become UTF-8 byte arrays via
    // TextEncoder; existing byte arrays pass through.
    private toBytes(content: string | Uint8Array): Uint8Array {
        return typeof content === "string" ? MemoryPartFS._encoder.encode(content) : content;
    }

    /** Absolute → relative POSIX key. Accepts already-relative inputs too. */
    private relKey(absPath: string): string {
        const rel = absPath.startsWith(this.root) ? path.relative(this.root, absPath) : absPath;
        return this.normalizeRel(rel);
    }

    private normalizeRel(rel: string): string {
        return rel.replace(/^[/\\]+/, "").split(path.sep).join("/");
    }

    /** Snapshot of every part as `[relPath, bytes]`, for repacking. */
    entries(): Array<[string, Uint8Array]> {
        return [...this.parts.entries()];
    }

    list(extensions?: ReadonlyArray<string>): string[] {
        const out: string[] = [];
        for (const rel of this.parts.keys()) {
            if (!extensions || extensions.some((e) => rel.endsWith(e))) {
                out.push(path.join(this.root, rel));
            }
        }
        return out.sort();
    }

    exists(absPath: string): boolean {
        return this.parts.has(this.relKey(absPath));
    }

    private get(absPath: string): Uint8Array {
        const key = this.relKey(absPath);
        const buf = this.parts.get(key);
        if (buf === undefined) {
            const err = new Error(`ENOENT: no in-memory part '${key}'`) as NodeJS.ErrnoException;
            err.code = "ENOENT";
            throw err;
        }
        return buf;
    }

    async readText(absPath: string): Promise<string> {
        return MemoryPartFS._decoder.decode(this.get(absPath));
    }

    async readBytes(absPath: string): Promise<Uint8Array> {
        return this.get(absPath);
    }

    readTextSync(absPath: string): string {
        return MemoryPartFS._decoder.decode(this.get(absPath));
    }

    readBytesSync(absPath: string): Uint8Array {
        return this.get(absPath);
    }

    write(absPath: string, content: string | Uint8Array): Promise<void> {
        this.parts.set(this.relKey(absPath), this.toBytes(content));
        return Promise.resolve();
    }

    remove(absPath: string): Promise<void> {
        this.parts.delete(this.relKey(absPath));
        return Promise.resolve();
    }
}
