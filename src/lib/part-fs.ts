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

import { existsSync, mkdirSync, promises as fs, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";

export interface PartFS {
    /** The (possibly virtual) unpacked-package root these paths are relative to. */
    readonly root: string;
    /** Absolute paths of every part, optionally filtered by extension; sorted. */
    list(extensions?: ReadonlyArray<string>): string[];
    /** Whether a part exists at this absolute path. */
    exists(absPath: string): boolean;
    readText(absPath: string): Promise<string>;
    readBytes(absPath: string): Promise<Buffer>;
    readTextSync(absPath: string): string;
    readBytesSync(absPath: string): Buffer;
    /** Write a part; parent "directories" are created implicitly. */
    write(absPath: string, content: string | Buffer): Promise<void>;
    /** Delete a part. Missing parts are a no-op. */
    remove(absPath: string): Promise<void>;
}

/** Disk-backed PartFS — preserves the original `node:fs` behavior exactly. */
export class DiskPartFS implements PartFS {
    readonly root: string;

    constructor(unpackedDir: string) {
        this.root = path.resolve(unpackedDir);
    }

    list(extensions?: ReadonlyArray<string>): string[] {
        const results: string[] = [];
        const stack: string[] = [this.root];
        while (stack.length > 0) {
            const current = stack.pop();
            if (current === undefined) break;
            let entries: string[];
            try {
                entries = readdirSync(current);
            } catch {
                continue;
            }
            for (const name of entries) {
                const full = path.join(current, name);
                let st;
                try {
                    st = statSync(full);
                } catch {
                    continue;
                }
                if (st.isDirectory()) {
                    stack.push(full);
                } else if (st.isFile()) {
                    if (!extensions || extensions.some((e) => full.endsWith(e))) {
                        results.push(full);
                    }
                }
            }
        }
        return results.sort();
    }

    exists(absPath: string): boolean {
        return existsSync(absPath);
    }

    readText(absPath: string): Promise<string> {
        return fs.readFile(absPath, "utf-8");
    }

    readBytes(absPath: string): Promise<Buffer> {
        return fs.readFile(absPath);
    }

    readTextSync(absPath: string): string {
        return readFileSync(absPath, "utf-8");
    }

    readBytesSync(absPath: string): Buffer {
        return readFileSync(absPath);
    }

    async write(absPath: string, content: string | Buffer): Promise<void> {
        await fs.mkdir(path.dirname(absPath), { recursive: true });
        await fs.writeFile(absPath, content);
    }

    async remove(absPath: string): Promise<void> {
        await fs.rm(absPath, { force: true });
    }
}

/**
 * In-memory PartFS — no disk, no native deps. Keys are normalized to the part's
 * path relative to `root` (POSIX separators), so it round-trips with the same
 * `path.join(root, rel)` / `path.relative(root, abs)` arithmetic the validators
 * use. Construct from a `Map<relPath, content>` (e.g. a JSZip unpack).
 */
export class MemoryPartFS implements PartFS {
    readonly root: string;
    private readonly parts = new Map<string, Buffer>();

    constructor(parts?: Iterable<[string, string | Buffer]>, root = "/__mem__") {
        this.root = root;
        if (parts) {
            for (const [rel, content] of parts) {
                this.parts.set(this.normalizeRel(rel), this.toBuffer(content));
            }
        }
    }

    private toBuffer(content: string | Buffer): Buffer {
        return Buffer.isBuffer(content) ? content : Buffer.from(content, "utf-8");
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
    entries(): Array<[string, Buffer]> {
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

    private get(absPath: string): Buffer {
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
        return this.get(absPath).toString("utf-8");
    }

    async readBytes(absPath: string): Promise<Buffer> {
        return this.get(absPath);
    }

    readTextSync(absPath: string): string {
        return this.get(absPath).toString("utf-8");
    }

    readBytesSync(absPath: string): Buffer {
        return this.get(absPath);
    }

    write(absPath: string, content: string | Buffer): Promise<void> {
        this.parts.set(this.relKey(absPath), this.toBuffer(content));
        return Promise.resolve();
    }

    remove(absPath: string): Promise<void> {
        this.parts.delete(this.relKey(absPath));
        return Promise.resolve();
    }
}
