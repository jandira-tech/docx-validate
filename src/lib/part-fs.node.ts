/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// DiskPartFS — the Node-only, disk-backed PartFS. Split out of `part-fs.ts` so
// that `node:fs` never enters the browser bundle. The browser uses MemoryPartFS
// (in `part-fs.ts`) exclusively.

import { existsSync, promises as fs, readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";

import type { PartFS } from "./part-fs";

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

    async write(absPath: string, content: string | Uint8Array): Promise<void> {
        await fs.mkdir(path.dirname(absPath), { recursive: true });
        await fs.writeFile(absPath, content);
    }

    async remove(absPath: string): Promise<void> {
        await fs.rm(absPath, { force: true });
    }
}
