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

import { promises as fs } from "node:fs";
import path from "node:path";

import type JSZip from "jszip";

import { assertPlaceableEntries, indexByPartName, normalizePartName } from "./validation-context";

/**
 * `text` with every control character (C0, DEL, C1) and bidirectional
 * formatting character written as a `\u{…}` escape. An entry name comes from
 * the archive, and messages embed it, so a crafted name must not reach the
 * terminal that prints them as anything but text (#675).
 */
export const escapeControlCharacters = (text: string): string =>
    text.replace(/[\p{Cc}\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/gu, (char) => `\\u{${char.codePointAt(0)!.toString(16)}}`);

/**
 * Resolve a zip entry against `outputDir` and refuse path-traversal names
 * (Zip Slip). Uses `path.relative` rather than `startsWith` so a sibling
 * directory that shares a prefix (`/tmp/dest` vs `/tmp/dest-evil`) cannot
 * sneak through, and so a relative `outputDir` is compared after resolve.
 *
 * Returns the absolute destination path when the entry stays inside the
 * output directory; throws otherwise.
 */
export const resolveSafeZipEntry = (outputDir: string, entryName: string): string => {
    const resolvedOut = path.resolve(outputDir);
    const resolved = path.resolve(resolvedOut, entryName);
    const relative = path.relative(resolvedOut, resolved);
    if (relative === "") {
        return resolved;
    }
    if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
        throw new Error(`Refusing to extract entry outside output dir: ${escapeControlCharacters(entryName)}`);
    }
    return resolved;
};

/**
 * Refuse a destination whose path below `root` runs through, or ends at, a
 * symbolic link already on disk: the write would follow it out of `root`
 * (#675). `root` itself may be a link; the caller chose it.
 */
const assertNoLinkBelow = async (root: string, destination: string, entryName: string): Promise<void> => {
    let current = root;
    for (const segment of path.relative(root, destination).split(path.sep)) {
        current = path.join(current, segment);
        let stat;
        try {
            stat = await fs.lstat(current);
        } catch (err) {
            // Nothing on disk from here down, or a file where a folder is
            // needed, which the write itself then refuses.
            if ((err as NodeJS.ErrnoException).code === "ENOENT" || (err as NodeJS.ErrnoException).code === "ENOTDIR") {
                return;
            }
            throw err;
        }
        if (stat.isSymbolicLink()) {
            throw new Error(
                `Refusing to extract ${escapeControlCharacters(entryName)}: ${escapeControlCharacters(path.relative(root, current).split(path.sep).join("/"))} is a symbolic link`,
            );
        }
    }
};

/**
 * Extract every entry of `zip` under its own name into `outputDir`, a
 * directory the caller chose (`unpack`) or a fresh temporary one.
 *
 * Every entry is checked before the first write, so a refused entry leaves
 * no output: a name that leaves the directory, including a stored name JSZip
 * resolves inside it (`../evil.txt` loads as `evil.txt`, keeping the stored
 * name in `unsafeOriginalName`; #676), and a destination under a symbolic
 * link already in the directory. Extraction creates no links, so checking
 * up front covers every write. When a write or an entry's read fails part
 * way, the files and folders this extraction created are removed; a file it
 * overwrote keeps the new bytes.
 */
export const extractZipEntries = async (zip: JSZip, outputDir: string): Promise<void> => {
    const root = path.resolve(outputDir);
    const entries: { name: string; resolved: string; file: JSZip.JSZipObject }[] = [];
    zip.forEach((relativePath, file) => {
        const stored = (file as JSZip.JSZipObject & { unsafeOriginalName?: string }).unsafeOriginalName;
        if (stored !== undefined) {
            resolveSafeZipEntry(root, stored);
        }
        entries.push({ name: relativePath, resolved: resolveSafeZipEntry(root, relativePath), file });
    });
    for (const { name, resolved } of entries) {
        await assertNoLinkBelow(root, resolved, name);
    }
    const created: string[] = [];
    try {
        for (const { resolved, file } of entries) {
            const folder = await fs.mkdir(file.dir ? resolved : path.dirname(resolved), { recursive: true });
            if (folder !== undefined) {
                created.push(folder);
            }
            if (file.dir) {
                continue;
            }
            const content = await file.async("nodebuffer");
            try {
                await fs.writeFile(resolved, content, { flag: "wx" });
                created.push(resolved);
            } catch (err) {
                // Only EEXIST means the file was there before this write.
                if ((err as NodeJS.ErrnoException).code !== "EEXIST") {
                    created.push(resolved);
                    throw err;
                }
                await fs.writeFile(resolved, content);
            }
        }
    } catch (err) {
        // Everything below a folder mkdir created is this extraction's own.
        for (const entry of created.reverse()) {
            await fs.rm(entry, { recursive: true, force: true }).catch(() => undefined);
        }
        throw err;
    }
};

/**
 * Every file under `dir`, by its `/`-separated path relative to `dir`. Walks
 * one directory at a time: `Dirent.parentPath` needs Node 20.12 and a
 * recursive `readdir` Node 20.1, and the package supports Node 20.
 */
const filesUnder = async (dir: string, prefix = ""): Promise<string[]> => {
    const files: string[] = [];
    for (const entry of await fs.readdir(path.join(dir, prefix), { withFileTypes: true })) {
        const name = prefix === "" ? entry.name : `${prefix}/${entry.name}`;
        if (entry.isDirectory()) {
            files.push(...(await filesUnder(dir, name)));
        } else if (entry.isFile()) {
            files.push(name);
        }
    }
    return files;
};

/**
 * Extract the package view of `zip` for validation: each part once, under
 * its part name, with the bytes of the entry `indexByPartName` keeps. An
 * entry the view drops is read but not written, so no filesystem lets it
 * overwrite a kept part, and an unreadable one still fails (#371, #369). An
 * item name the view refuses stops the extraction (#372, #378). Distinct
 * case/NFC part or folder spellings are refused before writing, on every
 * host (#441). The exclusive writes also guard against a part landing on a
 * file an earlier part wrote through additional filesystem folding: a filesystem that
 * folds case or Unicode normalization holds the two names as one file, and
 * the merged bytes are not what the package holds (#373). `unpack` keeps the raw
 * names (`extractZipEntries`).
 */
export const extractPackageParts = async (zip: JSZip, outputDir: string): Promise<void> => {
    const files = Object.values(zip.files).filter((entry) => !entry.dir);
    assertPlaceableEntries(files);
    const kept = indexByPartName(files.map((entry) => [entry.name, entry] as const));
    for (const entry of files) {
        const content = await entry.async("nodebuffer");
        const part = normalizePartName(entry.name);
        if (kept.get(part) !== entry) {
            continue;
        }
        const resolved = resolveSafeZipEntry(outputDir, part);
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        try {
            // Each part name is written once, so the file exists only when
            // the filesystem holds this name and an earlier part's as one.
            await fs.writeFile(resolved, content, { flag: "wx" });
        } catch (err) {
            if ((err as NodeJS.ErrnoException).code === "EEXIST") {
                throw new Error(
                    `Cannot write part ${JSON.stringify(part)}: this filesystem already holds another part under the same file name`,
                    {
                        cause: err,
                    },
                );
            }
            throw err;
        }
    }
    // Such a filesystem can also merge two parts' folders: the later file
    // then lands in the folder the earlier part named, under a name the
    // package does not hold. Every part must be found under its own (#376).
    const held = new Set(await filesUnder(outputDir));
    for (const part of kept.keys()) {
        if (!held.has(part)) {
            throw new Error(`Cannot write part ${JSON.stringify(part)}: this filesystem holds it under another name`);
        }
    }
};
