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
        throw new Error(`Refusing to extract entry outside output dir: ${entryName}`);
    }
    return resolved;
};

export const extractZipEntries = async (zip: JSZip, outputDir: string): Promise<void> => {
    const entries: { name: string; file: JSZip.JSZipObject }[] = [];
    zip.forEach((relativePath, file) => {
        entries.push({ name: relativePath, file });
    });
    for (const { name, file } of entries) {
        const resolved = resolveSafeZipEntry(outputDir, name);
        if (file.dir) {
            await fs.mkdir(resolved, { recursive: true });
            continue;
        }
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        await fs.writeFile(resolved, await file.async("nodebuffer"));
    }
};

/**
 * Extract the package view of `zip` for validation: each part once, under
 * its part name, with the bytes of the entry `indexByPartName` keeps. An
 * entry the view drops is read but not written, so no filesystem lets it
 * overwrite a kept part, and an unreadable one still fails (#371, #369). An
 * item name the view refuses stops the extraction (#372, #378), and so does a
 * part that would land on a file an earlier part wrote: a filesystem that
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
    const held = new Set(
        (await fs.readdir(outputDir, { recursive: true, withFileTypes: true }))
            .filter((entry) => entry.isFile())
            .map((entry) => path.relative(outputDir, path.join(entry.parentPath, entry.name)).split(path.sep).join("/")),
    );
    for (const part of kept.keys()) {
        if (!held.has(part)) {
            throw new Error(`Cannot write part ${JSON.stringify(part)}: this filesystem holds it under another name`);
        }
    }
};
