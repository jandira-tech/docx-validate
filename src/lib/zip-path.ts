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
