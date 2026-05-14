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
import JSZip from "jszip";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

async function extractAll(zip: JSZip, outputPath: string): Promise<void> {
    const entries: Array<{ name: string; file: JSZip.JSZipObject }> = [];
    zip.forEach((relativePath, file) => {
        entries.push({ name: relativePath, file });
    });
    for (const { name, file } of entries) {
        const target = path.join(outputPath, name);
        const resolved = path.resolve(target);
        if (!resolved.startsWith(`${outputPath}${path.sep}`) && resolved !== outputPath) {
            throw new Error(`Refusing to extract entry outside output dir: ${name}`);
        }
        if (file.dir) {
            await fs.mkdir(resolved, { recursive: true });
            continue;
        }
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        const data = await file.async("nodebuffer");
        await fs.writeFile(resolved, data);
    }
}

async function main() {
    const targetPath = process.argv[2];
    if (!targetPath) {
        console.error("Usage: tsx scripts/repair-and-overwrite.ts <path-to-docx>");
        process.exit(1);
    }

    const resolved = path.resolve(targetPath);
    const buf = await fs.readFile(resolved);
    const zip = await JSZip.loadAsync(buf);

    await withTempDir(async (tempDir) => {
        // 1. Extract
        await extractAll(zip, tempDir);

        // 2. Repair
        const validator = new DOCXSchemaValidator({
            unpackedDir: tempDir,
            profile: "strict",
        });
        const repairs = await validator.repair();
        console.log(`Repairs applied: ${repairs}`);

        // 3. Repack
        const outZip = new JSZip();
        async function addDir(dir: string, root: string) {
            const entries = await fs.readdir(dir, { withFileTypes: true });
            for (const entry of entries) {
                const full = path.join(dir, entry.name);
                const relative = path.relative(root, full);
                if (entry.isDirectory()) {
                    outZip.folder(relative);
                    await addDir(full, root);
                } else {
                    const data = await fs.readFile(full);
                    outZip.file(relative, data);
                }
            }
        }
        await addDir(tempDir, tempDir);

        const output = await outZip.generateAsync({
            type: "nodebuffer",
            compression: "DEFLATE",
        });
        await fs.writeFile(resolved, output);
        console.log(`Overwritten: ${resolved}`);
    });
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
