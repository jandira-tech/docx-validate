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
 * #481 — the `--original` baseline archive must be read and loaded once per
 * validator, not once per XSD-failed part. `JSZip.loadAsync` is spied on and
 * the calls whose argument is the original's exact bytes are counted.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";

import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

class HarnessValidator extends BaseSchemaValidator {
    originalErrors(xmlFile: string): Promise<Set<string>> {
        return this._getOriginalFileErrors(xmlFile);
    }
}

/** `JSZip.loadAsync` calls whose first argument holds exactly `bytes`. */
const originalLoads = (load: { mock: { calls: readonly (readonly unknown[])[] } }, bytes: Uint8Array): number =>
    load.mock.calls.filter(
        ([data]) => data instanceof Uint8Array && data.length === bytes.length && Buffer.from(data).equals(Buffer.from(bytes)),
    ).length;

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("original archive load count (#481)", () => {
    it("loads the --original archive once across per-part baseline lookups", async () => {
        await withTempDir(async (dir) => {
            const originalPath = path.join(dir, "original.docx");
            const originalBytes = await buildDocx();
            await fs.writeFile(originalPath, originalBytes);

            const unpackedDir = path.join(dir, "unpacked");
            const validator = new HarnessValidator({ unpackedDir, originalFile: originalPath });
            const load = vi.spyOn(JSZip, "loadAsync");
            let loads: number;
            try {
                for (const part of ["word/document.xml", "word/styles.xml", "_rels/.rels"]) {
                    await validator.originalErrors(path.join(unpackedDir, part));
                }
                loads = originalLoads(load, originalBytes);
            } finally {
                // mockRestore clears mock.calls — count before restoring.
                load.mockRestore();
            }
            expect(loads!).toBe(1);
        });
    });

    it("loads the --original archive once per validate() run across failing parts", async () => {
        await withTempDir(async (dir) => {
            const originalPath = path.join(dir, "original.docx");
            const originalBytes = await buildDocx();
            await fs.writeFile(originalPath, originalBytes);

            // Two XSD-invalid parts: each failed part consults the original
            // archive for baseline subtraction under lenient.
            const modifiedPath = path.join(dir, "modified.docx");
            await fs.writeFile(
                modifiedPath,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml":
                        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                        `<w:document xmlns:w="${W_NS}"><w:body><w:bogus/></w:body></w:document>`,
                    "word/styles.xml":
                        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                        `<w:styles xmlns:w="${W_NS}"><w:style w:type="bogus" w:styleId="X"><w:name w:val="X"/></w:style></w:styles>`,
                }),
            );

            const load = vi.spyOn(JSZip, "loadAsync");
            let result;
            let loads = 0;
            try {
                result = await validate(modifiedPath, {
                    original: originalPath,
                    author: "Test Author",
                    profile: "lenient",
                });
                loads = originalLoads(load, originalBytes);
            } finally {
                load.mockRestore();
            }
            // Both parts must really have failed XSD — otherwise the count
            // assertion below could pass vacuously.
            expect(result!.issues.filter((issue) => issue.code === "xsd-summary")).toHaveLength(2);
            expect(loads).toBe(1);
        });
    });
});
