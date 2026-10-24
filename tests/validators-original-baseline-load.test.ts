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
 * validator run, not once per XSD-failed part. `JSZip.loadAsync` is spied on
 * and the calls whose argument is the original's exact bytes are counted.
 *
 * #503 — and only once per run: a re-run on the same validator, or a direct
 * `validateAgainstXsd()` call, reads the `--original` as it is now.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { PPTXSchemaValidator } from "../src/scripts/office/validators/pptx";

import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

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

/** Write `parts` as an unpacked package under `dir`. */
async function unpackTo(dir: string, parts: Readonly<Record<string, string>>): Promise<void> {
    for (const [name, content] of Object.entries(parts)) {
        const file = path.join(dir, name);
        await fs.mkdir(path.dirname(file), { recursive: true });
        await fs.writeFile(file, content);
    }
}

/** MINIMAL_PARTS with an element the schema rejects in place of the document's run. */
const DEFECTIVE_PARTS: Readonly<Record<string, string>> = {
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:r><w:t>Hello</w:t></w:r>", "<w:bogus/>"),
};

const xsdErrorPaths = (issues: readonly { code?: string; path?: string; severity: string }[]): string[] =>
    issues.flatMap((issue) => (issue.code === "xsd-error" && issue.severity === "error" ? [issue.path ?? ""] : []));

describe("original archive lifetime (#503)", () => {
    it("subtracts the current --original on a second validate() run, not the previous run's", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, DEFECTIVE_PARTS);
            const originalFile = path.join(dir, "original.docx");
            await fs.writeFile(originalFile, await buildDocx(DEFECTIVE_PARTS));
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            // The original holds the same defect, so the baseline hides it.
            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual([]);

            // The original no longer holds it: the defect is new and must surface.
            await fs.writeFile(originalFile, await buildDocx());
            const rerun = await validator.validate();
            expect(xsdErrorPaths(rerun.issues)).toEqual(["word/document.xml"]);
            expect(rerun.valid).toBe(false);
        });
    }, 60_000);

    it("reads an --original created after a run that found it missing", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, DEFECTIVE_PARTS);
            const originalFile = path.join(dir, "original.docx");
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual(["word/document.xml"]);

            await fs.writeFile(originalFile, await buildDocx(DEFECTIVE_PARTS));
            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual([]);
        });
    }, 60_000);

    it("recovers on the next run once an --original JSZip refused is replaced", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, DEFECTIVE_PARTS);
            const originalFile = path.join(dir, "original.docx");
            await fs.writeFile(originalFile, "not a zip archive");
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            await expect(validator.validate()).rejects.toThrow(/end of central directory/i);

            await fs.writeFile(originalFile, await buildDocx(DEFECTIVE_PARTS));
            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual([]);
        });
    }, 60_000);

    it("reads the current --original on each direct validateAgainstXsd() call", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, DEFECTIVE_PARTS);
            const originalFile = path.join(dir, "original.docx");
            await fs.writeFile(originalFile, await buildDocx(DEFECTIVE_PARTS));
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            expect(xsdErrorPaths((await validator.validateAgainstXsd()).issues)).toEqual([]);

            await fs.writeFile(originalFile, await buildDocx());
            expect(xsdErrorPaths((await validator.validateAgainstXsd()).issues)).toEqual(["word/document.xml"]);
        });
    }, 60_000);

    it("counts the current --original's paragraphs after a re-run", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, MINIMAL_PARTS);
            const originalFile = path.join(dir, "original.docx");
            await fs.writeFile(originalFile, await buildDocx());
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            await validator.validate();
            expect((await validator.compareParagraphCounts()).original).toBe(1);

            const paragraph = "<w:p><w:r><w:t>More</w:t></w:r></w:p>";
            await fs.writeFile(
                originalFile,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:body>", `<w:body>${paragraph}${paragraph}`),
                }),
            );
            await validator.validate();
            expect((await validator.compareParagraphCounts()).original).toBe(3);
        });
    }, 60_000);

    it("subtracts the current --original on a second PPTX validate() run", async () => {
        await withTempDir(async (dir) => {
            const defective = {
                ...MINIMAL_PPTX_PARTS,
                // Only parts directly under ppt/ are schema-validated.
                "ppt/presentation.xml": MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!.replace(
                    "</p:presentation>",
                    "<p:bogus/></p:presentation>",
                ),
            };
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, defective);
            const originalFile = path.join(dir, "original.pptx");
            await fs.writeFile(originalFile, await buildPptx(defective));
            const validator = new PPTXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual([]);

            await fs.writeFile(originalFile, await buildPptx());
            expect(xsdErrorPaths((await validator.validate()).issues)).toEqual(["ppt/presentation.xml"]);
        });
    }, 60_000);

    it("still loads the --original once within a run that also counts paragraphs", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await unpackTo(unpackedDir, DEFECTIVE_PARTS);
            const originalFile = path.join(dir, "original.docx");
            const originalBytes = await buildDocx();
            await fs.writeFile(originalFile, originalBytes);
            const validator = new DOCXSchemaValidator({ unpackedDir, originalFile, profile: "lenient" });

            const load = vi.spyOn(JSZip, "loadAsync");
            let loads = 0;
            try {
                await validator.validate();
                loads = originalLoads(load, originalBytes);
            } finally {
                load.mockRestore();
            }
            expect(loads).toBe(1);
        });
    }, 60_000);
});
