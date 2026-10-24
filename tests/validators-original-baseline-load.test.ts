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
 *
 * #520 — and each run keeps the `--original` it loaded: an overlapping
 * `validate()` on the same validator must never replace the baseline an
 * in-flight run is still subtracting against.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationResult } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator, type BaseSchemaValidatorOptions } from "../src/scripts/office/validators/base";
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
    }, 60_000);
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

interface Deferred {
    promise: Promise<void>;
    resolve: () => void;
}

const deferred = (): Deferred => {
    let resolve!: () => void;
    const promise = new Promise<void>((r) => {
        resolve = r;
    });
    return { promise, resolve };
};

/** `validate()` verdict reduced to what `--original` baseline subtraction controls. */
const baselineSummary = (result: ValidationResult): { valid: boolean; xsdErrors: string[] } => ({
    valid: result.valid,
    xsdErrors: result.issues
        .flatMap((issue) =>
            issue.code === "xsd-error" && issue.severity === "error" && issue.path !== undefined ? [`xsd-error@${issue.path}`] : [],
        )
        .sort(),
});

const DEFECTIVE_DOCX_PARTS: Readonly<Record<string, string>> = {
    ...DEFECTIVE_PARTS,
    "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace("</w:styles>", "<w:bogus/></w:styles>"),
};

const DEFECTIVE_PPTX_PARTS: Readonly<Record<string, string>> = {
    ...MINIMAL_PPTX_PARTS,
    "ppt/presentation.xml": MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!.replace("</p:presentation>", "<p:bogus/></p:presentation>"),
    // A second schema-mapped part under ppt/ so the walk subtracts twice.
    "ppt/z-extra.xml": MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!.replace("</p:presentation>", "<p:bogus/></p:presentation>"),
};

interface OverlapSpec {
    readonly extension: "docx" | "pptx";
    /** Part whose XSD failure loads the baseline and trips the pause gate. */
    readonly firstPart: string;
    /** Part that fails XSD after the gate, so the swapped baseline shows. */
    readonly secondPart: string;
    readonly minimal: Readonly<Record<string, string>>;
    readonly defective: Readonly<Record<string, string>>;
    readonly cleanSecondPart: string;
    readonly build: (parts: Readonly<Record<string, string>>) => Promise<Buffer>;
    readonly control: (options: BaseSchemaValidatorOptions) => { validate(): Promise<ValidationResult> };
    readonly gated: (options: BaseSchemaValidatorOptions, paused: Deferred, resume: Deferred) => { validate(): Promise<ValidationResult> };
}

interface OverlapOutcome {
    expectedA: { valid: boolean; xsdErrors: string[] };
    expectedB: { valid: boolean; xsdErrors: string[] };
    runA: { valid: boolean; xsdErrors: string[] };
    runB: { valid: boolean; xsdErrors: string[] };
}

/**
 * Review AO's barrier pattern (#520): run A pauses after the first failing
 * part has consulted the original; the `--original` bytes are then replaced
 * and run B executes to completion on the SAME validator; A resumes and both
 * results are compared against fresh-validator controls.
 *
 * - `remove-defects`: original A holds both defects, then a clean original
 *   replaces it — A must still subtract both.
 * - `add-defects`: original A lacks the second part's defect, then a fully
 *   defective original replaces it — A must still report the second part's
 *   error as new.
 */
async function overlappingBaselineRun(dir: string, spec: OverlapSpec, scenario: "remove-defects" | "add-defects"): Promise<OverlapOutcome> {
    const unpackedDir = path.join(dir, "unpacked");
    await unpackTo(unpackedDir, spec.defective);

    const originalFile = path.join(dir, `original.${spec.extension}`);
    const firstOriginal = scenario === "remove-defects" ? spec.defective : { ...spec.defective, [spec.secondPart]: spec.cleanSecondPart };
    await fs.writeFile(originalFile, await spec.build(firstOriginal));

    // Fresh-validator control for A, seeing the same original A will see.
    const expectedA = await spec.control({ unpackedDir, originalFile, profile: "lenient" }).validate();

    const paused = deferred();
    const resume = deferred();
    const validator = spec.gated({ unpackedDir, originalFile, profile: "lenient" }, paused, resume);
    const runA = validator.validate();
    await paused.promise;

    await fs.writeFile(originalFile, await spec.build(scenario === "remove-defects" ? spec.minimal : spec.defective));
    // Fresh-validator control for B, seeing the replaced original.
    const expectedB = await spec.control({ unpackedDir, originalFile, profile: "lenient" }).validate();

    const runB = validator.validate().then((result) => {
        resume.resolve();
        return result;
    });
    const [runAResult, runBResult] = await Promise.all([runA, runB]);
    return {
        expectedA: baselineSummary(expectedA),
        expectedB: baselineSummary(expectedB),
        runA: baselineSummary(runAResult),
        runB: baselineSummary(runBResult),
    };
}

const docxOverlap: OverlapSpec = {
    extension: "docx",
    firstPart: "word/document.xml",
    secondPart: "word/styles.xml",
    minimal: MINIMAL_PARTS,
    defective: DEFECTIVE_DOCX_PARTS,
    cleanSecondPart: MINIMAL_PARTS["word/styles.xml"]!,
    build: buildDocx,
    control: (options) => new DOCXSchemaValidator(options),
    gated: (options, paused, resume) =>
        new (class extends DOCXSchemaValidator {
            didPause = false;
            override async validateFileAgainstXsd(file: string) {
                const outcome = await super.validateFileAgainstXsd(file);
                if (!this.didPause && this.relPath(file).split(path.sep).join("/") === docxOverlap.firstPart) {
                    this.didPause = true;
                    paused.resolve();
                    await resume.promise;
                }
                return outcome;
            }
        })(options),
};

const pptxOverlap: OverlapSpec = {
    extension: "pptx",
    firstPart: "ppt/presentation.xml",
    secondPart: "ppt/z-extra.xml",
    minimal: MINIMAL_PPTX_PARTS,
    defective: DEFECTIVE_PPTX_PARTS,
    cleanSecondPart: MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!,
    build: buildPptx,
    control: (options) => new PPTXSchemaValidator(options),
    gated: (options, paused, resume) =>
        new (class extends PPTXSchemaValidator {
            didPause = false;
            override async validateFileAgainstXsd(file: string) {
                const outcome = await super.validateFileAgainstXsd(file);
                if (!this.didPause && this.relPath(file).split(path.sep).join("/") === pptxOverlap.firstPart) {
                    this.didPause = true;
                    paused.resolve();
                    await resume.promise;
                }
                return outcome;
            }
        })(options),
};

describe("overlapping runs keep their own --original (#520)", () => {
    it.each(["remove-defects", "add-defects"] as const)(
        "DOCX run A keeps the baseline it loaded when run B overlaps (%s)",
        async (scenario) => {
            await withTempDir(async (dir) => {
                const outcome = await overlappingBaselineRun(dir, docxOverlap, scenario);
                expect(outcome.runA).toEqual(outcome.expectedA);
                expect(outcome.runB).toEqual(outcome.expectedB);
            });
        },
        60_000,
    );

    it.each(["remove-defects", "add-defects"] as const)(
        "PPTX run A keeps the baseline it loaded when run B overlaps (%s)",
        async (scenario) => {
            await withTempDir(async (dir) => {
                const outcome = await overlappingBaselineRun(dir, pptxOverlap, scenario);
                expect(outcome.runA).toEqual(outcome.expectedA);
                expect(outcome.runB).toEqual(outcome.expectedB);
            });
        },
        60_000,
    );
});
