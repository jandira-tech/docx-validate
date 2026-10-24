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
import { describe, expect, it, vi } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator, withoutRestatedIgnorablePrefixes } from "../src/scripts/office/validators/base";

import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

class HarnessValidator extends BaseSchemaValidator {
    schemaMappingKey(part: string): string | null {
        return this._schemaMappingKey(path.join(this.unpackedDir, part));
    }

    originalErrors(xmlFile: string): Promise<Set<string>> {
        return this._getOriginalFileErrors(xmlFile);
    }
}

const RELS_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

async function writeFile(p: string, content: string): Promise<void> {
    await fs.mkdir(path.dirname(p), { recursive: true });
    await fs.writeFile(p, content, "utf-8");
}

describe("direct relationships outside legacy lenient (#434)", () => {
    it.each([
        ["validateFileReferences", '<Relationship Id="nested" Type="urn:image" Target="../customXml/missing.xml"/>', "rels-broken"],
        ["validateRelationshipElements", '<Relationship Id="nested"/>', "rels-empty-element"],
        ["validateAllRelationshipIds", '<Relationship Id="rId1" Type="urn:styles" Target="styles.xml"/>', "rels-id-duplicate"],
    ] as const)("%s ignores nested entries and preserves lenient behavior", async (method, nested, code) => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<x:Wrap xmlns:x="urn:x">${nested}</x:Wrap></Relationships>`,
                ),
            };
            for (const [part, xml] of Object.entries(parts)) await writeFile(path.join(dir, part), xml);
            const strict = await new HarnessValidator({ unpackedDir: dir, profile: "strict" })[method]();
            expect(strict.issues.filter((issue) => issue.code === code)).toEqual([]);
            const lenient = await new HarnessValidator({ unpackedDir: dir, profile: "lenient" })[method]();
            expect(lenient.issues).toContainEqual(expect.objectContaining({ code, severity: "error" }));
        });
    });
});

describe("BaseSchemaValidator", () => {
    describe("original-file extraction", () => {
        it("rejects an out-of-root part before reading or writing its ZIP payload", async () => {
            await withTempDir(async (dir) => {
                const originalFile = path.join(dir, "original.docx");
                await fs.writeFile(originalFile, await new JSZip().generateAsync({ type: "nodebuffer" }));
                const readPayload = vi.fn().mockResolvedValue(Buffer.from("malicious"));
                const zip = new JSZip();
                // JSZip normally sanitizes traversal names. Model a ZIP provider
                // returning one so this extraction boundary is tested directly.
                const entry = { async: readPayload } as unknown as JSZip.JSZipObject;
                zip.files["../escaped.xml"] = entry;
                const load = vi.spyOn(JSZip, "loadAsync").mockResolvedValue(zip);
                const escaped = path.join(dir, "escaped.xml");
                const validator = new HarnessValidator({ unpackedDir: path.join(dir, "unpacked"), originalFile });
                try {
                    await expect(validator.originalErrors(escaped)).rejects.toThrow("Refusing to extract entry outside output dir");
                    expect(readPayload).not.toHaveBeenCalled();
                    await expect(fs.access(escaped)).rejects.toThrow();
                } finally {
                    load.mockRestore();
                }
            });
        });

        it("keeps nested part paths for schema selection", async () => {
            await withTempDir(async (dir) => {
                const originalFile = path.join(dir, "original.docx");
                const zip = new JSZip();
                zip.file("word/theme/theme1.xml", '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>');
                await fs.writeFile(originalFile, await zip.generateAsync({ type: "nodebuffer" }));
                const validator = new HarnessValidator({ unpackedDir: path.join(dir, "unpacked"), originalFile });
                const errors = await validator.originalErrors(path.join(dir, "unpacked/word/theme/theme1.xml"));
                expect(errors.size).toBeGreaterThan(0);
                expect([...errors].some((message) => message.includes("themeElements"))).toBe(true);
            });
        });
    });

    describe("validateXml", () => {
        it("flags malformed XML", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<not-closed`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateXml();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "xml-syntax")).toBe(true);
            });
        });

        it("passes on well-formed XML", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateXml();
                expect(result.valid).toBe(true);
                expect(result.issues).toEqual([]);
            });
        });
    });

    describe("validateUniqueIds", () => {
        it("detects duplicate file-scoped IDs (e.g. duplicate w:comment id)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS}>
            <w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A"/>
            <w:comment w:id="0" w:author="B" w:date="2026-01-02T00:00:00Z" w:initials="B"/>
          </w:comments>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateUniqueIds();
                expect(result.valid).toBe(false);
                expect(result.issues.length).toBe(1);
                expect(result.issues[0].code).toBe("id-duplicate-file");
                expect(result.issues[0].message).toContain("'0'");
            });
        });

        it("passes when all IDs are unique", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS}>
            <w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A"/>
            <w:comment w:id="1" w:author="B" w:date="2026-01-02T00:00:00Z" w:initials="B"/>
          </w:comments>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateUniqueIds();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateFileReferences", () => {
        describe("OPC part-name equivalence (#346, #393)", { timeout: 120_000 }, () => {
            const cases = [
                {
                    name: "WORD folder with lowercase relationship targets",
                    parts: Object.fromEntries(Object.entries(MINIMAL_PARTS).map(([name, xml]) => [name.replace(/^word\//, "WORD/"), xml])),
                    expected: [],
                },
                {
                    name: "percent-encoded unreserved character in the main target",
                    parts: {
                        ...MINIMAL_PARTS,
                        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace("word/document.xml", "word/%64ocument.xml"),
                    },
                    expected: [],
                },
                {
                    name: "case and percent variants in a relative target",
                    parts: {
                        ...MINIMAL_PARTS,
                        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                            "styles.xml",
                            "%73TYLES.xml",
                        ),
                    },
                    expected: [],
                },
                {
                    name: "equivalent package-absolute target",
                    parts: {
                        ...MINIMAL_PARTS,
                        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                            "styles.xml",
                            "/WORD/%73tyles.xml",
                        ),
                    },
                    expected: [],
                },
                {
                    name: "target naming no part",
                    parts: {
                        ...MINIMAL_PARTS,
                        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace("styles.xml", "missing.xml"),
                    },
                    expected: [
                        { code: "rels-broken", path: "word/_rels/document.xml.rels", message: "Broken reference to missing.xml" },
                        { code: "rels-unreferenced", path: "word/styles.xml", message: "Unreferenced file: word/styles.xml" },
                        {
                            code: "rels-target-missing",
                            path: "word/_rels/document.xml.rels",
                            message: "Relationship target 'missing.xml' resolves to 'word/missing.xml' which does not exist",
                        },
                    ],
                },
                {
                    // An escaped reserved character names another part (#297, #344).
                    name: "reserved escape in the main target",
                    parts: {
                        ...MINIMAL_PARTS,
                        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace("word/document.xml", "word%2Fdocument.xml"),
                    },
                    expected: [
                        { code: "rels-broken", path: "_rels/.rels", message: "Broken reference to word%2Fdocument.xml" },
                        { code: "rels-unreferenced", path: "word/document.xml", message: "Unreferenced file: word/document.xml" },
                        {
                            code: "rels-target-missing",
                            path: "_rels/.rels",
                            message: "Relationship target 'word%2Fdocument.xml' resolves to 'word%2Fdocument.xml' which does not exist",
                        },
                    ],
                },
                {
                    // _RELS/.RELS names the package relationships part (#364).
                    name: "package relationships stored as _RELS/.RELS",
                    parts: Object.fromEntries(
                        Object.entries(MINIMAL_PARTS).map(([name, xml]) => [name === "_rels/.rels" ? "_RELS/.RELS" : name, xml]),
                    ),
                    expected: [],
                },
                {
                    name: "part named by no relationship",
                    parts: { ...MINIMAL_PARTS, "word/orphan.xml": "<extra/>" },
                    expected: [{ code: "rels-unreferenced", path: "word/orphan.xml", message: "Unreferenced file: word/orphan.xml" }],
                },
            ];

            // These are filesystem integration tests: exercise the public API
            // against generated packages rather than host-dependent path mocks.
            for (const profile of ["lenient", "strict"] as const) {
                for (const input of ["directory", "packed"] as const) {
                    for (const { name, parts, expected } of cases) {
                        it(`${profile}, ${input}: ${name}`, async () => {
                            await withTempDir(async (dir) => {
                                const source = path.join(dir, "specimen.docx");
                                const bytes = input === "packed" ? await buildDocx(parts) : undefined;
                                if (bytes) {
                                    await fs.writeFile(source, bytes);
                                } else {
                                    for (const [entry, xml] of Object.entries(parts)) {
                                        await writeFile(path.join(source, entry), xml);
                                    }
                                }

                                const result = await validate(source, { profile });
                                const references = result.issues
                                    .filter((issue) =>
                                        ["rels-broken", "rels-unreferenced", "rels-target-missing"].includes(issue.code ?? ""),
                                    )
                                    .map(({ code, path, message }) => ({ code, path, message }));
                                expect(references).toEqual(expected);
                                expect(result.valid).toBe(expected.length === 0);

                                if (bytes) {
                                    expect(await fs.readFile(source)).toEqual(bytes);
                                } else {
                                    for (const [entry, xml] of Object.entries(parts)) {
                                        expect(await fs.readFile(path.join(source, entry), "utf-8")).toBe(xml);
                                    }
                                }
                            });
                        });
                    }
                }
            }
        });

        it("detects broken Target references in .rels", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles-missing.xml"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateFileReferences();
                expect(result.valid).toBe(false);
                const broken = result.issues.find((i) => i.code === "rels-broken");
                expect(broken).toBeDefined();
                expect(broken!.message).toContain("word/styles-missing.xml");
            });
        });

        it("passes when all references resolve and no orphan files exist", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateFileReferences();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateAgainstXsd", () => {
        it("validates a known-good .rels file against opc-relationships.xsd", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const outcome = await v.validateFileAgainstXsd(path.join(dir, "_rels", ".rels"));
                expect(outcome.valid).toBe(true);
                expect(outcome.errors.size).toBe(0);
            });
        });

        it("flags a known-bad .rels file (missing required Target attribute)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const outcome = await v.validateFileAgainstXsd(path.join(dir, "_rels", ".rels"));
                expect(outcome.valid).toBe(false);
                expect(outcome.errors.size).toBeGreaterThan(0);
            });
        });

        it("returns null (skip) for a file without a registered schema", async () => {
            await withTempDir(async (dir) => {
                const p = path.join(dir, "random.xml");
                await writeFile(p, `<?xml version="1.0"?><something/>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const outcome = await v.validateFileAgainstXsd(p);
                expect(outcome.valid).toBeNull();
            });
        });

        it("validates many files in one validator without leaking state across files", async () => {
            // Regression: the template-tag predicate used `.test()` against a
            // `g`-flagged regex literal shared across all files, which advances
            // `lastIndex` between calls and silently mis-skips matches on
            // subsequent files. With many files in a row containing template tags,
            // the second/third/etc. file would get inconsistent stripping. We use
            // four files here so the pattern would have to reset its state more
            // than once for the test to pass.
            await withTempDir(async (dir) => {
                const filenames = ["one", "two", "three", "four"];
                for (const name of filenames) {
                    await writeFile(
                        path.join(dir, "_rels", `${name}.rels`),
                        `${RELS_HEADER}
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <!-- {{template-tag-${name}}} -->
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
              </Relationships>`,
                    );
                }
                const v = new HarnessValidator({ unpackedDir: dir });
                for (const name of filenames) {
                    const outcome = await v.validateFileAgainstXsd(path.join(dir, "_rels", `${name}.rels`));
                    expect(outcome.valid).toBe(true);
                    expect(outcome.errors.size).toBe(0);
                }
            });
        });

        it("caches parsed XSDs across calls (second call hits cache)", async () => {
            // Sanity check that the static cache is doing its job — second call to
            // _validateSingleFileXsd should not re-parse the schema. We can't
            // observe the cache directly, but we can at least confirm two calls
            // return the same outcome shape and don't throw.
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const a = await v.validateFileAgainstXsd(path.join(dir, "_rels", ".rels"));
                const b = await v.validateFileAgainstXsd(path.join(dir, "_rels", ".rels"));
                expect(a.valid).toBe(true);
                expect(b.valid).toBe(true);
            });
        });

        it("turns missing-XSD-on-disk into a per-file validation error (Python parity)", async () => {
            // Per Python's bare-except behaviour, schema-load issues surface as a
            // per-file error rather than throwing — that lets the
            // IGNORED_VALIDATION_ERRORS list filter out known-noisy schemas.
            // For loud failure on a broken engine, callers should invoke
            // BaseSchemaValidator.assertXsdValidationAvailable() at startup.
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`,
                );
                const v = new HarnessValidator({
                    unpackedDir: dir,
                    schemasDir: path.join(dir, "no-such-schemas-dir"),
                });
                const outcome = await v.validateFileAgainstXsd(path.join(dir, "_rels", ".rels"));
                expect(outcome.valid).toBe(false);
                expect(outcome.errors.size).toBeGreaterThan(0);
            });
        });

        it("never subtracts schema-load failures (#289)", async () => {
            // Infrastructure findings (xsd-schema-load-failed, engine
            // failures) are never part of the --original baseline
            // subtraction: they always surface as error-severity issues and
            // make the part invalid, under every profile.
            await withTempDir(async (dir) => {
                const file = path.join(dir, "word", "document.xml");
                await writeFile(file, MINIMAL_PARTS["word/document.xml"]!);
                const original = path.join(dir, "original.docx");
                await fs.writeFile(original, await buildDocx());
                const schemasDir = path.join(dir, "empty-schemas");
                await fs.mkdir(schemasDir);
                const validator = new BaseSchemaValidator({
                    unpackedDir: dir,
                    originalFile: original,
                    schemasDir,
                    profile: "word-valid",
                });
                expect((await validator.validateFileAgainstXsd(file)).valid).toBe(false);
            });
        });
    });

    describe("assertXsdValidationAvailable", () => {
        it("returns silently when the wasm-backed XSD engine works on this host", async () => {
            await expect(BaseSchemaValidator.assertXsdValidationAvailable()).resolves.toBeUndefined();
        });
    });

    describe("validateNamespaces", () => {
        it("flags a prefix in mc:Ignorable that is not declared on the root", async () => {
            await withTempDir(async (dir) => {
                // mc:Ignorable references "w14" but there is no xmlns:w14 declaration.
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?>` +
                        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
                        ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
                        ` mc:Ignorable="w14">` +
                        `<w:body/></w:document>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateNamespaces();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "ignorable-undeclared" && i.message.includes("w14"))).toBe(true);
            });
        });

        it("passes when all mc:Ignorable prefixes are declared on the root", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?>` +
                        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
                        ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
                        ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
                        ` mc:Ignorable="w14">` +
                        `<w:body/></w:document>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateNamespaces();
                expect(result.valid).toBe(true);
            });
        });

        it("passes when the mc:Ignorable prefix is declared on a child element (lxml nsmap parity)", async () => {
            // lxml's nsmap propagates xmlns:* declarations from all descendants.
            // TS must match: a prefix declared only on a child should still be
            // treated as "declared" for the purposes of mc:Ignorable validation.
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?>` +
                        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
                        ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
                        ` mc:Ignorable="w14">` +
                        // xmlns:w14 is on a child element only — not on the root
                        `<w:body><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/></w:body>` +
                        `</w:document>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateNamespaces();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateContentTypes", () => {
        it("flags a <document> root file not declared in [Content_Types].xml", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `</Types>`,
                );
                // document.xml has a <document> root but is NOT listed in Override entries.
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "ct-undeclared-part")).toBe(true);
            });
        });

        it("passes when [Content_Types].xml declares an Override for word/document.xml", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Override PartName="/word/document.xml"` +
                        ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(true);
            });
        });

        it("returns ct-missing when [Content_Types].xml is absent", async () => {
            await withTempDir(async (dir) => {
                // No [Content_Types].xml written.
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("ct-missing");
            });
        });

        it("flags a part covered by neither Override nor Default (ct-uncovered-part)", async () => {
            await withTempDir(async (dir) => {
                // Extensionless media — the recurring real-world case
                // (superdoc editors emit word/media/image_3094991244).
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "media", "image_3094991244"), "\x89PNG-blob");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(false);
                const hits = result.issues.filter((i) => i.code === "ct-uncovered-part");
                expect(hits.length).toBe(1);
                expect(hits[0].path).toBe("word/media/image_3094991244");
            });
        });

        it("does not double-flag uncovered known media extensions (ct-undeclared-ext owns those)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "media", "x.png"), "\x89PNG-blob");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "ct-undeclared-ext")).toBe(true);
                expect(result.issues.some((i) => i.code === "ct-uncovered-part")).toBe(false);
            });
        });

        it("flags a part whose extension shadows an Object.prototype property ('constructor' is not media)", async () => {
            await withTempDir(async (dir) => {
                // `"constructor" in {}` is true — an own-property check is
                // required or evil.constructor sails past the media-extension
                // lookup AND gets suggested with ContentType="undefined".
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "media", "evil.constructor"), "not media");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-uncovered-part" && i.path === "word/media/evil.constructor")).toBe(true);
                expect(result.issues.some((i) => i.code === "ct-undeclared-ext" && i.message.includes("undefined"))).toBe(false);
            });
        });

        it("matches URI-escaped Override PartNames against decoded package paths", async () => {
            await withTempDir(async (dir) => {
                // OPC PartName is a URI; the on-disk path is its decoded form.
                // The comparison must normalize both sides or a validly
                // declared part is flagged ct-uncovered-part.
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Override PartName="/word/document%20with%20space.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "document with space.xml"), "<document/>");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-uncovered-part" && i.path === "word/document with space.xml")).toBe(false);
            });
        });

        it("covers media declared by an OPC case-variant Override PartName in every pass", async () => {
            await withTempDir(async (dir) => {
                // OPC part-name comparison is ASCII case-insensitive, so
                // /Word/Media/IMAGE1.JPEG declares word/media/image1.jpeg.
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Override PartName="/Word/Media/IMAGE1.JPEG" ContentType="image/jpeg"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "media", "image1.jpeg"), "\xff\xd8\xff\xe0");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-undeclared-ext")).toBe(false);
                expect(result.issues.some((i) => i.code === "ct-uncovered-part")).toBe(false);
            });
        });

        it("covers a declarable XML root declared by an OPC case-variant Override PartName", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Override PartName="/Word/Document.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-undeclared-part")).toBe(false);
            });
        });

        it("still reports ct-undeclared-part when only a Default, not an Override, covers a declarable root", async () => {
            await withTempDir(async (dir) => {
                // document.xml is reachable through <Default Extension="xml">
                // but has no Override — the declarable-root check keys on
                // Override coverage, so this stays a defect.
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Default Extension="xml" ContentType="application/xml"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-undeclared-part" && i.path === "word/document.xml")).toBe(true);
            });
        });

        it("resolves an escaped Override PartName for media (percent-decoding)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Override PartName="/word/media/my%20img.jpeg" ContentType="image/jpeg"/>` +
                        `</Types>`,
                );
                await writeFile(path.join(dir, "word", "media", "my img.jpeg"), "\xff\xd8\xff\xe0");
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.issues.some((i) => i.code === "ct-undeclared-ext")).toBe(false);
                expect(result.issues.some((i) => i.code === "ct-uncovered-part")).toBe(false);
            });
        });

        it("treats _rels/.rels as extension 'rels' (OPC last-dot rule, not POSIX dotfile)", async () => {
            await withTempDir(async (dir) => {
                // POSIX path.extname('.rels') is '' — a naive implementation
                // would call the part extensionless and flag it. OPC derives
                // the extension from the last dot, so the Default covers it.
                await writeFile(
                    path.join(dir, "[Content_Types].xml"),
                    `<?xml version="1.0"?>` +
                        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                        `<Default Extension="xml" ContentType="application/xml"/>` +
                        `</Types>`,
                );
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}` + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateContentTypes();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateAllRelationshipIds", () => {
        it("flags a duplicate rId in a .rels file", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                // Two Relationship entries with the same Id.
                await writeFile(
                    path.join(dir, "word", "_rels", "document.xml.rels"),
                    `${RELS_HEADER}` +
                        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateAllRelationshipIds();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-id-duplicate" && i.message.includes("rId1"))).toBe(true);
            });
        });

        it("passes when all rIds in the .rels file are unique", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                await writeFile(
                    path.join(dir, "word", "_rels", "document.xml.rels"),
                    `${RELS_HEADER}` +
                        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
                        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateAllRelationshipIds();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateRelationshipElements", () => {
        const REL_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
        const REL_TYPE = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`;

        it("passes a well-formed self-closing <Relationship>", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` +
                        `<Relationship Id="rId1" Type="${REL_TYPE}" Target="word/document.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateRelationshipElements();
                expect(result.valid).toBe(true);
                expect(result.issues).toEqual([]);
            });
        });

        it("flags a <Relationship> missing required Id attribute", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` +
                        `<Relationship Type="${REL_TYPE}" Target="word/document.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateRelationshipElements();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-empty-element" && i.message.includes("Id"))).toBe(true);
            });
        });

        it("flags a <Relationship> missing required Type and Target attributes", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` + `<Relationship Id="rId1"/>` + `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateRelationshipElements();
                expect(result.valid).toBe(false);
                const issue = result.issues.find((i) => i.code === "rels-empty-element");
                expect(issue).toBeDefined();
                expect(issue!.message).toContain("Type");
                expect(issue!.message).toContain("Target");
            });
        });

        it("flags a <Relationship> with non-whitespace text body content", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` +
                        `<Relationship Id="rId1" Type="${REL_TYPE}" Target="word/document.xml">oops</Relationship>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateRelationshipElements();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-empty-element" && i.message.includes("self-closing"))).toBe(true);
            });
        });

        it("flags a <Relationship> with child element body content", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "_rels", ".rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` +
                        `<Relationship Id="rId1" Type="${REL_TYPE}" Target="word/document.xml"><foo/></Relationship>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateRelationshipElements();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-empty-element" && i.message.includes("self-closing"))).toBe(true);
            });
        });
    });

    describe("rels-missing-sidecar branch in validateAllRelationshipIds", () => {
        const R_NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
        const REL_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
        const REL_TYPE = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`;

        it("flags an XML part that uses r:id but has no .rels sidecar", async () => {
            await withTempDir(async (dir) => {
                // header1.xml uses r:embed but no word/_rels/header1.xml.rels exists.
                await writeFile(
                    path.join(dir, "word", "header1.xml"),
                    `<?xml version="1.0"?><w:hdr ${W_NS} ${R_NS}><w:p><w:r>` +
                        `<w:drawing><a:blip r:embed="rId1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></w:drawing>` +
                        `</w:r></w:p></w:hdr>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateAllRelationshipIds();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-missing-sidecar")).toBe(true);
            });
        });

        it("passes an XML part that has no r:id refs and no .rels sidecar (per OPC §9.3.1)", async () => {
            await withTempDir(async (dir) => {
                // header1.xml has no outgoing r:id references — no sidecar is required.
                await writeFile(
                    path.join(dir, "word", "header1.xml"),
                    `<?xml version="1.0"?><w:hdr ${W_NS}><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateAllRelationshipIds();
                expect(result.valid).toBe(true);
                expect(result.issues.some((i) => i.code === "rels-missing-sidecar")).toBe(false);
            });
        });

        it("passes an XML part that has r:id refs and a matching .rels sidecar", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "header1.xml"),
                    `<?xml version="1.0"?><w:hdr ${W_NS} ${R_NS}><w:p><w:r>` +
                        `<w:drawing><a:blip r:embed="rId1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></w:drawing>` +
                        `</w:r></w:p></w:hdr>`,
                );
                await writeFile(
                    path.join(dir, "word", "_rels", "header1.xml.rels"),
                    `${RELS_HEADER}<Relationships ${REL_NS}>` +
                        `<Relationship Id="rId1" Type="${REL_TYPE}" Target="styles.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateAllRelationshipIds();
                expect(result.valid).toBe(true);
                expect(result.issues.some((i) => i.code === "rels-missing-sidecar")).toBe(false);
            });
        });
    });

    describe("_isStrictXmlFile", () => {
        it("returns true for a file whose root namespace is a Strict OOXML URI", async () => {
            await withTempDir(async (dir) => {
                const STRICT_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?><w:document xmlns:w="${STRICT_NS}"><w:body/></w:document>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir });
                // _isStrictXmlFile is protected — access via any cast for unit testing.
                const isStrict = (v as unknown as Record<string, (f: string) => boolean>)._isStrictXmlFile(
                    path.join(dir, "word", "document.xml"),
                );
                expect(isStrict).toBe(true);
            });
        });

        it("returns false for a file whose root namespace is a Transitional OOXML URI", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><w:document ${W_NS}><w:body/></w:document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const isStrict = (v as unknown as Record<string, (f: string) => boolean>)._isStrictXmlFile(
                    path.join(dir, "word", "document.xml"),
                );
                expect(isStrict).toBe(false);
            });
        });

        it("returns false for a non-existent file (parse error → false, not throw)", async () => {
            await withTempDir(async (dir) => {
                const v = new HarnessValidator({ unpackedDir: dir });
                const isStrict = (v as unknown as Record<string, (f: string) => boolean>)._isStrictXmlFile(
                    path.join(dir, "word", "does-not-exist.xml"),
                );
                expect(isStrict).toBe(false);
            });
        });

        it("returns false for a file with no namespace on the root element", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><document><body/></document>`);
                const v = new HarnessValidator({ unpackedDir: dir });
                const isStrict = (v as unknown as Record<string, (f: string) => boolean>)._isStrictXmlFile(
                    path.join(dir, "word", "document.xml"),
                );
                expect(isStrict).toBe(false);
            });
        });
    });

    describe("validateNoEmptyRelsParts (profile-aware)", () => {
        const PR_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;

        it("flags an empty <Relationships/> as ERROR under strict", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "_rels", "fontTable.xml.rels"), `${RELS_HEADER}\n<Relationships ${PR_NS}/>`);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const result = await v.validateNoEmptyRelsParts();
                expect(result.valid).toBe(false);
                const issue = result.issues.find((i) => i.code === "rels-empty-part");
                expect(issue).toBeDefined();
                expect(issue?.severity).toBe("error");
                expect(issue?.path).toContain("fontTable.xml.rels");
            });
        });

        it("downgrades to WARNING under lenient (still surfaced, doesn't fail validation)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "_rels", "fontTable.xml.rels"), `${RELS_HEADER}\n<Relationships ${PR_NS}/>`);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "lenient" });
                const result = await v.validateNoEmptyRelsParts();
                expect(result.valid).toBe(true);
                const issue = result.issues.find((i) => i.code === "rels-empty-part");
                expect(issue).toBeDefined();
                expect(issue?.severity).toBe("warning");
            });
        });

        it("passes when the rels part has at least one <Relationship> child", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "_rels", "document.xml.rels"),
                    `${RELS_HEADER}\n<Relationships ${PR_NS}>` +
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
                        `</Relationships>`,
                );
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const result = await v.validateNoEmptyRelsParts();
                expect(result.valid).toBe(true);
                expect(result.issues).toEqual([]);
            });
        });

        it("ignores non-rels files entirely", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), `<?xml version="1.0"?><doc/>`);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const result = await v.validateNoEmptyRelsParts();
                expect(result.valid).toBe(true);
                expect(result.issues).toEqual([]);
            });
        });
    });

    describe("repairEmptyRelsParts", () => {
        const PR_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;

        it("deletes the file when <Relationships/> is empty", async () => {
            await withTempDir(async (dir) => {
                const relsFile = path.join(dir, "word", "_rels", "fontTable.xml.rels");
                await writeFile(relsFile, `${RELS_HEADER}\n<Relationships ${PR_NS}/>`);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const repairs = await v.repairEmptyRelsParts();
                expect(repairs).toBe(1);
                await expect(fs.access(relsFile)).rejects.toThrow();
                // Subsequent validation pass must not crash trying to read it.
                const after = await v.validateNoEmptyRelsParts();
                expect(after.valid).toBe(true);
                expect(after.issues).toEqual([]);
            });
        });

        it("leaves a non-empty rels part intact", async () => {
            await withTempDir(async (dir) => {
                const relsFile = path.join(dir, "word", "_rels", "document.xml.rels");
                const content =
                    `${RELS_HEADER}\n<Relationships ${PR_NS}>` +
                    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
                    `</Relationships>`;
                await writeFile(relsFile, content);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const repairs = await v.repairEmptyRelsParts();
                expect(repairs).toBe(0);
                expect(await fs.readFile(relsFile, "utf-8")).toBe(content);
            });
        });

        it("removes deleted files from xmlFiles so validateNoEmptyRelsParts does not re-read them", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "_rels", "a.xml.rels"), `${RELS_HEADER}\n<Relationships ${PR_NS}/>`);
                await writeFile(path.join(dir, "word", "_rels", "b.xml.rels"), `${RELS_HEADER}\n<Relationships ${PR_NS}/>`);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                expect(v.xmlFiles.filter((f) => f.endsWith(".rels"))).toHaveLength(2);
                const repairs = await v.repairEmptyRelsParts();
                expect(repairs).toBe(2);
                expect(v.xmlFiles.filter((f) => f.endsWith(".rels"))).toHaveLength(0);
            });
        });
    });

    describe("validateRels path traversal mitigation", () => {
        const PR_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
        it("closes the existence oracle: an escaping target that EXISTS on disk is still broken", async () => {
            await withTempDir(async (dir) => {
                // A real file outside the unpacked dir — pre-containment code
                // would count it as "referenced" (the oracle leak).
                await fs.writeFile(path.join(dir, "outside-secret.txt"), "secret", "utf-8");
                const relsFile = path.join(dir, "word", "_rels", "document.xml.rels");
                const content =
                    `${RELS_HEADER}\n<Relationships ${PR_NS}>` +
                    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../../outside-secret.txt"/>` +
                    `</Relationships>`;
                await writeFile(relsFile, content);
                const v = new HarnessValidator({ unpackedDir: dir });
                const result = await v.validateFileReferences();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-broken" && i.message.includes("../../outside-secret.txt"))).toBe(true);
            });
        });

        it("rejects targets routed through a symlink that escapes the unpacked directory", async () => {
            await withTempDir(async (dir) => {
                // outside/secret.txt exists beyond the unpacked root; the
                // symlink inside the root makes "link/secret.txt" lexically
                // contained while really escaping it.
                const unpacked = path.join(dir, "unpacked");
                await fs.mkdir(path.join(dir, "outside"), { recursive: true });
                await fs.mkdir(path.join(unpacked, "_rels"), { recursive: true });
                await fs.writeFile(path.join(dir, "outside", "secret.txt"), "secret", "utf-8");
                await fs.symlink(path.join(dir, "outside"), path.join(unpacked, "link"));
                const relsFile = path.join(unpacked, "_rels", ".rels");
                const content =
                    `${RELS_HEADER}\n<Relationships ${PR_NS}>` +
                    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="link/secret.txt"/>` +
                    `</Relationships>`;
                await fs.writeFile(relsFile, content, "utf-8");
                const v = new HarnessValidator({ unpackedDir: unpacked });
                const result = await v.validateFileReferences();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "rels-broken" && i.message.includes("link/secret.txt"))).toBe(true);
            });
        });

        it("rejects Target paths that traverse outside the unpacked directory", async () => {
            await withTempDir(async (dir) => {
                const relsFile = path.join(dir, "word", "_rels", "document.xml.rels");
                const content =
                    `${RELS_HEADER}\n<Relationships ${PR_NS}>` +
                    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../../../../etc/passwd"/>` +
                    `</Relationships>`;
                await writeFile(relsFile, content);
                const v = new HarnessValidator({ unpackedDir: dir, profile: "strict" });
                const result = await v.validateFileReferences();
                expect(result.valid).toBe(false);
                expect(
                    result.issues.some((i) => i.code === "rels-broken" && i.message.includes("Broken reference to ../../../../etc/passwd")),
                ).toBe(true);
            });
        });
    });
});

describe("withoutRestatedIgnorablePrefixes (#309)", () => {
    const undeclared = (prefix: string, path = "word/document.xml") => ({
        severity: "error" as const,
        code: "ignorable-undeclared",
        path,
        message: `Namespace '${prefix}' in Ignorable but not declared`,
    });
    const unbound = (directive: string, prefix: string, path = "word/document.xml") => ({
        severity: "error" as const,
        code: "mce-unbound-prefix",
        path,
        message: `Unbound prefix in mc:${directive}: ${prefix}`,
    });

    it("drops one processor finding per namespace-check finding on the same part and prefix", () => {
        const result = withoutRestatedIgnorablePrefixes({
            valid: false,
            issues: [undeclared("w15"), unbound("Ignorable", "w15"), unbound("Ignorable", "w15")],
        });
        expect(result.issues).toEqual([undeclared("w15"), unbound("Ignorable", "w15")]);
        expect(result.valid).toBe(false);
    });

    it("keeps findings on another part, another prefix, or another directive", () => {
        const issues = [
            undeclared("w15"),
            unbound("Ignorable", "w15", "word/styles.xml"),
            unbound("Ignorable", "wp14"),
            unbound("ProcessContent", "w15"),
            unbound("MustUnderstand", "w15"),
        ];
        expect(withoutRestatedIgnorablePrefixes({ valid: false, issues }).issues).toEqual(issues);
    });

    it("matches a namespace-check path written with Windows separators (#310 review)", () => {
        // validateNamespaces reports path.relative paths (backslashes on
        // Windows); the Word rule reports the normalized part name.
        const result = withoutRestatedIgnorablePrefixes({
            valid: false,
            issues: [undeclared("w15", "word\\document.xml"), unbound("Ignorable", "w15")],
        });
        expect(result.issues).toEqual([undeclared("w15", "word\\document.xml")]);
    });

    it("returns the result unchanged when the namespace check found nothing", () => {
        const result = { valid: false, issues: [unbound("Ignorable", "w15")] };
        expect(withoutRestatedIgnorablePrefixes(result)).toBe(result);
    });
});

// Filesystem-backed integration regressions: packages are built here, never fixtures.
describe("OPC-equivalent relationships parts (#398)", () => {
    const variants = ["RELS", "Rels", "%72els", "%52ELS"];
    const relationships = (body: string): string =>
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

    it.each(variants)("reports empty notes.%s like notes.rels, without repairing it", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                const part = `word/notes.${ext}`;
                const xml = relationships("");
                await writeFile(path.join(root, part), xml);
                const validator = new HarnessValidator({ unpackedDir: root, profile: "strict" });
                const result = await validator.validateNoEmptyRelsParts();
                expect(await fs.readFile(path.join(root, part), "utf8")).toBe(xml);
                return result;
            };
            const control = await run("rels");
            expect(control).toMatchObject({
                valid: false,
                issues: [{ code: "rels-empty-part", severity: "error", path: "word/notes.rels" }],
            });
            const variant = await run(extension);
            expect(variant).toEqual({ ...control, issues: control.issues.map((issue) => ({ ...issue, path: `word/notes.${extension}` })) });
        });
    });

    it.each(variants)("checks Relationship element content in .%s", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                await writeFile(
                    path.join(root, `word/_rels/document.xml.${ext}`),
                    relationships(`<Relationship Id="rId1" Type="styles" Target="styles.xml"><child/></Relationship>`),
                );
                return new HarnessValidator({ unpackedDir: root, profile: "lenient" }).validateRelationshipElements();
            };
            const control = await run("rels");
            expect(control).toMatchObject({ valid: false, issues: [{ code: "rels-empty-element" }] });
            expect(await run(extension)).toEqual({
                ...control,
                issues: control.issues.map((issue) => ({ ...issue, path: `word/_rels/document.xml.${extension}` })),
            });
        });
    });

    it.each(variants)("schema-validates a bogus attribute in .%s", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                const part = `word/_rels/document.xml.${ext}`;
                await writeFile(
                    path.join(root, part),
                    relationships(`<Relationship Id="rId1" Type="styles" Target="styles.xml" bogus="1"/>`),
                );
                const validator = new HarnessValidator({ unpackedDir: root, profile: "strict" });
                const result = await validator.validateAgainstXsd();
                return { result, validator, part };
            };
            const control = await run("rels");
            expect(control.result.valid).toBe(false);
            expect(control.result.issues).toContainEqual(expect.objectContaining({ code: "xsd-error", path: control.part }));
            const variant = await run(extension);
            expect(variant.result.valid).toBe(control.result.valid);
            expect(variant.result.issues.map(({ code }) => code)).toEqual(control.result.issues.map(({ code }) => code));
            expect(variant.result.issues.every(({ path: part }) => part === variant.part)).toBe(true);
            expect(variant.validator.xsdWalk.parts).toEqual([variant.part]);
        });
    });

    it.each(variants)("maps .%s to the relationships schema key", async (extension) => {
        await withTempDir(async (dir) => {
            const validator = new HarnessValidator({ unpackedDir: dir });
            expect(validator.schemaMappingKey("word/_rels/document.xml.rels")).toBe(".rels");
            expect(validator.schemaMappingKey(`word/_rels/document.xml.${extension}`)).toBe(".rels");
            // Other mapping rules retain their existing spelling semantics.
            expect(validator.schemaMappingKey("word/commentsIds.xml")).toBe("commentsIds.xml");
            expect(validator.schemaMappingKey("customXml/itemProps1.%78ml")).toBeNull();
        });
    });

    it.each(["RELS", "%72els", "%52ELS"])("reads a clean .%s sidecar without unreferenced findings", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                for (const [part, xml] of Object.entries(MINIMAL_PARTS)) {
                    await writeFile(
                        path.join(root, part === "word/_rels/document.xml.rels" ? `word/_rels/document.xml.${ext}` : part),
                        xml,
                    );
                }
                return new HarnessValidator({ unpackedDir: root }).validateFileReferences();
            };
            const control = await run("rels");
            expect(control).toEqual({ valid: true, issues: [] });
            expect(await run(extension)).toEqual(control);
        });
    });

    it("includes .XML parts in the XSD walk and passes iterating xmlFiles", async () => {
        await withTempDir(async (dir) => {
            const run = async (extension: string) => {
                const root = path.join(dir, extension === "xml" ? "control" : "variant");
                const part = `word/comments.${extension}`;
                await writeFile(
                    path.join(root, part),
                    `<w:comments ${W_NS} bogus="1"><w:comment w:id="0"/><w:comment w:id="0"/></w:comments>`,
                );
                await writeFile(path.join(root, "ignore.txt"), "not XML");
                const validator = new HarnessValidator({ unpackedDir: root, profile: "strict" });
                return { part, validator, schema: await validator.validateAgainstXsd(), ids: await validator.validateUniqueIds() };
            };
            const control = await run("xml");
            expect(control.schema.valid).toBe(false);
            expect(control.ids.issues).toContainEqual(expect.objectContaining({ code: "id-duplicate-file" }));
            const variant = await run("XML");
            expect(variant.schema.valid).toBe(control.schema.valid);
            expect(variant.schema.issues.map(({ code }) => code)).toEqual(control.schema.issues.map(({ code }) => code));
            expect(variant.schema.issues.every(({ path: part }) => part === variant.part)).toBe(true);
            expect(variant.ids).toEqual({ ...control.ids, issues: control.ids.issues.map((issue) => ({ ...issue, path: variant.part })) });
            expect(variant.validator.xsdWalk.parts).toEqual([variant.part]);
            expect(variant.validator.xmlFiles).toEqual([path.join(dir, "variant", variant.part)]);
        });
    });

    it.each(["RELS", "%72els"])("does not treat a .%s part as an XML source needing a sidecar", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                await writeFile(
                    path.join(root, `word/notes.${ext}`),
                    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="unused"/>`,
                );
                const validator = new HarnessValidator({ unpackedDir: root });
                // Exercise the source/sidecar discriminator independently of discovery.
                if (validator.xmlFiles.length === 0) validator.xmlFiles.push(path.join(root, `word/notes.${ext}`));
                return validator.validateAllRelationshipIds();
            };
            const control = await run("rels");
            expect(control).toEqual({ valid: true, issues: [] });
            expect(await run(extension)).toEqual(control);
        });
    });
});

describe("OPC-equivalent content-type relationship skip (#398)", () => {
    it.each(["RELS", "%72els"])("skips .%s as .rels in the declarable-root pass", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const root = path.join(dir, ext === "rels" ? "control" : "variant");
                const part = `word/notes.${ext}`;
                await writeFile(path.join(root, "[Content_Types].xml"), MINIMAL_PARTS["[Content_Types].xml"]);
                await writeFile(path.join(root, part), `<w:document ${W_NS}><w:body/></w:document>`);
                const validator = new HarnessValidator({ unpackedDir: root });
                // Keep discovery independent: this pass must skip relationships
                // parts even if their root resembles a declarable XML part.
                if (!validator.xmlFiles.includes(path.join(root, part))) validator.xmlFiles.push(path.join(root, part));
                return validator.validateContentTypes();
            };
            const control = await run("rels");
            expect(control).toEqual({ valid: true, issues: [] });
            expect(await run(extension)).toEqual(control);
        });
    });
});

describe("content-types stream directory lookup (#396)", () => {
    it("reads the stored uppercase stream and excludes it from part coverage", async () => {
        await withTempDir(async (dir) => {
            const name = "[CONTENT_TYPES].XML";
            const xml = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`;
            await writeFile(path.join(dir, name), xml);
            const result = await new HarnessValidator({ unpackedDir: dir }).validateContentTypes();
            expect(result.issues).toEqual([]);
            expect(result.valid).toBe(true);
            expect(await fs.readdir(dir)).toEqual([name]);
            expect(await fs.readFile(path.join(dir, name), "utf8")).toBe(xml);
        });
    });

    it("reports ct-parse with the stored uppercase stream name", async () => {
        await withTempDir(async (dir) => {
            const name = "[CONTENT_TYPES].XML";
            await writeFile(path.join(dir, name), "<broken");
            const result = await new HarnessValidator({ unpackedDir: dir }).validateContentTypes();
            expect(result.issues).toEqual([expect.objectContaining({ code: "ct-parse", path: name })]);
            expect(result.issues[0].message).toContain(name);
            expect(await fs.readFile(path.join(dir, name), "utf8")).toBe("<broken");
        });
    });

    it("does not skip the percent-escaped part in coverage", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "[CONTENT_TYPES].XML"),
                `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
            );
            await writeFile(path.join(dir, "%5BContent_Types%5D.xml"), "<x/>");
            const result = await new HarnessValidator({ unpackedDir: dir }).validateContentTypes();
            expect(result.issues).toEqual([expect.objectContaining({ code: "ct-uncovered-part", path: "%5BContent_Types%5D.xml" })]);
        });
    });

    it("reports a folder named [Content_Types].xml as a missing stream, not a malformed one", async () => {
        await withTempDir(async (dir) => {
            // A folder holds no stream: the entries list files only, where
            // existsSync matched the folder and readFile then failed (review S).
            await fs.mkdir(path.join(dir, "[Content_Types].xml"));
            const result = await new HarnessValidator({ unpackedDir: dir }).validateContentTypes();
            expect(result.valid).toBe(false);
            expect(result.issues.map((issue) => issue.code)).toEqual(["ct-missing"]);
        });
    });

    it("does not classify the uppercase stream as an unreferenced part", async () => {
        await withTempDir(async (dir) => {
            await writeFile(path.join(dir, "[CONTENT_TYPES].XML"), MINIMAL_PARTS["[Content_Types].xml"]);
            await writeFile(path.join(dir, "_rels/.rels"), MINIMAL_PARTS["_rels/.rels"]);
            await writeFile(path.join(dir, "word/document.xml"), MINIMAL_PARTS["word/document.xml"]);
            const result = await new HarnessValidator({ unpackedDir: dir }).validateFileReferences();
            expect(result.issues.filter((issue) => issue.code === "rels-unreferenced")).toEqual([]);
        });
    });
});
