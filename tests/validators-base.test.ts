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

import { describe, expect, it } from "vitest";

import { BaseSchemaValidator } from "../src/scripts/office/validators/base.ts";
import { withTempDir } from "../src/lib/run-cli.ts";

class HarnessValidator extends BaseSchemaValidator {
    // Subclass concrete-only so we can instantiate.
}

const RELS_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

async function writeFile(p: string, content: string): Promise<void> {
    await fs.mkdir(path.dirname(p), { recursive: true });
    await fs.writeFile(p, content, "utf-8");
}

describe("BaseSchemaValidator", () => {
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
            // IGNORED_VALIDATION_ERRORS list filter out known-noisy schemas like
            // docProps/core.xml whose `dcterms` import libxmljs2 cannot fully
            // resolve. For loud failure on a broken libxmljs2 binding, callers
            // should invoke BaseSchemaValidator.assertLibxmljsAvailable() at
            // startup.
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
    });

    describe("assertLibxmljsAvailable", () => {
        it("returns silently when libxmljs2 + XSD validation work on this host", () => {
            expect(() => BaseSchemaValidator.assertLibxmljsAvailable()).not.toThrow();
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
});
