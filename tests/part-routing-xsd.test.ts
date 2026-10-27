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

// #628: the schema walk routes the glossary document's parts to wml.xsd and
// the DrawingML diagram parts to dml-diagram.xsd. Both folders sit below a
// main content folder, so the folder rule never reached them and their parts
// were left out of the walk.

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { RunDeadline } from "../src/lib/run-deadline";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DGM = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';

class KeyValidator extends BaseSchemaValidator {
    schemaMappingKey(part: string): string | null {
        return this._schemaMappingKey(path.join(this.unpackedDir, part));
    }

    schemaFile(part: string): string | null {
        const schema = this._getSchemaPath(path.join(this.unpackedDir, part));
        return schema === null ? null : path.relative(this.schemasDir, schema).split(path.sep).join("/");
    }
}

/**
 * The walk reported `element` in `part` and nothing else: lenient adds a
 * per-part `N new error(s)` line before the engine messages.
 */
const reportsOnly = (errors: readonly string[], part: string, element: string): void => {
    expect(errors.length).toBeGreaterThan(0);
    expect(errors.every((error) => error.startsWith(`${part}: `))).toBe(true);
    expect(errors.some((error) => error.includes(`}${element}'`))).toBe(true);
};

/** Walk the schemas of a minimal package plus `extra` parts under `profile`. */
const schemaWalk = async (profile: Profile, extra: Readonly<Record<string, string>>) =>
    withTempDir(async (dir) => {
        for (const [part, xml] of Object.entries({ ...MINIMAL_PARTS, ...extra })) {
            await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
            await fs.writeFile(path.join(dir, part), xml, "utf-8");
        }
        const validator = new DOCXSchemaValidator({ unpackedDir: dir, profile });
        const result = await validator.validateAgainstXsd();
        return {
            errors: result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.path}: ${issue.message}`),
            walked: validator.xsdWalk.parts,
        };
    });

describe("schema routing of glossary and diagram parts (#628)", { timeout: 120_000 }, () => {
    it.each([
        ["word/glossary/document.xml", "word"],
        ["word/glossary/styles.xml", "word"],
        ["word/glossary/settings.xml", "word"],
        ["WORD/Glossary/fontTable.xml", "word"],
        ["word/diagrams/data1.xml", "diagram"],
        ["word/diagrams/layout2.xml", "diagram"],
        ["word/diagrams/quickStyle1.xml", "diagram"],
        ["word/diagrams/colors10.xml", "diagram"],
        ["ppt/diagrams/data1.xml", "diagram"],
        ["xl/diagrams/layout1.xml", "diagram"],
    ])("maps %s to the %s schema key", (part, key) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBe(key);
    });

    it.each([
        // Microsoft's diagram drawing cache (dsp:drawing) has no bundled schema.
        "word/diagrams/drawing1.xml",
        "word/diagrams/dataModel.xml",
        "word/diagrams/data1.xml/extra.xml",
        "word/glossary/nested/document.xml",
        "glossary/document.xml",
        "customXml/item1.xml",
    ])("leaves %s unmapped", (part) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBeNull();
    });

    it("routes a Transitional diagram part to the bundled diagram schema", () => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaFile("word/diagrams/data1.xml")).toBe("ISO-IEC29500-4_2016/dml-diagram.xsd");
    });

    it("routes a Strict diagram part to the ISO Strict diagram schema under strict", async () => {
        const { errors, walked } = await schemaWalk("strict", {
            "word/diagrams/data1.xml": `${HEAD}<dgm:dataModel xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"><dgm:ptLst/><dgm:bogusStrict/></dgm:dataModel>`,
        });
        expect(walked).toContain("word/diagrams/data1.xml");
        reportsOnly(errors, "word/diagrams/data1.xml", "bogusStrict");
    });

    describe.each(["strict", "lenient"] as const)("under %s", (profile) => {
        it("walks a valid glossary document and diagram data part without findings", async () => {
            const { errors, walked } = await schemaWalk(profile, {
                "word/glossary/document.xml": `${HEAD}<w:glossaryDocument ${W}/>`,
                "word/diagrams/data1.xml": `${HEAD}<dgm:dataModel ${DGM}><dgm:ptLst/></dgm:dataModel>`,
            });
            expect(errors).toEqual([]);
            expect(walked).toEqual(expect.arrayContaining(["word/glossary/document.xml", "word/diagrams/data1.xml"]));
        });

        it("reports a schema violation in the glossary document", async () => {
            const { errors } = await schemaWalk(profile, {
                "word/glossary/document.xml": `${HEAD}<w:glossaryDocument ${W}><w:bogusGlossary/></w:glossaryDocument>`,
            });
            reportsOnly(errors, "word/glossary/document.xml", "bogusGlossary");
        });

        it("reports a schema violation in a glossary styles part", async () => {
            const { errors } = await schemaWalk(profile, {
                "word/glossary/styles.xml": `${HEAD}<w:styles ${W}><w:bogusStyle/></w:styles>`,
            });
            reportsOnly(errors, "word/glossary/styles.xml", "bogusStyle");
        });

        it.each([
            ["data1.xml", `<dgm:dataModel ${DGM}><dgm:ptLst/><dgm:bogusDiagram/></dgm:dataModel>`],
            ["quickStyle1.xml", `<dgm:styleDef ${DGM}><dgm:bogusDiagram/></dgm:styleDef>`],
        ])("reports a schema violation in the diagram part %s", async (base, root) => {
            const { errors } = await schemaWalk(profile, { [`word/diagrams/${base}`]: `${HEAD}${root}` });
            reportsOnly(errors, `word/diagrams/${base}`, "bogusDiagram");
        });
    });
});

describe("schema audit of unmapped parts (#628)", { timeout: 120_000 }, () => {
    const unmapped = {
        "customXml/item1.xml": "<business>payload</business>",
        "word/charts/style1.xml": '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>',
    };
    const parts = {
        ...MINIMAL_PARTS,
        ...unmapped,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
                '<Relationship Id="rStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="charts/style1.xml"/></Relationships>',
        ),
    };
    const audit = (part: string) => ({
        severity: "info",
        message: "The bundled schemas map no schema to this part; XSD validation skipped for this part.",
        path: part,
        code: "xsd-no-schema",
    });
    const mapped = ["[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels", "word/document.xml", "word/styles.xml"];

    describe.each([false, true])("direct schema walk after validateXml: %s (#867)", (validateXmlFirst) => {
        it.each([
            ["malformed", "<business>", []],
            ["well-formed", unmapped["customXml/item1.xml"], [audit("customXml/item1.xml")]],
        ] as const)("audits only a part that parses: %s", async (shape, payload, expected) => {
            await withTempDir(async (dir) => {
                const directParts = {
                    ...MINIMAL_PARTS,
                    "customXml/item1.xml": payload,
                    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                        "</Relationships>",
                        '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
                    ),
                };
                for (const [part, xml] of Object.entries(directParts)) {
                    await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                    await fs.writeFile(path.join(dir, part), xml, "utf-8");
                }
                const validator = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
                if (validateXmlFirst) {
                    const xml = await validator.validateXml();
                    expect(xml.valid).toBe(shape === "well-formed");
                }
                const result = await validator.validateAgainstXsd();
                expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual(expected);
                expect(result.valid).toBe(true);
                expect(validator.xsdWalk).toEqual({ started: true, parts: mapped, eligible: mapped.length });
            });
        });
    });

    describe.each(["strict", "lenient"] as const)("under %s", (profile) => {
        it.each(["docx", "docm"])(
            "reports each unmapped %s part once in walk order without changing coverage or validity",
            async (suffix) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `unmapped.${suffix}`);
                    await fs.writeFile(file, await buildDocx(parts));
                    const result = await validate(file, { profile });
                    expect(result.valid).toBe(true);
                    expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([
                        audit("customXml/item1.xml"),
                        audit("word/charts/style1.xml"),
                    ]);
                    expect(result.xsdWalk).toEqual({ started: true, parts: mapped, eligible: mapped.length });
                    expect(result.metadata?.schemaCompleted).toBe(true);
                });
            },
        );

        it("reports no audit finding when every part maps to a schema", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "mapped.docx");
                await fs.writeFile(file, await buildDocx());
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([]);
                expect(result.xsdWalk).toEqual({ started: true, parts: mapped, eligible: mapped.length });
            });
        });

        it("reports malformed XML without an unmapped-schema audit finding", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "malformed.docx");
                await fs.writeFile(file, await buildDocx({ ...parts, "customXml/item1.xml": "<business>" }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.path === "customXml/item1.xml").map((issue) => issue.code)).toEqual([
                    "xml-syntax",
                ]);
            });
        });
    });

    it.each(["docx", "docm"])("does not audit unmapped %s parts under word-valid", async (suffix) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `unmapped.${suffix}`);
            await fs.writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([]);
            expect(result.xsdWalk).toEqual({ started: false, parts: [], eligible: 0 });
            expect(result.metadata?.schemaCompleted).toBe(false);
        });
    });

    it.each(["word-valid", "lenient"] as const)("does not audit unmapped parts in a partial %s run", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unmapped.docx");
            await fs.writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile, checks: ["xml-syntax"] });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([]);
            expect(result.xsdWalk).toEqual({ started: false, parts: [], eligible: 0 });
            expect(result.metadata?.schemaCompleted).toBe(false);
        });
    });

    it.each(["lenient", "strict"] as const)("reports the appropriate skip for a Strict unmapped part under %s", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "strict-unmapped.docx");
            await fs.writeFile(
                file,
                await buildDocx({
                    ...parts,
                    "customXml/item1.xml": '<a:payload xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"/>',
                }),
            );
            const result = await validate(file, { profile });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code?.startsWith("xsd-"))).toEqual(
                profile === "strict"
                    ? [audit("customXml/item1.xml"), audit("word/charts/style1.xml")]
                    : [
                          {
                              severity: "info",
                              message:
                                  "Document uses ISO OOXML Strict conformance class namespace URIs; " +
                                  "XSD validation skipped under this profile. The bundled ISO Strict schemas " +
                                  "(schemas/iso-strict/) validate these parts under the strict profile.",
                              code: "xsd-strict-skipped",
                          },
                          audit("word/charts/style1.xml"),
                      ],
            );
        });
    });

    // The minimal PPTX's slide, layout and master are unmapped too: no rule maps a
    // PresentationML part below `ppt/` (#864), so they are audited beside the payload.
    it.each([undefined, "lenient", "strict", "word-valid"] as const)("audits an unmapped PPTX part under profile %s", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unmapped.pptx");
            await fs.writeFile(
                file,
                await buildPptx({
                    ...MINIMAL_PPTX_PARTS,
                    "customXml/item1.xml": unmapped["customXml/item1.xml"],
                    "ppt/_rels/presentation.xml.rels": MINIMAL_PPTX_PARTS["ppt/_rels/presentation.xml.rels"]!.replace(
                        "</Relationships>",
                        '<Relationship Id="rCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
                    ),
                }),
            );
            const result = await validate(file, { profile });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([
                audit("customXml/item1.xml"),
                audit("ppt/slideLayouts/slideLayout1.xml"),
                audit("ppt/slideMasters/slideMaster1.xml"),
                audit("ppt/slides/slide1.xml"),
            ]);
            expect(result.xsdWalk.parts).not.toContain("customXml/item1.xml");
            expect(result.metadata?.schemaCompleted).toBe(true);
        });
    });

    it("keeps the lenient audit when --original has the same unmapped parts", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unmapped.docx");
            const original = path.join(dir, "original.docx");
            const bytes = await buildDocx(parts);
            await fs.writeFile(file, bytes);
            await fs.writeFile(original, bytes);
            const result = await validate(file, { profile: "lenient", original, author: "Test Author" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([
                audit("customXml/item1.xml"),
                audit("word/charts/style1.xml"),
            ]);
        });
    });

    it("does not audit parts the schema walk's deadline never reached", async () => {
        await withTempDir(async (dir) => {
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            let now = 0;
            const deadline = new RunDeadline(1, () => now);
            const validator = new DOCXSchemaValidator({
                unpackedDir: dir,
                profile: "strict",
                deadline,
                xsdValidator: {
                    validate: async () => {
                        now = 1;
                        return [];
                    },
                },
            });
            const result = await validator.validateAgainstXsd();
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([]);
            expect(validator.xsdWalk).toEqual({ started: true, parts: ["[Content_Types].xml"], eligible: 1, notExamined: 6 });
            expect(deadline.notRun).toEqual(["schema validation of 6 XML parts it had not reached"]);
        });
    });
});
