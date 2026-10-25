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
import type { Profile } from "../src/lib/types";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { MINIMAL_PARTS } from "./support/minimal-docx";

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

describe("schema routing of glossary and diagram parts (#628)", () => {
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
