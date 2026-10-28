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
//
// #864: the same walk routes the PresentationML parts one folder below ppt/
// to pml.xsd — the slides, slide layouts, slide masters, notes slides, notes
// and handout masters, legacy comments and tags whose root element pml.xsd
// declares — and keeps every other part there (PowerPoint's modern comments)
// unmapped.
//
// #872: directly in ppt/, the walk reaches pml.xsd by file name too — only
// the parts whose root pml.xsd, or a schema it imports, declares — so the
// Microsoft extension parts stored beside them (the modern comment authors
// list and the revision info) stay unmapped instead of failing the walk.
//
// #887: those five names are a fast path, nothing more: a part's identity in
// OPC comes from its relationship and content type, so every other name
// directly in ppt/ is routed by its root element — the walk reaches pml.xsd
// when the root is a global element of its schema closure for the part's
// conformance class — and a renamed main presentation or a relocated slide
// keeps its schema check, while the Microsoft extension roots stay unmapped.

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { RunDeadline } from "../src/lib/run-deadline";
import type { Profile } from "../src/lib/types";
import * as xmlHelpers from "../src/lib/xml-helpers";
import { schemaGlobals } from "../src/scripts/office/schema-globals";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator, defaultSchemasDir } from "../src/scripts/office/validators/base";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { PPTXSchemaValidator } from "../src/scripts/office/validators/pptx";
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

describe("schema routing of PresentationML parts below ppt/ (#864)", { timeout: 120_000 }, () => {
    const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PML_CT = "application/vnd.openxmlformats-officedocument.presentationml";
    const PPT_NS =
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:r="${OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
    const STRICT_PPT_NS =
        `xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" ` +
        `xmlns:r="http://purl.oclc.org/ooxml/officeDocument" ` +
        `xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"`;
    // The same shape tree and colour map the minimal PPTX's slide parts hold.
    const PPT_TREE =
        `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr/></p:spTree></p:cSld>`;
    const PPT_CLR_MAP =
        `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ` +
        `accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`;

    // The five part kinds the minimal package lacks, each the smallest part
    // pml.xsd accepts: `p:notes` and the legacy `p:cmLst` relate from the
    // slide, the masters and the tags from the presentation, as the Open XML
    // SDK part catalogue relates them (data/parts/NotesSlidePart.json,
    // SlideCommentsPart.json, NotesMasterPart.json, HandoutMasterPart.json,
    // UserDefinedTagsPart.json).
    const NOTES_SLIDE = `${HEAD}<p:notes ${PPT_NS}>${PPT_TREE}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
    const NOTES_MASTER = `${HEAD}<p:notesMaster ${PPT_NS}>${PPT_TREE}${PPT_CLR_MAP}</p:notesMaster>`;
    const HANDOUT_MASTER = `${HEAD}<p:handoutMaster ${PPT_NS}>${PPT_TREE}${PPT_CLR_MAP}</p:handoutMaster>`;
    const LEGACY_COMMENTS = `${HEAD}<p:cmLst ${PPT_NS}></p:cmLst>`;
    const TAGS = `${HEAD}<p:tagLst ${PPT_NS}></p:tagLst>`;

    /** The eight part kinds one folder below `ppt/` whose root pml.xsd declares. */
    const PPTX_PARTS = [
        "ppt/slides/slide1.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/notesSlides/notesSlide1.xml",
        "ppt/notesMasters/notesMaster1.xml",
        "ppt/handoutMasters/handoutMaster1.xml",
        "ppt/comments/comment1.xml",
        "ppt/tags/tag1.xml",
    ] as const;

    /** The minimal PPTX holding all eight kinds, each with its relationship and content-type override. */
    const pptxHoldingAllParts = (): Record<string, string> => ({
        ...MINIMAL_PPTX_PARTS,
        "ppt/notesSlides/notesSlide1.xml": NOTES_SLIDE,
        "ppt/notesMasters/notesMaster1.xml": NOTES_MASTER,
        "ppt/handoutMasters/handoutMaster1.xml": HANDOUT_MASTER,
        "ppt/comments/comment1.xml": LEGACY_COMMENTS,
        "ppt/tags/tag1.xml": TAGS,
        "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="${PML_CT}.notesSlide+xml"/>` +
                `<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${PML_CT}.notesMaster+xml"/>` +
                `<Override PartName="/ppt/handoutMasters/handoutMaster1.xml" ContentType="${PML_CT}.handoutMaster+xml"/>` +
                `<Override PartName="/ppt/comments/comment1.xml" ContentType="${PML_CT}.comments+xml"/>` +
                `<Override PartName="/ppt/tags/tag1.xml" ContentType="${PML_CT}.tags+xml"/></Types>`,
        ),
        "ppt/_rels/presentation.xml.rels": MINIMAL_PPTX_PARTS["ppt/_rels/presentation.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdNotesMaster" Type="${OFFICE_REL}/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
                `<Relationship Id="rIdHandoutMaster" Type="${OFFICE_REL}/handoutMaster" Target="handoutMasters/handoutMaster1.xml"/>` +
                `<Relationship Id="rIdTags" Type="${OFFICE_REL}/tags" Target="tags/tag1.xml"/></Relationships>`,
        ),
        "ppt/slides/_rels/slide1.xml.rels": MINIMAL_PPTX_PARTS["ppt/slides/_rels/slide1.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdNotesSlide" Type="${OFFICE_REL}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
                `<Relationship Id="rIdComments" Type="${OFFICE_REL}/comments" Target="../comments/comment1.xml"/></Relationships>`,
        ),
    });

    /** The part's XML with `<p:bogusDefect/>` as the root's first child. */
    const withBogusDefect = (xml: string): string => {
        const at = xml.indexOf(">", xml.indexOf("<p:")) + 1;
        return xml.slice(0, at) + "<p:bogusDefect/>" + xml.slice(at);
    };

    const noSchemaAudit = (part: string) => ({
        severity: "info",
        message: "The bundled schemas map no schema to this part; XSD validation skipped for this part.",
        path: part,
        code: "xsd-no-schema",
    });

    it.each([...PPTX_PARTS.map((part) => [part, "ppt"] as const), ["PPT/Slides/Slide2.XML", "ppt"] as const])(
        "maps %s to the ppt schema key",
        (part, key) => {
            expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBe(key);
        },
    );

    it.each([
        // PowerPoint's modern comments sit in the comments folder, but their
        // root belongs to the Microsoft 2018/8 namespace pml.xsd never declares.
        "ppt/comments/modernComment_1_0.xml",
        // The folder and the file name must agree …
        "ppt/slides/slideMaster1.xml",
        // … and the part must sit exactly one folder below ppt/.
        "ppt/slides/slide1.xml/extra.xml",
        "ppt/notesSlides/notesSlide1.xml/extra.xml",
        "word/slides/slide1.xml",
        // A folder named like an Object.prototype key names no part kind.
        "ppt/constructor/slide1.xml",
        "ppt/__proto__/slide1.xml",
    ])("leaves %s unmapped", (part) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBeNull();
    });

    it("routes a slide to the bundled PresentationML schema", () => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaFile("ppt/slides/slide1.xml")).toBe("ISO-IEC29500-4_2016/pml.xsd");
    });

    describe.each(["lenient", "strict"] as const)("under %s", (profile) => {
        it.each(PPTX_PARTS)("reports a schema violation in %s", async (part) => {
            await withTempDir(async (dir) => {
                const parts = pptxHoldingAllParts();
                const file = path.join(dir, "defect.pptx");
                await fs.writeFile(file, await buildPptx({ ...parts, [part]: withBogusDefect(parts[part]!) }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                const errors = result.issues.filter((issue) => issue.code === "xsd-error");
                expect(errors.length).toBeGreaterThan(0);
                expect(errors.every((issue) => issue.path === part)).toBe(true);
                expect(errors.some((issue) => issue.message.includes("bogusDefect"))).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
            });
        });

        it.each(PPTX_PARTS)("walks a valid %s without findings", async (part) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.pptx");
                await fs.writeFile(file, await buildPptx(pptxHoldingAllParts()));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
                expect(result.issues.filter((issue) => issue.path === part)).toEqual([]);
            });
        });
    });

    it("keeps PowerPoint's modern comments unmapped and audited", async () => {
        await withTempDir(async (dir) => {
            const wired = pptxHoldingAllParts();
            const parts = {
                ...wired,
                // The modern comments' root and relationship:
                // PowerPointCommentPart (Open XML SDK data/parts/
                // PowerPointCommentPart.json), whose p188 prefix is the
                // Microsoft 2018/8 namespace.
                "ppt/comments/modernComment_1_0.xml": `${HEAD}<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"></p188:cmLst>`,
                "[Content_Types].xml": wired["[Content_Types].xml"]!.replace(
                    "</Types>",
                    '<Override PartName="/ppt/comments/modernComment_1_0.xml" ContentType="application/vnd.ms-powerpoint.comments+xml"/></Types>',
                ),
                "ppt/slides/_rels/slide1.xml.rels": wired["ppt/slides/_rels/slide1.xml.rels"]!.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdModern" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" ' +
                        'Target="../comments/modernComment_1_0.xml"/></Relationships>',
                ),
            };
            const file = path.join(dir, "modern-comment.pptx");
            await fs.writeFile(file, await buildPptx(parts));
            for (const profile of ["lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).not.toContain("ppt/comments/modernComment_1_0.xml");
                expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([
                    noSchemaAudit("ppt/comments/modernComment_1_0.xml"),
                ]);
            }
        });
    });

    it("routes a Strict slide to the ISO Strict schema under strict, and skips it under lenient", async () => {
        await withTempDir(async (dir) => {
            const strictSlide = `${HEAD}<p:sld ${STRICT_PPT_NS}>${PPT_TREE}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
            const parts = { ...pptxHoldingAllParts(), "ppt/slides/slide1.xml": withBogusDefect(strictSlide) };
            const file = path.join(dir, "strict-slide.pptx");
            await fs.writeFile(file, await buildPptx(parts));

            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(
                strict.issues.some(
                    (issue) =>
                        issue.code === "xsd-error" && issue.path === "ppt/slides/slide1.xml" && issue.message.includes("bogusDefect"),
                ),
            ).toBe(true);
            expect(strict.xsdWalk.parts).toContain("ppt/slides/slide1.xml");

            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.valid).toBe(true);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-strict-skipped")).toHaveLength(1);
            expect(lenient.xsdWalk.parts).not.toContain("ppt/slides/slide1.xml");
        });
    });
});

describe("schema routing of the parts directly in ppt/ (#872)", { timeout: 120_000 }, () => {
    const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PML_CT = "application/vnd.openxmlformats-officedocument.presentationml";
    const PPT_NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
    const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const STRICT_PPT_NS =
        `xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" ` +
        `xmlns:r="http://purl.oclc.org/ooxml/officeDocument" ` +
        `xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"`;

    // The four part kinds the minimal PPTX lacks directly in `ppt/`, each the
    // smallest part pml.xsd accepts: `p:presentationPr`, `p:viewPr`,
    // `p:cmAuthorLst` and the DrawingML `a:tblStyleLst` pml.xsd reaches
    // through its dml-main.xsd import. `p:presentation` the minimal package
    // already holds.
    const PRES_PROPS = `${HEAD}<p:presentationPr ${PPT_NS}></p:presentationPr>`;
    const VIEW_PROPS = `${HEAD}<p:viewPr ${PPT_NS}></p:viewPr>`;
    const COMMENT_AUTHORS = `${HEAD}<p:cmAuthorLst ${PPT_NS}></p:cmAuthorLst>`;
    const TABLE_STYLES = `${HEAD}<a:tblStyleLst ${A_NS} def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"></a:tblStyleLst>`;

    // The Microsoft extension parts PowerPoint stores beside them, whose
    // roots belong to Microsoft namespaces no bundled schema declares: the
    // modern comment authors list and the revision info (#872).
    const AUTHORS =
        `${HEAD}<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">` +
        `<p188:author id="{00000000-0000-0000-0000-000000000001}" name="A" initials="A" userId="A" providerId="None"/></p188:authorLst>`;
    const REVISION_INFO =
        `${HEAD}<p1510:revInfo xmlns:p1510="http://schemas.microsoft.com/office/powerpoint/2015/10/main">` +
        `<p1510:revLst><p1510:client id="{00000000-0000-0000-0000-000000000001}" v="1" dt="2026-10-27T00:00:00.000"/></p1510:revLst></p1510:revInfo>`;

    /** The five part kinds directly in `ppt/` whose root pml.xsd, or a schema it imports, declares. */
    const DIRECT_PARTS = [
        "ppt/presentation.xml",
        "ppt/presProps.xml",
        "ppt/viewProps.xml",
        "ppt/commentAuthors.xml",
        "ppt/tableStyles.xml",
    ] as const;

    /** The minimal PPTX holding all five kinds, each with its relationship and content-type Override. */
    const pptxHoldingDirectParts = (): Record<string, string> => ({
        ...MINIMAL_PPTX_PARTS,
        "ppt/presProps.xml": PRES_PROPS,
        "ppt/viewProps.xml": VIEW_PROPS,
        "ppt/commentAuthors.xml": COMMENT_AUTHORS,
        "ppt/tableStyles.xml": TABLE_STYLES,
        "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/ppt/presProps.xml" ContentType="${PML_CT}.presProps+xml"/>` +
                `<Override PartName="/ppt/viewProps.xml" ContentType="${PML_CT}.viewProps+xml"/>` +
                `<Override PartName="/ppt/commentAuthors.xml" ContentType="${PML_CT}.commentAuthors+xml"/>` +
                `<Override PartName="/ppt/tableStyles.xml" ContentType="${PML_CT}.tableStyles+xml"/></Types>`,
        ),
        "ppt/_rels/presentation.xml.rels": MINIMAL_PPTX_PARTS["ppt/_rels/presentation.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdPresProps" Type="${OFFICE_REL}/presProps" Target="presProps.xml"/>` +
                `<Relationship Id="rIdViewProps" Type="${OFFICE_REL}/viewProps" Target="viewProps.xml"/>` +
                `<Relationship Id="rIdCommentAuthors" Type="${OFFICE_REL}/commentAuthors" Target="commentAuthors.xml"/>` +
                `<Relationship Id="rIdTableStyles" Type="${OFFICE_REL}/tableStyles" Target="tableStyles.xml"/></Relationships>`,
        ),
    });

    /** The minimal PPTX plus `part`, related from the presentation and typed by its Override. */
    const pptxHolding = (part: string, xml: string, relType: string, contentType: string): Record<string, string> => ({
        ...MINIMAL_PPTX_PARTS,
        [part]: xml,
        "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`,
        ),
        "ppt/_rels/presentation.xml.rels": MINIMAL_PPTX_PARTS["ppt/_rels/presentation.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdExtra" Type="${relType}" Target="${part.slice("ppt/".length)}"/></Relationships>`,
        ),
    });

    /** The part's XML with a `<rootPrefix:bogusDefect/>` as the root's first child. */
    const withBogusDefect = (xml: string): string => {
        const root = xml.indexOf("<", xml.indexOf("?>") + 1) + 1;
        const prefix = xml.slice(root, xml.indexOf(">", root)).split(":")[0]!;
        const at = xml.indexOf(">", root) + 1;
        return `${xml.slice(0, at)}<${prefix}:bogusDefect/>${xml.slice(at)}`;
    };

    const noSchemaAudit = (part: string) => ({
        severity: "info",
        message: "The bundled schemas map no schema to this part; XSD validation skipped for this part.",
        path: part,
        code: "xsd-no-schema",
    });

    it.each([
        ...DIRECT_PARTS.map((part) => [part, "ppt"] as const),
        ...["Presentation", "PresProps", "ViewProps", "CommentAuthors", "TableStyles"].map((base) => [`PPT/${base}.XML`, "ppt"] as const),
    ])("maps %s to the ppt schema key", (part, key) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBe(key);
    });

    it.each([
        // The Microsoft extension parts whose roots pml.xsd cannot judge.
        "ppt/authors.xml",
        "ppt/revisionInfo.xml",
        // A file name from the package names no entry, whatever it spells.
        "ppt/constructor.xml",
        "ppt/__proto__.xml",
        "ppt/other.xml",
    ])("leaves %s unmapped", (part) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBeNull();
    });

    it.each([
        ["word/document.xml", "word"],
        ["word/anything.xml", "word"],
        ["xl/workbook.xml", "xl"],
        ["xl/anything.xml", "xl"],
    ])("keeps the %s folder rule (%s)", (part, key) => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey(part)).toBe(key);
    });

    describe.each(["lenient", "strict"] as const)("under %s", (profile) => {
        it.each(DIRECT_PARTS)("reports a schema violation in %s", async (part) => {
            await withTempDir(async (dir) => {
                const parts = pptxHoldingDirectParts();
                const file = path.join(dir, "defect.pptx");
                await fs.writeFile(file, await buildPptx({ ...parts, [part]: withBogusDefect(parts[part]!) }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                const errors = result.issues.filter((issue) => issue.code === "xsd-error");
                expect(errors.length).toBeGreaterThan(0);
                expect(errors.every((issue) => issue.path === part)).toBe(true);
                expect(errors.some((issue) => issue.message.includes("bogusDefect"))).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
            });
        });

        it.each(DIRECT_PARTS)("walks a valid %s without findings", async (part) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.pptx");
                await fs.writeFile(file, await buildPptx(pptxHoldingDirectParts()));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
                expect(result.issues.filter((issue) => issue.path === part)).toEqual([]);
            });
        });

        it.each([
            // The relationship and content types the issue reproduced them by
            // (the Open XML SDK part catalogue's PowerPointAuthorsPart and
            // revision info part).
            [
                "ppt/authors.xml",
                AUTHORS,
                "http://schemas.microsoft.com/office/2018/10/relationships/authors",
                "application/vnd.ms-powerpoint.authors+xml",
            ],
            [
                "ppt/revisionInfo.xml",
                REVISION_INFO,
                "http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo",
                "application/vnd.ms-powerpoint.revisioninfo+xml",
            ],
        ] as const)("keeps %s unmapped and audited", async (part, xml, relType, contentType) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "extension.pptx");
                await fs.writeFile(file, await buildPptx(pptxHolding(part, xml, relType, contentType)));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === "xsd-error" || issue.code === "xsd-summary")).toEqual([]);
                expect(result.xsdWalk.parts).not.toContain(part);
                expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([noSchemaAudit(part)]);
                expect(result.metadata?.schemaCompleted).toBe(true);
            });
        });
    });

    it("routes a Strict presentation to the ISO Strict schema under strict, and skips it under lenient", async () => {
        await withTempDir(async (dir) => {
            const strictPresentation =
                `${HEAD}<p:presentation ${STRICT_PPT_NS}>` +
                `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
                `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
                `<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>` +
                `</p:presentation>`;
            const parts = { ...pptxHoldingDirectParts(), "ppt/presentation.xml": withBogusDefect(strictPresentation) };
            const file = path.join(dir, "strict-presentation.pptx");
            await fs.writeFile(file, await buildPptx(parts));

            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(
                strict.issues.some(
                    (issue) => issue.code === "xsd-error" && issue.path === "ppt/presentation.xml" && issue.message.includes("bogusDefect"),
                ),
            ).toBe(true);
            expect(strict.xsdWalk.parts).toContain("ppt/presentation.xml");

            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.valid).toBe(true);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-strict-skipped")).toHaveLength(1);
            expect(lenient.xsdWalk.parts).not.toContain("ppt/presentation.xml");
        });
    });
});

describe("schema routing of the parts directly in ppt/ by their root (#887)", { timeout: 120_000 }, () => {
    const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PML_CT = "application/vnd.openxmlformats-officedocument.presentationml";
    const PPT_NS =
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:r="${OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
    const STRICT_PPT_NS =
        `xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" ` +
        `xmlns:r="http://purl.oclc.org/ooxml/officeDocument" ` +
        `xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"`;
    // The same shape tree and colour map the minimal PPTX's slide parts hold.
    const PPT_TREE =
        `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr/></p:spTree></p:cSld>`;
    const PPT_CLR_MAP =
        `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ` +
        `accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`;

    // The six part kinds the minimal PPTX lacks, each the smallest part
    // pml.xsd accepts, related from the part the Open XML SDK part catalogue
    // relates them from (NotesSlidePart.json, NotesMasterPart.json,
    // HandoutMasterPart.json, SlideCommentsPart.json, UserDefinedTagsPart.json
    // and the slide synchronization data part): #887 stored them directly in
    // `ppt/`, where a part's identity comes from its relationship and content
    // type, not from the folder it sits in.
    const NOTES_SLIDE = `${HEAD}<p:notes ${PPT_NS}>${PPT_TREE}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
    const NOTES_MASTER = `${HEAD}<p:notesMaster ${PPT_NS}>${PPT_TREE}${PPT_CLR_MAP}</p:notesMaster>`;
    const HANDOUT_MASTER = `${HEAD}<p:handoutMaster ${PPT_NS}>${PPT_TREE}${PPT_CLR_MAP}</p:handoutMaster>`;
    const LEGACY_COMMENTS = `${HEAD}<p:cmLst ${PPT_NS}></p:cmLst>`;
    const TAGS = `${HEAD}<p:tagLst ${PPT_NS}></p:tagLst>`;
    const SLIDE_UPDATE_INFO =
        `${HEAD}<p:sldSyncPr ${PPT_NS} serverSldId="1" serverSldModifiedTime="2026-01-01T00:00:00Z" ` +
        `clientInsertedTime="2026-01-01T00:00:00Z"></p:sldSyncPr>`;

    /** The six kinds stored directly in `ppt/`: the part, the sidecar that relates it, and its relationship and content-type name. */
    const DIRECT_EXTRA_PARTS = [
        ["ppt/notesSlide1.xml", NOTES_SLIDE, "ppt/slides/_rels/slide1.xml.rels", "notesSlide"],
        ["ppt/notesMaster1.xml", NOTES_MASTER, "ppt/_rels/presentation.xml.rels", "notesMaster"],
        ["ppt/handoutMaster1.xml", HANDOUT_MASTER, "ppt/_rels/presentation.xml.rels", "handoutMaster"],
        ["ppt/comment1.xml", LEGACY_COMMENTS, "ppt/slides/_rels/slide1.xml.rels", "comments"],
        ["ppt/tag1.xml", TAGS, "ppt/_rels/presentation.xml.rels", "tags"],
        ["ppt/slideUpdateInfo1.xml", SLIDE_UPDATE_INFO, "ppt/slides/_rels/slide1.xml.rels", "slideUpdateInfo"],
    ] as const;

    /** The three part kinds the minimal PPTX already holds in its standard folders, each moved directly into `ppt/`. */
    const MOVED_PARTS = ["ppt/slides/slide1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slideMasters/slideMaster1.xml"] as const;

    /** The minimal PPTX plus `part`, related from `sidecar`'s owner and typed by its own Override. */
    const pptxHolding = (part: string, xml: string, sidecar: string, kind: string): Record<string, string> => {
        const owner = sidecar.replace("/_rels/", "/").slice(0, -".rels".length);
        return {
            ...MINIMAL_PPTX_PARTS,
            [part]: xml,
            "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/${part}" ContentType="${PML_CT}.${kind}+xml"/></Types>`,
            ),
            [sidecar]: MINIMAL_PPTX_PARTS[sidecar]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdExtra" Type="${OFFICE_REL}/${kind}" ` +
                    `Target="${path.posix.relative(path.posix.dirname(owner), part)}"/></Relationships>`,
            ),
        };
    };

    /**
     * The minimal PPTX with the part stored at `from` moved directly into
     * `ppt/`: its sidecar moves with it, and every relationship target and
     * content-type Override that named the old name names the new one — the
     * same relocation the issue reproduced them by.
     */
    const pptxWithPartMoved = (from: string): Record<string, string> => {
        const base = path.posix.basename(from);
        const target = `ppt/${base}`;
        const fromSidecar = `${path.posix.dirname(from)}/_rels/${base}.rels`;
        const toSidecar = `ppt/_rels/${base}.rels`;
        const move = (name: string): string => (name === from ? target : name === fromSidecar ? toSidecar : name);
        const parts: Record<string, string> = {};
        for (const [name, xml] of Object.entries(MINIMAL_PPTX_PARTS)) {
            let content = xml;
            if (name === "[Content_Types].xml") content = content.replaceAll(`/${from}`, `/${target}`);
            if (name.endsWith(".rels")) {
                // Resolve each target against the owner it names today and
                // re-relativize it against the owner it will name.
                const owner = name === "_rels/.rels" ? "" : name.replace("/_rels/", "/").slice(0, -".rels".length);
                const moved = move(owner);
                content = content.replace(/Target="([^"]+)"/g, (_, rel: string) => {
                    const resolved = path.posix.normalize(path.posix.join(path.posix.dirname(owner), rel));
                    return `Target="${path.posix.relative(path.posix.dirname(moved), move(resolved))}"`;
                });
            }
            parts[move(name)] = content;
        }
        return parts;
    };

    /** The minimal PPTX with its main presentation renamed: the package relationship, the Override and the sidecar's name move with it. */
    const pptxWithRenamedPresentation = (): Record<string, string> => {
        const parts: Record<string, string> = {};
        for (const [name, xml] of Object.entries(MINIMAL_PPTX_PARTS)) {
            parts[name.replaceAll("presentation.xml", "presentation2.xml")] = xml.replaceAll("presentation.xml", "presentation2.xml");
        }
        return parts;
    };

    /** The part's XML with a `<rootPrefix:bogusDefect/>` as the root's first child. */
    const withBogusDefect = (xml: string): string => {
        const root = xml.indexOf("<", xml.indexOf("?>") + 1) + 1;
        const prefix = xml.slice(root, xml.indexOf(">", root)).split(":")[0]!;
        const at = xml.indexOf(">", root) + 1;
        return `${xml.slice(0, at)}<${prefix}:bogusDefect/>${xml.slice(at)}`;
    };

    it.each([
        // The root decides, not the name: a renamed main presentation, a
        // slide under a name of its own, and a file name spelled like an
        // Object.prototype key all route to pml.xsd when the root declares
        // itself a global element of its closure.
        ["ppt/presentation2.xml", MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!],
        ["ppt/slide1.xml", MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!],
        ["ppt/constructor.xml", MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!],
        ["ppt/__proto__.xml", MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!],
    ])("maps %s to the ppt schema key by its root", async (part, xml) => {
        await withTempDir(async (dir) => {
            const pkg = path.join(dir, "pkg");
            await fs.mkdir(path.dirname(path.join(pkg, part)), { recursive: true });
            await fs.writeFile(path.join(pkg, part), xml, "utf-8");
            expect(new KeyValidator({ unpackedDir: pkg }).schemaMappingKey(part)).toBe("ppt");
            expect(new KeyValidator({ unpackedDir: pkg }).schemaFile(part)).toBe("ISO-IEC29500-4_2016/pml.xsd");
        });
    });

    it("leaves a part directly in ppt/ unmapped when no bundled schema declares its root, or it does not parse", async () => {
        await withTempDir(async (dir) => {
            const pkg = path.join(dir, "pkg");
            await fs.mkdir(path.join(pkg, "ppt"), { recursive: true });
            await fs.writeFile(
                path.join(pkg, "ppt/authors.xml"),
                `${HEAD}<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"></p188:authorLst>`,
                "utf-8",
            );
            await fs.writeFile(path.join(pkg, "ppt/broken.xml"), "<p:sld", "utf-8");
            const validator = new KeyValidator({ unpackedDir: pkg });
            expect(validator.schemaMappingKey("ppt/authors.xml")).toBeNull();
            expect(validator.schemaMappingKey("ppt/broken.xml")).toBeNull();
            expect(validator.schemaMappingKey("ppt/absent.xml")).toBeNull();
        });
    });

    it("keeps the earlier basename routes before the ppt/ root route", () => {
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey("ppt/core.xml")).toBe("core.xml");
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey("ppt/app.xml")).toBe("app.xml");
        expect(new KeyValidator({ unpackedDir: "/pkg" }).schemaMappingKey("ppt/custom.xml")).toBe("custom.xml");
    });

    it.each([undefined, "lenient", "strict"] as const)(
        "reports a schema violation in a renamed ppt/presentation2.xml under %s",
        async (profile) => {
            await withTempDir(async (dir) => {
                const parts = pptxWithRenamedPresentation();
                const file = path.join(dir, "renamed-defect.pptx");
                await fs.writeFile(
                    file,
                    await buildPptx({ ...parts, "ppt/presentation2.xml": withBogusDefect(parts["ppt/presentation2.xml"]!) }),
                );
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                const errors = result.issues.filter((issue) => issue.code === "xsd-error");
                expect(errors.length).toBeGreaterThan(0);
                expect(errors.every((issue) => issue.path === "ppt/presentation2.xml")).toBe(true);
                expect(errors.some((issue) => issue.message.includes("bogusDefect"))).toBe(true);
                expect(result.xsdWalk.parts).toContain("ppt/presentation2.xml");
            });
        },
    );

    describe.each(["lenient", "strict"] as const)("under %s", (profile) => {
        it("walks a clean renamed ppt/presentation2.xml without findings", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "renamed-control.pptx");
                await fs.writeFile(file, await buildPptx(pptxWithRenamedPresentation()));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).toContain("ppt/presentation2.xml");
                expect(result.issues.filter((issue) => issue.path === "ppt/presentation2.xml")).toEqual([]);
            });
        });

        it.each(MOVED_PARTS)("reports a schema violation in the relocated %s", async (from) => {
            await withTempDir(async (dir) => {
                const part = `ppt/${path.posix.basename(from)}`;
                const parts = pptxWithPartMoved(from);
                const file = path.join(dir, "moved-defect.pptx");
                await fs.writeFile(file, await buildPptx({ ...parts, [part]: withBogusDefect(parts[part]!) }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                const errors = result.issues.filter((issue) => issue.code === "xsd-error");
                expect(errors.length).toBeGreaterThan(0);
                expect(errors.every((issue) => issue.path === part)).toBe(true);
                expect(errors.some((issue) => issue.message.includes("bogusDefect"))).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
            });
        });

        it.each(MOVED_PARTS)("walks a clean relocated %s without findings", async (from) => {
            await withTempDir(async (dir) => {
                const part = `ppt/${path.posix.basename(from)}`;
                const file = path.join(dir, "moved-control.pptx");
                await fs.writeFile(file, await buildPptx(pptxWithPartMoved(from)));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
                expect(result.issues.filter((issue) => issue.path === part)).toEqual([]);
            });
        });

        it.each(DIRECT_EXTRA_PARTS)("reports a schema violation in %s", async (part, xml, sidecar, kind) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "direct-defect.pptx");
                await fs.writeFile(file, await buildPptx({ ...pptxHolding(part, xml, sidecar, kind), [part]: withBogusDefect(xml) }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                const errors = result.issues.filter((issue) => issue.code === "xsd-error");
                expect(errors.length).toBeGreaterThan(0);
                expect(errors.every((issue) => issue.path === part)).toBe(true);
                expect(errors.some((issue) => issue.message.includes("bogusDefect"))).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
            });
        });

        it.each(DIRECT_EXTRA_PARTS)("walks a valid %s without findings", async (part, xml, sidecar, kind) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "direct-control.pptx");
                await fs.writeFile(file, await buildPptx(pptxHolding(part, xml, sidecar, kind)));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.xsdWalk.parts).toContain(part);
                expect(result.issues.filter((issue) => issue.path === part)).toEqual([]);
            });
        });
    });

    it("routes a Strict slide stored as ppt/slide1.xml to the ISO Strict schema under strict, and skips it under lenient", async () => {
        await withTempDir(async (dir) => {
            const strictSlide = `${HEAD}<p:sld ${STRICT_PPT_NS}>${PPT_TREE}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
            const parts = { ...pptxWithPartMoved("ppt/slides/slide1.xml"), "ppt/slide1.xml": withBogusDefect(strictSlide) };
            const file = path.join(dir, "strict-slide.pptx");
            await fs.writeFile(file, await buildPptx(parts));

            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(
                strict.issues.some(
                    (issue) => issue.code === "xsd-error" && issue.path === "ppt/slide1.xml" && issue.message.includes("bogusDefect"),
                ),
            ).toBe(true);
            expect(strict.xsdWalk.parts).toContain("ppt/slide1.xml");

            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.valid).toBe(true);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
            expect(lenient.issues.filter((issue) => issue.code === "xsd-strict-skipped")).toHaveLength(1);
            expect(lenient.xsdWalk.parts).not.toContain("ppt/slide1.xml");
        });
    });

    // A validator walked twice routes each part by the root it holds now:
    // the root memo is cleared at the start of every walk, as the main
    // document list is (#657), so a part rewritten between the walks is not
    // routed by the root it held before (#902).
    it("routes a part rewritten between two walks of one validator by its new root (#902)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const validator = new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" });
            const first = await validator.validateAgainstXsd();
            expect(first.issues.filter((issue) => issue.path === "ppt/extra.xml").map((issue) => issue.code)).toEqual(["xsd-no-schema"]);

            await fs.writeFile(path.join(dir, "ppt/extra.xml"), withBogusDefect(MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!), "utf-8");
            const second = await validator.validateAgainstXsd();
            expect(second.valid).toBe(false);
            expect(
                second.issues.some(
                    (issue) => issue.code === "xsd-error" && issue.path === "ppt/extra.xml" && issue.message.includes("bogusDefect"),
                ),
            ).toBe(true);
            expect(validator.xsdWalk.parts).toContain("ppt/extra.xml");
        });
    });

    // A reused validator routes each direct validateFileAgainstXsd call by
    // the root the part holds at that call too: the public method drops the
    // file's root memo entry before it validates, as a walk clears the memo
    // wholesale, so a part rewritten between two direct calls is not routed
    // by the root the first call read (#910).
    it("routes a part rewritten between two direct validateFileAgainstXsd calls by its new root (#910)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const validator = new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" });
            const extra = path.join(dir, "ppt/extra.xml");
            expect((await validator.validateFileAgainstXsd(extra)).valid).toBe(null);

            await fs.writeFile(extra, withBogusDefect(MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!), "utf-8");
            const second = await validator.validateFileAgainstXsd(extra);
            expect(second.valid).toBe(false);
            expect([...second.errors].some((error) => error.includes("bogusDefect"))).toBe(true);
        });
    });

    it("routes a part rewritten from a slide to an unknown root between two direct calls by its new root (#910)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": withBogusDefect(MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!),
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const validator = new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" });
            const extra = path.join(dir, "ppt/extra.xml");
            const first = await validator.validateFileAgainstXsd(extra);
            expect(first.valid).toBe(false);
            expect([...first.errors].some((error) => error.includes("bogusDefect"))).toBe(true);

            await fs.writeFile(extra, `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`, "utf-8");
            expect((await validator.validateFileAgainstXsd(extra)).valid).toBe(null);
        });
    });

    // The walk still reads each part's root once: its route through the
    // shared memo must not become the public method, whose per-call drop
    // would re-read every part's root mid-walk. The `node:fs` namespace
    // cannot be spied on, so count the parses of one marked part's bytes
    // instead: a mapped part is parsed once for its root (`rootQName`) and
    // once for XSD preprocessing, so two parses over one walk pin the root
    // read at one, and two more for a direct call after the walk show the
    // call read the root again instead of taking the walk's (#910).
    it("reads a walked part's root once per walk while direct calls re-read it (#910)", async () => {
        await withTempDir(async (dir) => {
            const marker = "read-once-marker";
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!.replace('name=""', `name="${marker}"`),
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const spy = vi.spyOn(xmlHelpers, "parseXml");
            const markedParses = () => spy.mock.calls.filter((call) => String(call[0]).includes(marker)).length;
            try {
                const validator = new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" });
                const result = await validator.validateAgainstXsd();
                expect(result.valid).toBe(true);
                expect(validator.xsdWalk.parts).toContain("ppt/extra.xml");
                expect(markedParses()).toBe(2);

                const direct = await validator.validateFileAgainstXsd(path.join(dir, "ppt/extra.xml"));
                expect(direct.valid).toBe(true);
                expect(markedParses()).toBe(4);
            } finally {
                spy.mockRestore();
            }
        });
    });

    // The walk keeps calling the public method, so a subclass that
    // overrides validateFileAgainstXsd still sees every part the walk
    // validates; calling super from the override keeps the walk's memo, so
    // the root is still read once (#910).
    it("calls an overriding validateFileAgainstXsd for every part a walk validates, reading each root once (#910)", async () => {
        await withTempDir(async (dir) => {
            const marker = "override-marker";
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!.replace('name=""', `name="${marker}"`),
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const seen: string[] = [];
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    seen.push(this.relPath(file).split(path.sep).join("/"));
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            const spy = vi.spyOn(xmlHelpers, "parseXml");
            try {
                const result = await validator.validateAgainstXsd();
                expect(result.valid).toBe(true);
                expect(seen).toContain("ppt/extra.xml");
                expect(seen).toEqual(validator.xsdWalk.parts);
                expect(spy.mock.calls.filter((call) => String(call[0]).includes(marker))).toHaveLength(2);
            } finally {
                spy.mockRestore();
            }
        });
    });

    // An async descendant keeps the walk's store after its call settles,
    // but not its right to reuse the root read for that call (#926).
    it.each([false, true])("re-reads a rewritten part in a task started by a settled walk call (rejected: %s) (#926)", async (reject) => {
        await withTempDir(async (dir) => {
            const parts = { ...MINIMAL_PPTX_PARTS, "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]! };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const extra = path.join(dir, "ppt/extra.xml");
            let release!: () => void;
            const gate = new Promise<void>((resolve) => {
                release = resolve;
            });
            let later: ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]> | undefined;
            const failure = new Error("walk call rejected");
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    if (file === extra && !later) {
                        later = (async () => {
                            await gate;
                            return super.validateFileAgainstXsd(file);
                        })();
                        if (reject) throw failure;
                    }
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            if (reject) await expect(validator.validateAgainstXsd()).rejects.toBe(failure);
            else expect((await validator.validateAgainstXsd()).valid).toBe(true);

            await fs.writeFile(extra, `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`, "utf-8");
            expect(later).toBeDefined();
            release();
            const inherited = await later!;
            const direct = await validator.validateFileAgainstXsd(extra);
            expect(direct.valid).toBe(null);
            expect(inherited).toEqual(direct);
        });
    });

    // Only the part the walk called for keeps its memo: a direct call for
    // another part reads its root again, even inside the override (#926).
    it("routes another part rewritten between two direct calls inside a walk by its new root (#926)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!,
                "ppt/zzz.xml": `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const extra = path.join(dir, "ppt/extra.xml");
            const other = path.join(dir, "ppt/zzz.xml");
            let second: Awaited<ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]>> | undefined;
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    if (file === extra) {
                        expect((await super.validateFileAgainstXsd(other)).valid).toBe(null);
                        await fs.writeFile(other, withBogusDefect(MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!), "utf-8");
                        second = await super.validateFileAgainstXsd(other);
                    }
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            await validator.validateAgainstXsd();
            expect(second?.valid).toBe(false);
            expect([...second!.errors].some((error) => error.includes("bogusDefect"))).toBe(true);
        });
    });

    // A second validator may read the same path during the first one's
    // walk; that call still drops its own old root memo (#926).
    it("re-reads a rewritten part in another validator called inside a walk (#926)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const extra = path.join(dir, "ppt/extra.xml");
            const other = new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" });
            expect((await other.validateFileAgainstXsd(extra)).valid).toBe(null);
            await fs.writeFile(extra, withBogusDefect(MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!), "utf-8");
            let direct: Awaited<ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]>> | undefined;
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    if (file === extra) direct = await other.validateFileAgainstXsd(file);
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            await validator.validateAgainstXsd();
            expect(direct?.valid).toBe(false);
            expect([...direct!.errors].some((error) => error.includes("bogusDefect"))).toBe(true);
        });
    });

    // The walk hands its root memo to the first base call once. A second
    // call for the same part reads the root again, whether the override
    // calls super again, recurses, or awaits a task it starts (#941).
    it.each(["super-twice", "recursive-this", "awaited-descendant"] as const)(
        "reads the root afresh when the walk's own call validates a rewritten part again: %s (#941)",
        async (kind) => {
            await withTempDir(async (dir) => {
                const parts = { ...MINIMAL_PPTX_PARTS, "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]! };
                for (const [part, xml] of Object.entries(parts)) {
                    await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                    await fs.writeFile(path.join(dir, part), xml, "utf-8");
                }
                const extra = path.join(dir, "ppt/extra.xml");
                let entered = false;
                let second: Awaited<ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]>> | undefined;
                const validator = new (class extends PPTXSchemaValidator {
                    override async validateFileAgainstXsd(file: string): ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]> {
                        if (file === extra && !entered) {
                            entered = true;
                            expect((await super.validateFileAgainstXsd(file)).valid).toBe(true);
                            await fs.writeFile(
                                file,
                                `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
                                "utf-8",
                            );
                            if (kind === "recursive-this") second = await this.validateFileAgainstXsd(file);
                            else if (kind === "awaited-descendant") second = await (async () => super.validateFileAgainstXsd(file))();
                            else second = await super.validateFileAgainstXsd(file);
                            return second;
                        }
                        return super.validateFileAgainstXsd(file);
                    }
                })({ unpackedDir: dir, profile: "strict" });
                const result = await validator.validateAgainstXsd();
                expect(second?.valid).toBe(null);
                expect(second?.errors.size).toBe(0);
                expect(result.valid).toBe(true);
                expect(second).toEqual(await validator.validateFileAgainstXsd(extra));
            });
        },
    );

    // Claiming the memo must not drop it from the first base call: an
    // override that calls super once still reads each root once (#941).
    it("keeps the walk's root for an override calling super once without rewriting the part (#941)", async () => {
        await withTempDir(async (dir) => {
            const marker = "claim-once-marker";
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!.replace('name=""', `name="${marker}"`),
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const seen: string[] = [];
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    seen.push(this.relPath(file).split(path.sep).join("/"));
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            const spy = vi.spyOn(xmlHelpers, "parseXml");
            try {
                const result = await validator.validateAgainstXsd();
                expect(result).toEqual({ valid: true, issues: [] });
                expect(seen).toContain("ppt/extra.xml");
                expect(seen).toEqual(validator.xsdWalk.parts);
                expect(spy.mock.calls.filter((call) => String(call[0]).includes(marker))).toHaveLength(2);
            } finally {
                spy.mockRestore();
            }
        });
    });

    // The claim belongs to the walk's own part: a call for an earlier part
    // made first inside the walk's call reads that part's root again (#941).
    it("reads another part's root afresh when it is the first base call inside the walk's own call (#941)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "ppt/aaa.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!,
                "ppt/extra.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"]!,
            };
            for (const [part, xml] of Object.entries(parts)) {
                await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await fs.writeFile(path.join(dir, part), xml, "utf-8");
            }
            const earlier = path.join(dir, "ppt/aaa.xml");
            const extra = path.join(dir, "ppt/extra.xml");
            let first: Awaited<ReturnType<PPTXSchemaValidator["validateFileAgainstXsd"]>> | undefined;
            const validator = new (class extends PPTXSchemaValidator {
                override async validateFileAgainstXsd(file: string) {
                    if (file === extra && first === undefined) {
                        await fs.writeFile(
                            earlier,
                            `${HEAD}<p:unknown xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
                            "utf-8",
                        );
                        first = await super.validateFileAgainstXsd(earlier);
                    }
                    return super.validateFileAgainstXsd(file);
                }
            })({ unpackedDir: dir, profile: "strict" });
            await validator.validateAgainstXsd();
            // The walk read the earlier part's slide root before it reached
            // ppt/extra.xml, so only a fresh read routes the rewrite as unmapped.
            expect(validator.xsdWalk.parts.indexOf("ppt/aaa.xml")).toBeLessThan(validator.xsdWalk.parts.indexOf("ppt/extra.xml"));
            expect(validator.xsdWalk.parts.indexOf("ppt/aaa.xml")).toBeGreaterThan(-1);
            expect(first?.valid).toBe(null);
            expect(first?.errors.size).toBe(0);
        });
    });
});

describe("the global element declarations the ppt/ root route reads (#887)", () => {
    const PML = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_PML = "http://purl.oclc.org/ooxml/presentationml/main";
    const STRICT_DML = "http://purl.oclc.org/ooxml/drawingml/main";
    const schemas = (schema: string): string => path.join(defaultSchemasDir(), schema);

    it.each([
        // The roots the route reads them for: the slide, the presentation,
        // and the DrawingML table styles pml.xsd reaches through its
        // dml-main.xsd import — declarations the bundled files own, so the
        // route never carries a hand-listed copy of them.
        ["ISO-IEC29500-4_2016/pml.xsd", `{${PML}}sld`],
        ["ISO-IEC29500-4_2016/pml.xsd", `{${PML}}presentation`],
        ["ISO-IEC29500-4_2016/pml.xsd", `{${DML}}tblStyleLst`],
        ["iso-strict/pml.xsd", `{${STRICT_PML}}sld`],
        ["iso-strict/pml.xsd", `{${STRICT_PML}}presentation`],
        ["iso-strict/pml.xsd", `{${STRICT_DML}}tblStyleLst`],
    ] as const)("declares %s's %s", (schema, key) => {
        expect(schemaGlobals(schemas(schema)).has(key)).toBe(true);
    });

    it("declares no Microsoft authorLst in either conformance class", () => {
        const key = "{http://schemas.microsoft.com/office/powerpoint/2018/8/main}authorLst";
        expect(schemaGlobals(schemas("ISO-IEC29500-4_2016/pml.xsd")).has(key)).toBe(false);
        expect(schemaGlobals(schemas("iso-strict/pml.xsd")).has(key)).toBe(false);
    });

    it("reads each entry schema's closure once", () => {
        expect(schemaGlobals(schemas("ISO-IEC29500-4_2016/pml.xsd"))).toBe(schemaGlobals(schemas("ISO-IEC29500-4_2016/pml.xsd")));
    });

    it("follows a local import and include, and skips a URL or missing schemaLocation", async () => {
        await withTempDir(async (dir) => {
            const XSD = "http://www.w3.org/2001/XMLSchema";
            await fs.writeFile(
                path.join(dir, "root.xsd"),
                `<xsd:schema xmlns:xsd="${XSD}" targetNamespace="urn:root">` +
                    `<xsd:import namespace="urn:remote" schemaLocation="http://example.com/remote.xsd"/>` +
                    `<xsd:import namespace="urn:missing" schemaLocation="missing.xsd"/>` +
                    `<xsd:import namespace="urn:local" schemaLocation="local.xsd"/>` +
                    `<xsd:include schemaLocation="chameleon.xsd"/>` +
                    `<xsd:element name="root"/></xsd:schema>`,
                "utf-8",
            );
            await fs.writeFile(
                path.join(dir, "local.xsd"),
                `<xsd:schema xmlns:xsd="${XSD}" targetNamespace="urn:local"><xsd:element name="localRoot"/></xsd:schema>`,
                "utf-8",
            );
            // An include that declares no target namespace of its own (the
            // chameleon case) declares its globals in the including one's.
            await fs.writeFile(
                path.join(dir, "chameleon.xsd"),
                `<xsd:schema xmlns:xsd="${XSD}"><xsd:element name="chameleon"/></xsd:schema>`,
                "utf-8",
            );
            expect([...schemaGlobals(path.join(dir, "root.xsd"))].sort()).toEqual([
                "{urn:local}localRoot",
                "{urn:root}chameleon",
                "{urn:root}root",
            ]);
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

    // The minimal PPTX's slide, layout and master map to pml.xsd since #864
    // routed the PresentationML parts one folder below `ppt/`; only the customXml
    // payload is unmapped, so it alone is audited.
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
            expect(result.issues.filter((issue) => issue.code === "xsd-no-schema")).toEqual([audit("customXml/item1.xml")]);
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
