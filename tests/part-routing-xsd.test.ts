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
