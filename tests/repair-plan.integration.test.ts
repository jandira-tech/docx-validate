import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const documentWith = (inner: string): string => MINIMAL_PARTS["word/document.xml"].replace("<w:t>Hello</w:t>", inner);
const drawing = (inner: string): string => `<w:drawing>${inner}</w:drawing>`;
const graphic = `<a:graphic xmlns:a="${A}"><a:graphicData uri="urn:test"/></a:graphic>`;
const anchorAttributes = `xmlns:wp="${WP}" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"`;
const scalarDrawing = (value: string): string =>
    drawing(
        `<wp:anchor ${anchorAttributes}><wp:simplePos x="0" y="0"/>` +
            `<wp:positionH relativeFrom="column"><wp:posOffset>${value}</wp:posOffset></wp:positionH>` +
            `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
            `<wp:extent cx="1" cy="1"/><wp:wrapNone/><wp:docPr id="1" name="s"/>${graphic}</wp:anchor>`,
    );

interface OutcomeCase {
    name: string;
    code: string;
    parts: Record<string, string>;
    profile?: Profile;
    suffix?: ".docx" | ".docm" | ".pptx";
    before?: number;
    remaining: number;
    repairs: number;
    hint?: string;
    excludedHint?: RegExp;
}

const upperMain: Record<string, string> = { ...MINIMAL_PARTS, "word/DOCUMENT.xml": documentWith(scalarDrawing(" 1 ")) };
delete upperMain["word/document.xml"];
upperMain["_rels/.rels"] = upperMain["_rels/.rels"].replace("word/document.xml", "word/DOCUMENT.xml");
upperMain["[Content_Types].xml"] = upperMain["[Content_Types].xml"].replace("/word/document.xml", "/word/DOCUMENT.xml");
upperMain["word/_rels/DOCUMENT.xml.rels"] = upperMain["word/_rels/document.xml.rels"];
delete upperMain["word/_rels/document.xml.rels"];

const strictInline: Record<string, string> = {
    ...MINIMAL_PARTS,
    "word/document.xml": documentWith(drawing(`<wp:inline xmlns:wp="${WP}">${graphic}</wp:inline>`)),
};
for (const key of Object.keys(strictInline)) {
    strictInline[key] = strictInline[key]
        .replaceAll(W, "http://purl.oclc.org/ooxml/wordprocessingml/main")
        .replaceAll(WP, "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing")
        .replaceAll(A, "http://purl.oclc.org/ooxml/drawingml/main")
        .replaceAll(R, "http://purl.oclc.org/ooxml/officeDocument/relationships");
}

const customFiletime: Record<string, string> = {
    ...MINIMAL_PARTS,
    "docProps/custom.xml": `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9997-08002B2CF9AE}" pid="2" name="p"><vt:filetime/></property></Properties>`,
};
customFiletime["[Content_Types].xml"] = customFiletime["[Content_Types].xml"].replace(
    "</Types>",
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>',
);
customFiletime["_rels/.rels"] = customFiletime["_rels/.rels"].replace(
    "</Relationships>",
    `<Relationship Id="custom" Type="${R}/custom-properties" Target="docProps/custom.xml"/></Relationships>`,
);

const cases: OutcomeCase[] = [
    {
        name: "review 1: newline-only whitespace stays unavailable",
        code: "ws-missing-preserve",
        parts: { ...MINIMAL_PARTS, "word/document.xml": documentWith("<w:t>\nHello\n</w:t>") },
        remaining: 1,
        repairs: 0,
    },
    {
        name: "review 1: default-namespace text stays unavailable",
        code: "ws-missing-preserve",
        parts: { ...MINIMAL_PARTS, "word/document.xml": documentWith(`<t xmlns="${W}"> Hello </t>`) },
        remaining: 1,
        repairs: 0,
    },
    {
        name: "review 2: uppercase relationship suffix stays unavailable",
        code: "rels-empty-part",
        parts: {
            ...MINIMAL_PARTS,
            "word/_rels/styles.xml.RELS": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
        remaining: 1,
        repairs: 0,
    },
    {
        name: "review 3: uppercase main part scalar stays unavailable",
        code: "word-drawing-scalar-whitespace",
        parts: upperMain,
        remaining: 1,
        repairs: 0,
    },
    {
        name: "review 3: effective-view-only scalar stays unavailable",
        code: "word-drawing-scalar-whitespace",
        parts: {
            ...MINIMAL_PARTS,
            "word/document.xml": documentWith(
                scalarDrawing(
                    '<mc:AlternateContent><mc:Choice Requires="w14"> 1 </mc:Choice><mc:Fallback> 2 </mc:Fallback></mc:AlternateContent>',
                ),
            ),
        },
        remaining: 1,
        repairs: 0,
    },
    {
        name: "review 4: anchor scaffold marker stays unavailable",
        code: "xsd-error",
        profile: "strict",
        parts: { ...MINIMAL_PARTS, "word/document.xml": documentWith(drawing(`<wp:anchor ${anchorAttributes}>${graphic}</wp:anchor>`)) },
        remaining: 1,
        repairs: 0,
        hint: "required DrawingML children in schema order",
        excludedHint: /insert missing inline-picture scaffolding/,
    },
    {
        name: "review 4: ISO Strict inline scaffold marker stays unavailable",
        code: "xsd-error",
        profile: "strict",
        parts: strictInline,
        remaining: 1,
        repairs: 0,
        hint: "required DrawingML children in schema order",
    },
    ...(["word-valid", "lenient", "strict"] as const).map((profile): OutcomeCase => ({
        name: `review 5: PPTX Ignorable stays unavailable under ${profile}`,
        code: "ignorable-undeclared",
        suffix: ".pptx",
        profile,
        parts: {
            ...MINIMAL_PPTX_PARTS,
            "ppt/slides/slide1.xml": MINIMAL_PPTX_PARTS["ppt/slides/slide1.xml"].replace(
                "<p:sld ",
                '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p14" ',
            ),
        },
        remaining: 1,
        repairs: 0,
    })),
    {
        name: "review 6: Choice Requires clears indirectly through Ignorable repair",
        code: "word-choice-requires-undeclared",
        parts: {
            ...MINIMAL_PARTS,
            "word/document.xml": documentWith(
                '<mc:AlternateContent><mc:Choice Requires="w15"><w:t>A</w:t></mc:Choice><mc:Fallback><w:t>B</w:t></mc:Fallback></mc:AlternateContent>',
            ).replace('mc:Ignorable="w14"', 'mc:Ignorable="w14 w15"'),
        },
        remaining: 0,
        repairs: 1,
    },
    {
        name: "review 7: empty custom filetime receives usable manual guidance",
        code: "xsd-error",
        profile: "strict",
        parts: customFiletime,
        remaining: 1,
        repairs: 0,
        hint: "custom property a valid xs:dateTime value or remove/correct it",
        excludedHint: /whitespace|core document properties/,
    },
    {
        name: "partly cleared whitespace group reports exact counts",
        code: "ws-missing-preserve",
        before: 2,
        parts: { ...MINIMAL_PARTS, "word/document.xml": documentWith("<w:t> Hello </w:t><w:t>\nHello\n</w:t>") },
        remaining: 1,
        repairs: 1,
        hint: "add xml:space='preserve'",
    },
    {
        name: "unknown styles receive manual definition guidance",
        code: "style-reference-undefined",
        parts: { ...MINIMAL_PARTS, "word/document.xml": documentWith('<w:rPr><w:rStyle w:val="MyCustomStyle"/></w:rPr><w:t>Hello</w:t>') },
        remaining: 1,
        repairs: 0,
        hint: "define the missing styles",
        excludedHint: /leave unknown style IDs reported/,
    },
    {
        name: "DOCM reports the same cleared outcome",
        code: "ws-missing-preserve",
        suffix: ".docm",
        parts: {
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.ms-word.document.macroEnabled.main+xml",
            ),
            "word/document.xml": documentWith("<w:t> Hello </w:t>"),
        },
        remaining: 0,
        repairs: 1,
    },
    {
        name: "PPTX reports a cleared base repair",
        code: "rels-empty-part",
        suffix: ".pptx",
        profile: "lenient",
        parts: {
            ...MINIMAL_PPTX_PARTS,
            "ppt/slides/_rels/unused.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
        remaining: 0,
        repairs: 1,
    },
];

describe("repair plan observed outcomes (#622 review regressions)", () => {
    it.each(cases)("$name", { timeout: 60_000 }, async (c) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `specimen${c.suffix ?? ".docx"}`);
            await fs.writeFile(file, await (c.suffix === ".pptx" ? buildPptx(c.parts) : buildDocx(c.parts)));
            const options = { profile: c.profile ?? "word-valid" };
            const before = await validate(file, options);
            const after = await validate(file, { ...options, autoRepair: true });
            const findings = before.issues.filter((i) => i.code === c.code);
            expect(findings).toHaveLength(c.before ?? 1);
            expect(before.issues.some((i) => i.code?.startsWith("repair-plan"))).toBe(false);
            expect(after.repairs).toBe(c.repairs);
            const remaining = after.issues.filter((i) => i.code === c.code);
            expect(remaining).toHaveLength(c.remaining);
            const plans = after.issues.filter((i) => i.code?.startsWith("repair-plan") && i.message.includes(`[${c.code}]`));
            expect(plans).toHaveLength(1);
            const [plan] = plans;
            const cleared = findings.length - remaining.length;
            expect(plan).toMatchObject({
                severity: "info",
                path: findings[0].path,
                code: remaining.length === 0 ? "repair-plan" : cleared > 0 ? "repair-plan-partial" : "repair-plan-unavailable",
            });
            expect(plan.message).toContain(`Before repair found ${findings.length} [${c.code}] issue(s)`);
            expect(plan.message).toContain(`Auto-repair cleared ${cleared}; ${remaining.length} remain`);
            expect(plan.message).not.toContain("Auto-repair will");
            if (remaining.length > 0) expect(plan.message).toContain("Fix manually:");
            else expect(plan.message).not.toContain("Fix manually:");
            if (c.hint) expect(plan.message).toContain(c.hint);
            if (c.excludedHint) expect(plan.message).not.toMatch(c.excludedHint);
        });
    });
});
