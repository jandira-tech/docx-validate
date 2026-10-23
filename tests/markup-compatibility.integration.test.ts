/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { NS, type Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const positions = ["H", "V"]
    .map(
        (axis) =>
            `<mc:AlternateContent><mc:Choice Requires="wp14"><wp:position${axis} relativeFrom="page"><wp14:pctPos${axis}Offset>45500</wp14:pctPos${axis}Offset></wp:position${axis}></mc:Choice><mc:Fallback><wp:position${axis} relativeFrom="page"><wp:posOffset>4576445</wp:posOffset></wp:position${axis}></mc:Fallback></mc:AlternateContent>`,
    )
    .join("");
const drawing = `<w:r><w:drawing><wp:anchor xmlns:wp="${WP}" xmlns:wp14="${WP14}" xmlns:a="${NS.A}" mc:Ignorable="wp14" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${positions}<wp:extent cx="1" cy="1"/><wp:wrapNone/><wp:docPr id="1" name="Drawing"/><a:graphic><a:graphicData uri="urn:test"/></a:graphic></wp:anchor></w:drawing></w:r>`;

const check = async (parts: Readonly<Record<string, string>>, profile: Profile = "strict") =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "test.docx");
        const bytes = await buildDocx(parts);
        await fs.writeFile(file, bytes);
        const result = await validate(file, { profile });
        expect(await fs.readFile(file)).toEqual(bytes);
        return result;
    });

const chartParts = (extension: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
        "</Types>",
        '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
    ),
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
        "</Relationships>",
        '<Relationship Id="chart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>',
    ),
    "word/charts/chart1.xml": `<c:chartSpace xmlns:c="${C}" xmlns:mc="${NS.MC}" xmlns:c16="urn:chart-extension" mc:Ignorable="c16"><c:chart><c:plotArea><c:layout/><c:pieChart/></c:plotArea></c:chart><c:extLst>${extension}</c:extLst></c:chartSpace>`,
});

describe("markup compatibility XSD integration", { timeout: 240_000 }, () => {
    it("diagnoses the real missing-Ignorable parts once each without changing package bytes", async () => {
        for (const [specimen, parts] of [
            ["empty.missing-content-type", ["word/document.xml"]],
            ["sample-document.missing-paraids", ["word/header1.xml", "word/footer1.xml"]],
        ] as const) {
            const file = new URL(`./fixtures/unknown/unknown/${specimen}.docx`, import.meta.url);
            const before = await fs.readFile(file);
            const result = await validate(file.pathname, { profile: "lenient" });
            expect(result.valid, JSON.stringify(result.issues)).toBe(true);
            const warnings = result.issues.filter((issue) => issue.code === "mce-extension-not-ignorable");
            expect(warnings.map((issue) => issue.path).sort()).toEqual([...parts].sort());
            for (const warning of warnings) {
                expect(warning.severity).toBe("warning");
                expect(warning.message).toContain(warning.path);
                expect(warning.message).toContain("http://schemas.microsoft.com/office/word/2010/wordml");
                expect(warning.message).toContain("w14:paraId");
                expect(warning.message).toContain("w14:textId");
            }
            expect(await fs.readFile(file)).toEqual(before);
        }
    });
    for (const profile of ["lenient", "strict", "word-valid"] as const) {
        it(`removes non-ignorable Office attributes before XSD with a stable warning under ${profile}`, async () => {
            const result = await check(
                {
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(
                        "<w:p ",
                        `<w:p xmlns:wp14="${WP14}" wp14:paraId="1" wp14:textId="2" `,
                    ),
                },
                profile,
            );
            expect(result.valid, JSON.stringify(result.issues)).toBe(true);
            expect(result.xsdWalk.parts).toContain("word/document.xml");
            const warnings = result.issues.filter((issue) => issue.code === "mce-extension-not-ignorable");
            expect(warnings).toHaveLength(1);
            expect(warnings[0]).toMatchObject({ severity: "warning", path: "word/document.xml" });
            for (const value of ["word/document.xml", WP14, "wp14:paraId", "wp14:textId", "2 attributes", "mc:Ignorable"]) {
                expect(warnings[0].message).toContain(value);
            }
            expect(result.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
        });
        it(`keeps truly unknown non-ignorable attributes and elements visible to XSD under ${profile}`, async () => {
            for (const content of [
                '<w:r xmlns:foo="urn:example:foo" foo:flag="keep"/>',
                '<foo:foreign xmlns:foo="urn:example:foo"/>',
                `<wp14:foreign xmlns:wp14="${WP14}"/>`,
            ]) {
                const result = await check(
                    { ...MINIMAL_PARTS, "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace("</w:p>", `${content}</w:p>`) },
                    profile,
                );
                expect(result.valid).toBe(profile === "word-valid");
                expect(result.issues).toContainEqual(
                    expect.objectContaining({
                        code: "xsd-error",
                        severity: profile === "word-valid" ? "warning" : "error",
                        path: "word/document.xml",
                    }),
                );
                expect(result.issues.some((issue) => issue.code === "mce-extension-not-ignorable")).toBe(false);
            }
        });
    }
    it("validates fallback anchor positions without manufacturing missing-position errors", async () => {
        const result = await check({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace("</w:p>", `${drawing}</w:p>`),
        });
        expect(result.xsdWalk.parts, JSON.stringify(result)).toContain("word/document.xml");
        expect(result.issues.filter((issue) => issue.code === "xsd-error" && /position[HV]/.test(issue.message))).toEqual([]);
        expect(result.valid, JSON.stringify(result.issues)).toBe(true);
    });
    it("surfaces compatibility errors with stable codes and the package part path", async () => {
        const result = await check({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(
                "</w:p>",
                '<mc:AlternateContent><mc:Choice Requires="missing"/></mc:AlternateContent><w:r xmlns:future="urn:future" mc:MustUnderstand="future"/></w:p>',
            ),
        });
        expect(result.issues).toContainEqual(
            expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
        );
        expect(result.issues).toContainEqual(
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/document.xml" }),
        );
        expect(result.valid).toBe(false);
    });
    it("keeps MustUnderstand failures blocking under word-valid", async () => {
        const result = await check(
            {
                ...MINIMAL_PARTS,
                "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(
                    "</w:p>",
                    '<w:r xmlns:future="urn:future" mc:MustUnderstand="future"/></w:p>',
                ),
            },
            "word-valid",
        );
        expect(result.issues).toContainEqual(
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/document.xml" }),
        );
        expect(result.valid).toBe(false);
    });
    it("preserves a chart's ignorable extension child through strict XSD validation", async () => {
        const result = await check(chartParts('<c:ext uri="test"><c16:uniqueId val="id"/></c:ext>'));
        expect(result.xsdWalk.parts).toContain("word/charts/chart1.xml");
        expect(result.issues.filter((issue) => issue.code === "xsd-error")).toEqual([]);
        expect(result.valid, JSON.stringify(result.issues)).toBe(true);
    });
    it("still rejects a genuinely empty chart extension", async () => {
        const result = await check(chartParts('<c:ext uri="test"/>'));
        expect(result.issues).toContainEqual(
            expect.objectContaining({
                code: "xsd-error",
                path: "word/charts/chart1.xml",
                message: expect.stringMatching(/chart}ext.*Missing child/),
            }),
        );
        expect(result.valid).toBe(false);
    });
    for (const specimen of ["k02_1", "k02_2", "k03_1", "k03_2", "k05_1", "k05_2", "k12_1"]) {
        it(`does not manufacture anchor or extension errors in ${specimen}`, async () => {
            const file = new URL(`./lane/jubarte-0.13.0/${specimen}.redline.docx`, import.meta.url);
            const before = await fs.readFile(file);
            const result = await validate(file.pathname, { profile: "strict" });
            expect(result.xsdWalk.started).toBe(true);
            expect(result.xsdWalk.parts).toContain("word/document.xml");
            expect(
                result.issues.filter(
                    (issue) =>
                        issue.code === "xsd-error" &&
                        (/position[HV]/.test(issue.message) || /chart}ext.*Missing child/.test(issue.message)),
                ),
            ).toEqual([]);
            expect(await fs.readFile(file)).toEqual(before);
        });
    }
    it("scopes mapped extension capabilities to the schema validating the part", async () => {
        const result = await check({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"].replace(
                "</w:p>",
                `<mc:AlternateContent xmlns:w15="${NS.W15}"><mc:Choice Requires="w15"><w:unknown/></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
                "</Types>",
                '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/></Types>',
            ),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
                "</Relationships>",
                '<Relationship Id="people" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/></Relationships>',
            ),
            "word/people.xml": `<w15:people xmlns:w15="${NS.W15}" xmlns:mc="${NS.MC}" mc:Ignorable="w15"/>`,
        });
        expect(result.xsdWalk.parts).toContain("word/people.xml");
        expect(result.valid, JSON.stringify(result.issues)).toBe(true);
    });
    it("does not manufacture an error for a part whose root element is ignored (#284)", async () => {
        const fixture = path.join(import.meta.dirname, "fixtures/external/open-xml-sdk/mcdoc.docx");
        for (const profile of ["strict", "lenient"] as const) {
            const result = await validate(fixture, { profile });
            const part = result.issues.filter((issue) => issue.path === "word/txbx1.xml");
            expect(
                part.filter((issue) => issue.severity === "error"),
                profile,
            ).toEqual([]);
            expect(
                part.map((issue) => [issue.severity, issue.code]),
                profile,
            ).toEqual([["info", "mce-root-ignored"]]);
            expect(result.xsdWalk.parts, profile).toContain("word/txbx1.xml");
        }
    });
});

describe("markup compatibility: standalone extension parts (#288)", { timeout: 120_000 }, () => {
    it("schema-validates word/commentsExtended.xml with its own w15 attributes intact", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "word", "commentsExtended.xml");
            await fs.mkdir(path.dirname(file), { recursive: true });
            await fs.writeFile(file, `<w15:commentsEx xmlns:w15="${NS.W15}"><w15:commentEx w15:paraId="00000001"/></w15:commentsEx>`);
            for (const profile of ["lenient", "strict"] as const) {
                const outcome = await new BaseSchemaValidator({ unpackedDir: dir, profile }).validateFileAgainstXsd(file);
                expect(outcome.valid, profile).toBe(true);
                expect(outcome.issues ?? [], profile).toEqual([]);
            }
        }));
});
