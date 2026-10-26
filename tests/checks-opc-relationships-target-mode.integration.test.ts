/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { CHART_DRAWING, referencedChartSpecimenParts } from "./support/chart-specimens";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const CODE = "opc-relationships-target-mode";
const PACKAGE = "_rels/.rels";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const HEADER = "word/_rels/header1.xml.rels";
const GLOSSARY = "word/glossary/document.xml";
const GLOSSARY_SIDECAR = "word/glossary/_rels/document.xml.rels";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const append = (xml: string, child: string) => xml.replace("</Relationships>", `${child}</Relationships>`);
const hyperlink = (mode: string) => `<Relationship Id="rId99" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="${mode}"/>`;

/** Session 57 (s57f/s57g): a reached glossary with one docPart and an optional sidecar. */
const glossaryParts = (body: string, mode?: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/>`),
    [GLOSSARY]:
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        '<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>' +
        `<w:docPartBody>${body}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`,
    ...(mode === undefined
        ? {}
        : {
              [GLOSSARY_SIDECAR]: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hyperlink(mode)}</Relationships>`,
          }),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/${GLOSSARY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>`,
    ),
});

/** Session 55 (s55b): a real chart the drawing names, with only TargetMode changed. */
const chart = async (mode?: string) => {
    const parts = referencedChartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml", false);
    if (mode !== undefined) {
        parts[SIDECAR] = parts[SIDECAR]!.replace('Target="charts/chart1.xml"/>', `Target="charts/chart1.xml" TargetMode="${mode}"/>`);
    }
    return buildDocx(parts);
};

/** Session 55 (s55c): the fixture is read only; change diagramData rId35 in memory. */
const diagram = async (mode?: string) => {
    const zip = await JSZip.loadAsync(await readFile(new URL("./fixtures/working/word-clean-strict01.docx", import.meta.url)));
    if (mode !== undefined) {
        const rels = await zip.file(SIDECAR)!.async("string");
        const original =
            '<Relationship Id="rId35" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData" Target="diagrams/data1.xml"/>';
        expect(rels).toContain(original);
        zip.file(SIDECAR, rels.replace(original, original.replace("/>", ` TargetMode="${mode}"/>`)));
    }
    return zip.generateAsync({ type: "nodebuffer" });
};

/** Session 55 (s55d/s55e): an unreferenced hyperlink in the main or package relationships part. */
const unreferenced = (part: string, mode: string) => buildDocx({ ...MINIMAL_PARTS, [part]: append(MINIMAL_PARTS[part]!, hyperlink(mode)) });

/** Session 55 (s55e): a referenced header whose own sidecar holds only the unreferenced hyperlink. */
const header = (mode: string) =>
    buildDocx({
        ...MINIMAL_PARTS,
        [MAIN]: MINIMAL_PARTS[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
            "<w:sectPr>",
            '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>',
        ),
        [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/>`),
        "word/header1.xml": `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>`,
        [HEADER]: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hyperlink(mode)}</Relationships>`,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
        ),
    });

const refused = [
    { label: "s55b chart external", make: () => chart("external"), part: SIDECAR },
    { label: "s55c diagramData external", make: () => diagram("external"), part: SIDECAR },
    { label: "s55d unreferenced hyperlink external", make: () => unreferenced(SIDECAR, "external"), part: SIDECAR },
    { label: "s55e unreferenced hyperlink trailing space", make: () => unreferenced(SIDECAR, "External "), part: SIDECAR },
    { label: "s55e unreferenced hyperlink empty", make: () => unreferenced(SIDECAR, ""), part: SIDECAR },
    { label: "s55e package hyperlink external", make: () => unreferenced(PACKAGE, "external"), part: PACKAGE },
];
const clean = [
    { label: "s55b chart absent", make: () => chart() },
    { label: "s55b chart Internal", make: () => chart("Internal") },
    { label: "s55c diagramData absent", make: () => diagram() },
    { label: "s55c diagramData Internal", make: () => diagram("Internal") },
    { label: "s55d main hyperlink External", make: () => unreferenced(SIDECAR, "External") },
    { label: "s55e package hyperlink External", make: () => unreferenced(PACKAGE, "External") },
    { label: "s55e header hyperlink external", make: () => header("external") },
    { label: "s55e header hyperlink External", make: () => header("External") },
];

describe(CODE, () => {
    for (const selection of ["full word-valid", "named word-valid"] as const) {
        const options = selection === "full word-valid" ? {} : { checks: [CODE] };
        describe(selection, () => {
            it.each(refused)("refuses $label (#753)", async ({ make, part }) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, "specimen.docx");
                    await writeFile(file, await make());
                    const result = await validate(file, { profile: "word-valid", ...options });
                    expect(result.valid, result.issues.map((issue) => issue.message).join("\n")).toBe(false);
                    expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([
                        expect.objectContaining({ code: CODE, path: part, severity: "error" }),
                    ]);
                });
            });
            it.each(clean)("keeps $label valid (#753)", async ({ make }) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, "specimen.docx");
                    await writeFile(file, await make());
                    const result = await validate(file, { profile: "word-valid", ...options });
                    expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                    expect(result.valid, result.issues.map((issue) => issue.message).join("\n")).toBe(true);
                });
            });
        });
    }

    it.each(refused)("reports $label as an error under strict", async ({ make, part }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await make());
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(expect.objectContaining({ code: CODE, path: part, severity: "error" }));
        });
    });
});

describe("hyperlink TargetMode ownership (#773)", () => {
    const parts = (mode: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        [MAIN]: MINIMAL_PARTS[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
            "<w:r><w:t>Hello</w:t></w:r>",
            '<w:hyperlink r:id="rId99"><w:r><w:t>link</w:t></w:r></w:hyperlink>',
        ),
        [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, hyperlink(mode)),
    });

    it.each(["external", "EXTERNAL", "External "])("reports only the target-mode error for %j in a full word-valid run", async (mode) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts(mode)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({ code: CODE, path: SIDECAR, message: expect.stringContaining("rId99") }),
            ]);
            expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
        });
    });

    for (const code of [CODE, "word-hyperlink-relationship"]) {
        it.each(["external", "EXTERNAL", "External "])(`pins the named ${code} run for %j`, async (mode) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts(mode)));
                const result = await validate(file, { profile: "word-valid", checks: [code] });
                expect(result.valid).toBe(code !== CODE);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual(
                    code === CODE ? [expect.objectContaining({ code: CODE, path: SIDECAR })] : [],
                );
            });
        });
    }

    it.each(["Internal", "absent", "another Type", "missing relationship"])("keeps the hyperlink finding for %s", async (shape) => {
        const specimen = parts("Internal");
        if (shape === "absent") specimen[SIDECAR] = specimen[SIDECAR]!.replace(' TargetMode="Internal"', "");
        if (shape === "another Type") {
            specimen[SIDECAR] = specimen[SIDECAR]!.replace(`${R}/hyperlink`, `${R}/image`).replace(
                'TargetMode="Internal"',
                'TargetMode="external"',
            );
        }
        if (shape === "missing relationship") specimen[SIDECAR] = MINIMAL_PARTS[SIDECAR]!;
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(specimen));
            for (const options of [{}, { checks: ["word-hyperlink-relationship"] }]) {
                const result = await validate(file, { profile: "word-valid", ...options });
                expect(result.valid).toBe(false);
                expect(result.issues).toContainEqual(
                    expect.objectContaining({
                        code: "word-hyperlink-relationship",
                        severity: "error",
                        message: expect.stringContaining("rId99"),
                    }),
                );
            }
        });
    });

    it.each(["header", "glossary"])("pins hyperlink finding ownership in a reached %s's sidecar (#775)", async (story) => {
        const specimen = { ...MINIMAL_PARTS };
        const part = story === "header" ? "word/header1.xml" : "word/glossary/document.xml";
        const sidecar = story === "header" ? HEADER : "word/glossary/_rels/document.xml.rels";
        const type = story === "header" ? "header" : "glossaryDocument";
        const contentType = story === "header" ? "header" : "document.glossary";
        const link = `<w:p><w:hyperlink r:id="rId99"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`;
        specimen[MAIN] = specimen[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `);
        if (story === "header") {
            specimen[MAIN] = specimen[MAIN]!.replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdS"/>');
        }
        specimen[SIDECAR] = append(specimen[SIDECAR]!, `<Relationship Id="rIdS" Type="${R}/${type}" Target="${part.slice(5)}"/>`);
        specimen[part] =
            story === "header"
                ? `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${link}</w:hdr>`
                : `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartBody>${link}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
        specimen[sidecar] =
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hyperlink("external")}</Relationships>`;
        specimen["[Content_Types].xml"] = specimen["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${contentType}+xml"/></Types>`,
        );
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(specimen));
            for (const options of [{}, { checks: ["word-hyperlink-relationship"] }]) {
                const result = await validate(file, { profile: "word-valid", ...options });
                const full = options.checks === undefined;
                expect(result.valid).toBe(story === "glossary" && !full);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual(
                    story === "header"
                        ? [expect.objectContaining({ code: "word-hyperlink-relationship", path: part })]
                        : full
                          ? [expect.objectContaining({ code: CODE, path: sidecar })]
                          : [],
                );
                expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(story === "header");
            }
        });
    });
});

describe("session 57 TargetMode and absent sidecars (#775)", () => {
    const text = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";

    it("refuses s57f's unreferenced glossary hyperlink with lowercase external", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(glossaryParts(text, "external")));
            for (const options of [{}, { checks: [CODE] }]) {
                const result = await validate(file, { profile: "word-valid", ...options });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                    expect.objectContaining({ code: CODE, path: GLOSSARY_SIDECAR, severity: "error" }),
                ]);
            }
        });
    });

    it.each([
        { label: "s57f glossary External calibration", make: () => buildDocx(glossaryParts(text, "External")) },
        { label: "s55e referenced header external", make: () => header("external") },
    ])("keeps $label valid", async ({ make }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await make());
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid, result.issues.map((issue) => issue.message).join("\n")).toBe(true);
            expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
        });
    });

    const drawing = CHART_DRAWING.replace(' w14:paraId="00000002" w14:textId="00000002"', "")
        .replace('r:id="rId9"', 'r:id="rId1"')
        .replace(
            "<w:p>",
            '<w:p xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">',
        );

    for (const story of ["footnotes", "comments", "glossary", "main"] as const) {
        it.each([true, false])(`pins ${story} with no sidecar, chart=%s (s57f/s57g)`, async (chart) => {
            const body = chart ? drawing : text;
            const part = story === "glossary" ? GLOSSARY : story === "main" ? MAIN : `word/${story}.xml`;
            const parts = story === "glossary" ? glossaryParts(body) : { ...MINIMAL_PARTS };
            if (story === "main") {
                delete parts[SIDECAR];
                parts[MAIN] = parts[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
                    "<w:sectPr>",
                    `${body}<w:sectPr>`,
                );
            } else if (story === "footnotes" || story === "comments") {
                const note = story === "footnotes" ? "footnote" : "comment";
                const id = story === "footnotes" ? "1" : "0";
                parts[MAIN] = parts[MAIN]!.replace("</w:r></w:p>", `</w:r><w:r><w:${note}Reference w:id="${id}"/></w:r></w:p>`);
                parts[SIDECAR] = append(parts[SIDECAR]!, `<Relationship Id="rIdN" Type="${R}/${story}" Target="${story}.xml"/>`);
                parts[part] =
                    `<w:${story} xmlns:w="${W}" xmlns:r="${R}"><w:${note} w:id="${id}"${story === "comments" ? ' w:author="a"' : ""}>${body}</w:${note}></w:${story}>`;
                parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
                    "</Types>",
                    `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}+xml"/></Types>`,
                );
            }
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid, result.issues.map((issue) => issue.message).join("\n")).toBe(!chart);
                expect(result.issues.filter((issue) => issue.code === "word-part-reference-relationship")).toEqual(
                    chart ? [expect.objectContaining({ path: part, severity: "error", message: expect.stringContaining("rId1") })] : [],
                );
            });
        });
    }
});
