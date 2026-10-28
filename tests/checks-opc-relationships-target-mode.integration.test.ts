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
import { referencedChartSpecimenParts } from "./support/chart-specimens";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const CODE = "opc-relationships-target-mode";
const PACKAGE = "_rels/.rels";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const HEADER = "word/_rels/header1.xml.rels";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const append = (xml: string, child: string) => xml.replace("</Relationships>", `${child}</Relationships>`);
const hyperlink = (mode: string) => `<Relationship Id="rId99" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="${mode}"/>`;

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
