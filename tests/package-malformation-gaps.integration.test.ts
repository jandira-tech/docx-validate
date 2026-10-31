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
import { writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { validate } from "../src/scripts/office/validate";
import { withTempDir } from "../src/lib/run-cli";
import { compatibilityMainParts, compatibilityStoryParts } from "./support/word-compatibility-parts";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const customXml =
    '<x:root xmlns:x="urn:probe" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="missing"><x:child/></mc:Choice><mc:Fallback/></mc:AlternateContent></x:root>';
const cases: { name: string; code: string; strict: "error" | "warning"; mutate: (parts: Record<string, string>) => void }[] = [
    {
        name: "stale override",
        code: "ct-stale-override",
        // Word opened a stale Override clean (sessions 3c to 7, #414).
        strict: "warning",
        mutate: (p) => {
            p["[Content_Types].xml"] = p["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/absent.xml" ContentType="application/xml"/></Types>',
            );
        },
    },
    {
        name: "unclosed field",
        code: "word-field-unbalanced",
        strict: "error",
        mutate: (p) => {
            p["word/document.xml"] = p["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
            );
        },
    },
    {
        name: "unpaired move",
        code: "word-move-range-unpaired",
        strict: "error",
        mutate: (p) => {
            p["word/document.xml"] = p["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                '<w:moveFromRangeStart w:id="10" w:name="a" w:author="Probe" w:date="2026-10-30T00:00:00Z"/><w:r><w:t>Hello</w:t></w:r><w:moveFromRangeEnd w:id="10"/>',
            );
        },
    },
    {
        name: "custom XML Requires",
        code: "mce-malformed",
        strict: "error",
        mutate: (p) => {
            p["customXml/item.xml"] = customXml;
            p["word/_rels/document.xml.rels"] = p["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="custom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item.xml"/></Relationships>',
            );
        },
    },
];
const run = async (parts: Record<string, string>, profile: "word-valid" | "strict" | "lenient", checks?: readonly string[]) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "probe.docx");
        await writeFile(file, await buildDocx(parts));
        return validate(file, { profile, ...(checks === undefined ? {} : { checks }) });
    });

describe("package malformation gap integration", () => {
    for (const row of cases)
        it(`${row.name}: named word-valid selection agrees with full validation`, async () => {
            const parts = { ...MINIMAL_PARTS };
            row.mutate(parts);
            const id = row.code === "mce-malformed" ? "xml-package-markup-compatibility" : row.code;
            const result = await run(parts, "word-valid", [id]);
            expect(result.issues).toContainEqual(expect.objectContaining({ code: row.code, severity: "warning" }));
            expect(result.valid).toBe(true);
            expect(result.metadata?.executedChecks).toContain(id);
        });
    for (const row of cases)
        for (const profile of ["word-valid", "lenient", "strict"] as const) {
            it(`${row.name}: ${profile}`, async () => {
                const parts = { ...MINIMAL_PARTS };
                row.mutate(parts);
                const result = await run(parts, profile);
                if (profile === "lenient") {
                    expect(result.issues.some((i) => i.code === row.code)).toBe(false);
                    expect(result.valid).toBe(true);
                } else {
                    expect(result.issues).toContainEqual(
                        expect.objectContaining({ code: row.code, severity: profile === "strict" ? row.strict : "warning" }),
                    );
                    expect(result.valid).toBe(profile === "word-valid" || row.strict === "warning");
                }
            });
        }
    for (const profile of ["word-valid", "lenient", "strict"] as const)
        it(`minimal calibration: ${profile}`, async () => {
            const result = await run({ ...MINIMAL_PARTS }, profile);
            expect(result.valid).toBe(true);
            expect(result.issues.filter((i) => i.severity !== "info")).toEqual([]);
        });
});

describe("Word-read markup compatibility regardless of stored name (#1217)", () => {
    const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const defects = [
        {
            name: "MustUnderstand",
            code: "mce-must-understand",
            mutate: (xml: string) => xml.replace("<w:body>", '<w:body xmlns:p="urn:unknown" mc:MustUnderstand="p">'),
        },
        {
            name: "Requires",
            code: "mce-malformed",
            mutate: (xml: string) =>
                xml.replace(
                    "<w:body>",
                    '<w:body><mc:AlternateContent><mc:Choice Requires="missing"><w:p/></mc:Choice><mc:Fallback/></mc:AlternateContent>',
                ),
        },
        {
            name: "character data",
            code: "mce-malformed",
            mutate: (xml: string) =>
                xml.replace(
                    "<w:body>",
                    '<w:body><mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice>Stray<mc:Fallback/></mc:AlternateContent>',
                ),
        },
    ];
    for (const main of ["word/document.bin", "other/main.xml"]) {
        for (const defect of defects) {
            for (const checks of [
                undefined,
                ["word-markup-compatibility"],
                ["word-markup-compatibility", "xml-package-markup-compatibility"],
            ]) {
                it(`${defect.name} blocks ${main}, selection: ${checks?.join(",") ?? "full"}`, async () => {
                    const parts = compatibilityMainParts(main);
                    parts[main] = defect.mutate(parts[main]!);
                    const result = await run(parts, "word-valid", checks);
                    expect(result.issues.filter((issue) => issue.code === defect.code && issue.path === main)).toEqual([
                        expect.objectContaining({
                            severity: "error",
                            ...(defect.name === "character data" ? { message: expect.stringContaining("sessions 65 and 67") } : {}),
                        }),
                    ]);
                    expect(result.valid).toBe(false);
                });
            }
        }
        it(`accepts the renamed main calibration ${main}`, async () => {
            expect((await run(compatibilityMainParts(main), "word-valid")).valid).toBe(true);
        });
    }
    for (const extension of ["xml", "bin"]) {
        for (const used of [false, true]) {
            for (const defect of defects.slice(0, 2)) {
                for (const checks of [undefined, ["word-markup-compatibility", "xml-package-markup-compatibility"]]) {
                    it(`${used ? "used" : "unused"} header.${extension} ${defect.name}, selection: ${checks?.join(",") ?? "full"}`, async () => {
                        const part = `word/header1.${extension}`;
                        const main = defect.mutate(MINIMAL_PARTS["word/document.xml"]!);
                        const body = main.slice(main.indexOf("<w:body"), main.indexOf("</w:body>") + "</w:body>".length);
                        const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="${mc}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${body}</w:hdr>`;
                        const result = await run(compatibilityStoryParts(part, "header", xml, used), "word-valid", checks);
                        expect(result.issues.filter((issue) => issue.code === defect.code && issue.path === part)).toEqual([
                            expect.objectContaining({ severity: used ? "error" : "warning" }),
                        ]);
                        expect(result.valid).toBe(!used);
                    });
                }
            }
        }
    }
});
