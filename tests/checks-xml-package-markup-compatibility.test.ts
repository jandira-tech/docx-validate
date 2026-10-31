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
import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const root = (body: string, attrs = "") => `<x:root xmlns:x="urn:probe" xmlns:mc="${MC}" ${attrs}>${body}</x:root>`;
const run = async (xml: string, part = "customXml/item.xml", profile: "word-valid" | "strict" = "word-valid") =>
    (await createDefaultCheckRegistry()
        .get("xml-package-markup-compatibility")
        ?.run(ValidationContext.fromEntries({ [part]: xml }), profile)) ?? [];

describe("xml-package-markup-compatibility", () => {
    for (const profile of ["word-valid", "strict"] as const) {
        it(`reports out-of-scope Requires under ${profile}`, async () => {
            const issues = await run(
                root(
                    '<mc:AlternateContent><mc:Choice Requires="p"><x:child xmlns:p="urn:p"/></mc:Choice><mc:Fallback/></mc:AlternateContent>',
                ),
                undefined,
                profile,
            );
            expect(issues).toEqual([
                expect.objectContaining({
                    code: "mce-malformed",
                    severity: profile === "strict" ? "error" : "warning",
                    path: "customXml/item.xml",
                }),
            ]);
        });
        it(`reports descendant Ignorable with sibling-only declaration under ${profile}`, async () => {
            expect(await run(root('<x:a mc:Ignorable="p"/><x:b xmlns:p="urn:p"/>'), undefined, profile)).toEqual([
                expect.objectContaining({ code: "mce-unbound-prefix", severity: profile === "strict" ? "error" : "warning" }),
            ]);
        });
    }
    it("accepts an ancestor declaration", async () => {
        expect(await run(root('<x:a mc:Ignorable="p"/>', 'xmlns:p="urn:p"'))).toEqual([]);
    });
    it("does not treat a descendant declaration as a root binding", async () => {
        expect(await run(root('<x:a xmlns:p="urn:p"/>', 'mc:Ignorable="p"'))).toEqual([
            expect.objectContaining({ code: "mce-unbound-prefix", severity: "warning" }),
        ]);
    });
    it("ignores contents of an unselected Choice", async () => {
        expect(
            await run(
                root(
                    '<mc:AlternateContent><mc:Choice xmlns:p="urn:p" Requires="p"><x:a mc:Ignorable="missing"/></mc:Choice><mc:Fallback/></mc:AlternateContent>',
                ),
            ),
        ).toEqual([]);
    });
    it("does not duplicate the Word or content-types checks", async () => {
        for (const part of ["word/document.xml", "_rels/.rels", "[Content_Types].xml"]) {
            expect(await run(root('<x:a mc:Ignorable="p"/>'), part)).toEqual([]);
        }
    });
    it("does not parse an untyped binary part or recovered XML", async () => {
        expect(await run(root('<x:a mc:Ignorable="p"/>'), "customXml/item.bin")).toEqual([]);
        expect(await run("<broken")).toEqual([]);
    });
    for (const profile of ["word-valid", "strict"] as const) {
        for (const main of ["word/document.bin", "other/main.xml"]) {
            it(`leaves the related ${main} to the Word check under ${profile} (#1217)`, async () => {
                const parts = { ...MINIMAL_PARTS };
                delete parts["word/document.xml"];
                parts["_rels/.rels"] = parts["_rels/.rels"]!.replace("word/document.xml", main);
                parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace("/word/document.xml", `/${main}`);
                parts[main] = root('<x:a mc:Ignorable="missing"/>');
                const issues = await createDefaultCheckRegistry()
                    .get("xml-package-markup-compatibility")
                    ?.run(ValidationContext.fromEntries(parts), profile);
                expect(issues).toEqual([]);
                expect(
                    await createDefaultCheckRegistry().get("word-markup-compatibility")?.run(ValidationContext.fromEntries(parts), profile),
                ).toEqual([expect.objectContaining({ path: main, severity: "error", code: "mce-unbound-prefix" })]);
            });
        }
        for (const [part, segment] of [
            ["other/header.bin", "header"],
            ["other/settings.xml", "settings"],
        ] as const) {
            it(`assigns a related ${segment} outside word/ to its compatibility check under ${profile} (#1217)`, async () => {
                const parts = { ...MINIMAL_PARTS };
                parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${segment}" Target="../${part}"/></Relationships>`,
                );
                parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                    "<w:sectPr>",
                    segment === "header"
                        ? '<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="header" w:type="default"/>'
                        : "<w:sectPr>",
                );
                parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
                    "</Types>",
                    `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${segment}+xml"/></Types>`,
                );
                parts[part] = root("", 'xmlns:p="urn:unknown" mc:MustUnderstand="p"');
                const ctx = ValidationContext.fromEntries(parts);
                const packageIssues = await createDefaultCheckRegistry().get("xml-package-markup-compatibility")?.run(ctx, profile);
                const wordIssues = await createDefaultCheckRegistry().get("word-markup-compatibility")?.run(ctx, profile);
                if (segment === "header") {
                    expect(packageIssues).toEqual([]);
                    expect(wordIssues).toEqual([expect.objectContaining({ path: part, severity: "error" })]);
                } else {
                    // Settings are not story parts; the package check retains them.
                    expect(packageIssues).toEqual([
                        expect.objectContaining({ path: part, severity: profile === "strict" ? "error" : "warning" }),
                    ]);
                    expect(wordIssues).toEqual([]);
                }
            });
        }
    }
});
