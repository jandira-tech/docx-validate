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

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const field = (...types: string[]) => types.map((type) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`).join("");
const run = async (body: string, profile: "word-valid" | "strict" = "word-valid", ns = W, part = "word/document.xml") => {
    const ctx = ValidationContext.fromEntries({
        [part]: `<w:document xmlns:w="${ns}" xmlns:mc="${MC}"><w:body>${body}</w:body></w:document>`,
    });
    return (await createDefaultCheckRegistry().get("word-field-unbalanced")?.run(ctx, profile)) ?? [];
};

describe("word-field-unbalanced", () => {
    for (const profile of ["word-valid", "strict"] as const) {
        for (const types of [["begin"], ["end"], ["separate"], ["begin", "separate", "separate", "end"]]) {
            it(`reports ${types.join("/")} under ${profile}`, async () => {
                expect(await run(`<w:p>${field(...types)}</w:p>`, profile)).toEqual([
                    expect.objectContaining({
                        code: "word-field-unbalanced",
                        severity: profile === "strict" ? "error" : "warning",
                        path: "word/document.xml",
                    }),
                ]);
            });
        }
    }
    it("accepts nested fields, optional separators and a field spanning paragraphs", async () => {
        expect(await run(`<w:p>${field("begin", "separate", "begin", "end")}</w:p><w:p>${field("end")}</w:p>`)).toEqual([]);
        expect(await run(`<w:p>${field("begin", "end")}</w:p>`)).toEqual([]);
    });
    it("supports ISO Strict namespaces", async () => {
        expect(await run(`<w:p>${field("begin")}</w:p>`, "strict", WS)).toEqual([
            expect.objectContaining({ code: "word-field-unbalanced" }),
        ]);
    });
    it("does not pair across textboxes or notes", async () => {
        expect(await run(`<w:p>${field("begin")}<w:txbxContent><w:p>${field("end")}</w:p></w:txbxContent></w:p>`)).toHaveLength(2);
        expect(
            await run(
                `<w:footnote w:id="1"><w:p>${field("begin")}</w:p></w:footnote><w:footnote w:id="2"><w:p>${field("end")}</w:p></w:footnote>`,
            ),
        ).toHaveLength(2);
    });
    it("uses the selected compatibility branch", async () => {
        expect(
            await run(
                `<w:p><mc:AlternateContent><mc:Choice xmlns:p="urn:p" Requires="p">${field("begin")}</mc:Choice><mc:Fallback>${field("begin", "end")}</mc:Fallback></mc:AlternateContent></w:p>`,
            ),
        ).toEqual([]);
    });
    it("does not scan unrelated parts or missing field types", async () => {
        expect(await run(`<w:p>${field("begin")}</w:p>`, "strict", W, "customXml/item.xml")).toEqual([]);
        expect(await run("<w:p><w:r><w:fldChar/></w:r></w:p>")).toEqual([]);
    });
    it("skips malformed XML", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": "<broken" });
        expect((await createDefaultCheckRegistry().get("word-field-unbalanced")?.run(ctx, "strict")) ?? []).toEqual([]);
    });
});
