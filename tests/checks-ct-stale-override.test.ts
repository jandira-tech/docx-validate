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

const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const run = async (xml: string | undefined, parts: Record<string, string> = {}, profile: "word-valid" | "strict" = "word-valid") => {
    const ctx = ValidationContext.fromEntries({ ...(xml === undefined ? {} : { "[Content_Types].xml": xml }), ...parts });
    return (await createDefaultCheckRegistry().get("ct-stale-override")?.run(ctx, profile)) ?? [];
};
const types = (body: string) => `<Types xmlns="${CT}" xmlns:mc="${MC}">${body}</Types>`;
const override = (name: string) => `<Override PartName="${name}" ContentType="application/xml"/>`;

describe("ct-stale-override", () => {
    for (const profile of ["word-valid", "strict"] as const)
        it(`reports the absent target under ${profile}`, async () => {
            expect(await run(types(override("/word/absent.xml")), {}, profile)).toEqual([
                expect.objectContaining({
                    code: "ct-stale-override",
                    severity: "warning",
                    path: "[Content_Types].xml",
                }),
            ]);
        });
    it("names each stale declaration", async () => {
        const result = await run(types(override("/a.xml") + override("/b.xml")));
        expect(result).toHaveLength(2);
        expect(result.map((i) => i.message).join(" ")).toContain("/a.xml");
        expect(result.map((i) => i.message).join(" ")).toContain("/b.xml");
    });
    it("uses OPC equivalence, including unreserved escapes", async () => {
        expect(await run(types(override("/WORD/%73tyles.xml")), { "word/styles.xml": "<styles/>" })).toEqual([]);
    });
    for (const profile of ["word-valid", "strict"] as const) {
        for (const stream of ["[Content_Types].xml", "[content_types].XML"]) {
            it(`does not count ${stream} as an Override target under ${profile}`, async () => {
                const ctx = ValidationContext.fromEntries({ [stream]: types(override("/%5BContent_Types%5D.xml")) });
                expect(await createDefaultCheckRegistry().get("ct-stale-override")?.run(ctx, profile)).toEqual([
                    expect.objectContaining({ code: "ct-stale-override", severity: "warning", path: stream }),
                ]);
                expect(await run(types(override("/%5BContent_Types%5D.xml")), { "%5BContent_Types%5D.xml": "<part/>" }, profile)).toEqual(
                    [],
                );
            });
        }
    }
    it("looks up many Overrides without rescanning the part names for each", async () => {
        const parts = Object.fromEntries(Array.from({ length: 256 }, (_, i) => [`customXml/item${i}.xml`, "<part/>"]));
        const declarations = Object.keys(parts).map((name) => override(`/${name.toUpperCase()}`));
        const ctx = ValidationContext.fromEntries({
            "[Content_Types].xml": types(declarations.join("") + override("/absent.xml")),
            ...parts,
        });
        let nameReads = 0;
        // Count work on the input rather than relying on wall-clock timing.
        ctx.partNames.forEach((name, index) => {
            Object.defineProperty(ctx.partNames, index, {
                get: () => {
                    nameReads += 1;
                    return name;
                },
            });
        });
        const result = await createDefaultCheckRegistry().get("ct-stale-override")?.run(ctx, "word-valid");
        expect(result).toEqual([
            expect.objectContaining({ code: "ct-stale-override", severity: "warning", message: expect.stringContaining("/absent.xml") }),
        ]);
        expect(nameReads).toBeLessThan(ctx.partNames.length * 8);
    });
    it("does not replace grammar checks", async () => {
        expect(await run(types('<Override ContentType="application/xml"/>' + override("") + override("/word//bad.xml")))).toEqual([]);
    });
    it("ignores declarations in an unselected branch", async () => {
        expect(
            await run(
                types(
                    `<mc:AlternateContent><mc:Choice xmlns:p="urn:unknown" Requires="p">${override("/absent.xml")}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
                ),
            ),
        ).toEqual([]);
    });
    it("skips missing and malformed streams", async () => {
        expect(await run(undefined)).toEqual([]);
        expect(await run("<broken")).toEqual([]);
    });
});
