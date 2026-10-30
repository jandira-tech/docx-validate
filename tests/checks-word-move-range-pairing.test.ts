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
const start = (kind: string, name: string, id: string) =>
    `<w:${kind}RangeStart w:id="${id}" w:name="${name}" w:author="Probe" w:date="2026-10-30T00:00:00Z"/>`;
const end = (kind: string, id: string) => `<w:${kind}RangeEnd w:id="${id}"/>`;
const range = (kind: string, name: string, id: string) => start(kind, name, id) + end(kind, id);
const run = async (markers: string, profile: "word-valid" | "strict" = "word-valid", ns = W) => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": `<w:document xmlns:w="${ns}" xmlns:mc="${MC}"><w:body><w:p>${markers}</w:p></w:body></w:document>`,
    });
    return (await createDefaultCheckRegistry().get("word-move-range-unpaired")?.run(ctx, profile)) ?? [];
};

describe("word-move-range-unpaired", () => {
    for (const profile of ["word-valid", "strict"] as const) {
        for (const kind of ["moveFrom", "moveTo"])
            it(`reports ${kind} without counterpart under ${profile}`, async () => {
                expect(await run(range(kind, "moveA", "10"), profile)).toEqual([
                    expect.objectContaining({
                        code: "word-move-range-unpaired",
                        severity: profile === "strict" ? "error" : "warning",
                        path: "word/document.xml",
                    }),
                ]);
            });
    }
    it("reports mismatched names", async () => {
        expect(await run(range("moveFrom", "a", "10") + range("moveTo", "b", "11"))).toHaveLength(2);
    });
    it("reports missing and orphan ends independently of name pairing", async () => {
        expect(await run(start("moveFrom", "a", "10") + range("moveTo", "a", "11"))).toHaveLength(1);
        expect(await run(end("moveFrom", "10"))).toHaveLength(1);
        expect(await run(end("moveFrom", "10") + start("moveFrom", "a", "10") + range("moveTo", "a", "11"))).toHaveLength(2);
    });
    it("does not match an end of the other range kind", async () => {
        expect(await run(start("moveFrom", "a", "10") + end("moveTo", "10"))).toHaveLength(3);
    });
    it("does not balance across separate textboxes", async () => {
        expect(
            await run(
                `<w:txbxContent><w:p>${range("moveFrom", "a", "10")}</w:p></w:txbxContent><w:txbxContent><w:p>${range("moveTo", "a", "11")}</w:p></w:txbxContent>`,
            ),
        ).toHaveLength(2);
    });
    it("accepts valid pairs and exact large decimal ids", async () => {
        expect(await run(range("moveFrom", "a", "9007199254740993") + range("moveTo", "a", "9007199254740994"))).toEqual([]);
        expect(await run(start("moveFrom", "a", "+010") + end("moveFrom", "10") + range("moveTo", "a", "11"))).toEqual([]);
    });
    it("supports ISO Strict namespaces", async () => {
        expect(await run(range("moveFrom", "a", "10"), "strict", WS)).toEqual([
            expect.objectContaining({ code: "word-move-range-unpaired" }),
        ]);
    });
    it("does not require ranges to nest", async () => {
        expect(await run(start("moveFrom", "a", "10") + start("moveTo", "a", "11") + end("moveFrom", "10") + end("moveTo", "11"))).toEqual(
            [],
        );
    });
    it("reports unequal counterpart counts", async () => {
        expect(await run(range("moveFrom", "a", "10") + range("moveFrom", "a", "12") + range("moveTo", "a", "11"))).toHaveLength(1);
    });
    it("ignores an unselected branch", async () => {
        expect(
            await run(
                `<mc:AlternateContent><mc:Choice xmlns:p="urn:p" Requires="p">${range("moveFrom", "a", "10")}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
            ),
        ).toEqual([]);
    });
    it("leaves missing id/name attributes to their own rules", async () => {
        expect(await run("<w:moveFromRangeStart/><w:moveFromRangeEnd/>")).toEqual([]);
    });
});
