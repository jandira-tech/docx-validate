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

// #665: a relationship with TargetMode="External" names a resource outside
// the package, never a part (ECMA-376 Part 2 §9.3), so the file-reference
// walk does not resolve its target as a package path under strict and
// word-valid. Lenient keeps its historical scan. Word 16.115 opened the
// External `.rels`, `/_rels/` and `../customXml/` targets word-valid used to
// block clean, as direct children and in a selected w14 Choice (session 40,
// s40a/s40b, docs/plans/evidence/word-probes-2026-10-24.md).

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { DOCUMENT_RELS, MCE_DECLARATIONS } from "./support/relationships-markup-compatibility-specimens";

const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/** The minimal package plus one hyperlink relationship to `target`, a direct child or inside a selected w14 Choice. */
const withHyperlink = (target: string, placement: "direct" | "choice", targetMode: string | null = "External"): Record<string, string> => {
    const mode = targetMode === null ? "" : ` TargetMode="${targetMode}"`;
    const relationship = `<Relationship Id="rId9" Type="${HYPERLINK}" Target="${target}"${mode}/>`;
    const content =
        placement === "direct"
            ? relationship
            : `<mc:AlternateContent><mc:Choice Requires="w14">${relationship}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
    const rels = MINIMAL_PARTS[DOCUMENT_RELS]!;
    return {
        ...MINIMAL_PARTS,
        [DOCUMENT_RELS]: (placement === "direct" ? rels : rels.replace("<Relationships ", `<Relationships ${MCE_DECLARATIONS} `)).replace(
            "</Relationships>",
            `${content}</Relationships>`,
        ),
    };
};

const run = async (parts: Record<string, string>, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        await fs.writeFile(file, await buildDocx(parts));
        const result = await validate(file, { profile });
        const errors = [...new Set(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code))];
        return { valid: result.valid, errors, broken: result.issues.filter((issue) => issue.code === "rels-broken") };
    });

const EXTERNAL_TARGETS = [
    "file:///C:/missing.pdf",
    "file:///missing.rels",
    "file:///C:/docs/_rels/missing.xml",
    "../customXml/missing.xml",
    "\\\\server\\share\\missing.docx",
    "ftp://files.invalid/missing.pdf",
];

describe("TargetMode=External targets are not package paths (#665)", () => {
    describe.each(["direct", "choice"] as const)("as a %s relationship", (placement) => {
        it.each(EXTERNAL_TARGETS)("%s is no broken reference under strict or word-valid", async (target) => {
            for (const profile of ["strict", "word-valid"] as const) {
                const { valid, errors, broken } = await run(withHyperlink(target, placement), profile);
                expect(broken, profile).toEqual([]);
                // Stored OPC relationships allow no AlternateContent child, so
                // strict refuses the wrapper on its own rule, whatever the target (#632).
                const expected = placement === "choice" && profile === "strict" ? ["opc-relationships-markup-compatibility"] : [];
                expect(errors, profile).toEqual(expected);
                expect(valid, profile).toBe(expected.length === 0);
            }
        });
    });

    it("keeps an External .rels target a w:hyperlink names valid under strict and word-valid", async () => {
        const parts = withHyperlink("file:///missing.rels", "direct");
        parts["word/document.xml"] = parts["word/document.xml"]!.replace(
            "<w:document ",
            '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        ).replace("<w:r><w:t>Hello</w:t></w:r>", '<w:hyperlink r:id="rId9"><w:r><w:t>Hello</w:t></w:r></w:hyperlink>');
        expect(parts["word/document.xml"]).toContain('<w:hyperlink r:id="rId9">');
        for (const profile of ["strict", "word-valid"] as const) {
            expect(await run(parts, profile), profile).toEqual({ valid: true, errors: [], broken: [] });
        }
    });

    it.each(EXTERNAL_TARGETS)("keeps lenient's historical rels-broken for %s", async (target) => {
        const { broken } = await run(withHyperlink(target, "direct"), "lenient");
        expect(broken.map((issue) => issue.message)).toEqual([`Broken reference to ${target}`]);
    });

    it.each([
        ["absent", null],
        ["Internal", "Internal"],
    ] as const)("still resolves a target whose TargetMode is %s as a part, under every profile", async (_label, targetMode) => {
        for (const profile of ["strict", "word-valid", "lenient"] as const) {
            const { valid, broken } = await run(withHyperlink("missing.pdf", "direct", targetMode), profile);
            expect(
                broken.map((issue) => issue.message),
                profile,
            ).toEqual(["Broken reference to missing.pdf"]);
            expect(valid, profile).toBe(profile === "word-valid");
        }
    });
});
