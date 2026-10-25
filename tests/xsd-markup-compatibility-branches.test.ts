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

// #627: the schema walk validates the AlternateContent branch markup
// compatibility selects (ISO/IEC 29500-3 §10.2), not nothing and not every
// branch. The selection follows the schema view's understood namespaces: the
// first mc:Choice whose Requires prefixes are all understood, else mc:Fallback.

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** Validate a minimal package whose main document has `rootAttrs` and `body`, under `profile`. */
const schemaWalk = async (profile: Profile, rootAttrs: string, body: string) =>
    withTempDir(async (dir) => {
        const parts: Record<string, string> = {
            ...MINIMAL_PARTS,
            "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${rootAttrs}><w:body>${body}<w:sectPr/></w:body></w:document>`,
        };
        for (const [part, xml] of Object.entries(parts)) {
            await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
            await fs.writeFile(path.join(dir, part), xml, "utf-8");
        }
        const result = await new DOCXSchemaValidator({ unpackedDir: dir, profile }).validateAgainstXsd();
        return result.issues.filter((issue) => issue.path === "word/document.xml" && issue.severity === "error");
    });

describe.each(["strict", "lenient"] as const)(
    "the schema walk validates the selected AlternateContent branch (#627) under %s",
    (profile) => {
        it("validates an understood mc:Choice payload: a bogus element in it is a schema error", async () => {
            const errors = await schemaWalk(
                profile,
                `${W} ${MC} ${M}`,
                `<w:p><mc:AlternateContent><mc:Choice Requires="m"><w:bogusChoicePayload/></mc:Choice>` +
                    `<mc:Fallback><w:r><w:t>legacy</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
            );
            expect(errors.some((issue) => issue.message.includes("bogusChoicePayload"))).toBe(true);
        });

        it("validates mc:Fallback when no mc:Choice exists", async () => {
            const errors = await schemaWalk(
                profile,
                `${W} ${MC}`,
                `<w:p><mc:AlternateContent><mc:Fallback><w:bogusFallbackPayload/></mc:Fallback></mc:AlternateContent></w:p>`,
            );
            expect(errors.some((issue) => issue.message.includes("bogusFallbackPayload"))).toBe(true);
        });

        it("validates mc:Fallback, not a Choice whose Requires the schema view does not understand", async () => {
            // The draft patch on #627 always took the first Choice. ISO/IEC 29500-3
            // §10.2 selects a Choice only when every namespace it requires is
            // understood; the bundled schemas do not understand w14.
            const errors = await schemaWalk(
                profile,
                `${W} ${MC} ${W14}`,
                `<w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:bogusUnselectedChoice/></mc:Choice>` +
                    `<mc:Fallback><w:bogusSelectedFallback/></mc:Fallback></mc:AlternateContent></w:p>`,
            );
            expect(errors.some((issue) => issue.message.includes("bogusSelectedFallback"))).toBe(true);
            expect(errors.some((issue) => issue.message.includes("bogusUnselectedChoice"))).toBe(false);
        });

        it("does not validate the unselected mc:Fallback of a selected Choice", async () => {
            const errors = await schemaWalk(
                profile,
                `${W} ${MC} ${M}`,
                `<w:p><mc:AlternateContent><mc:Choice Requires="m"><w:r><w:t>modern</w:t></w:r></mc:Choice>` +
                    `<mc:Fallback><w:bogusNeverValidated/></mc:Fallback></mc:AlternateContent></w:p>`,
            );
            expect(errors).toEqual([]);
        });

        it("drops attributes in a namespace mc:Ignorable lists", async () => {
            const errors = await schemaWalk(
                profile,
                `${W} ${MC} ${W14} mc:Ignorable="w14"`,
                `<w:p w14:paraId="2313C800" w14:textId="6454BBE0"><w:r><w:t>ok</w:t></w:r></w:p>`,
            );
            expect(errors).toEqual([]);
        });

        it("selects nothing, rather than an un-validatable wps payload, for a Choice requiring wps with no Fallback", async () => {
            const errors = await schemaWalk(
                profile,
                `${W} ${MC} ${WP} ${A} ${WPS}`,
                `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>` +
                    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="Box"/>` +
                    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
                    `<wps:wsp><wps:spPr/></wps:wsp></a:graphicData></a:graphic></wp:inline>` +
                    `</w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>`,
            );
            expect(errors).toEqual([]);
        });
    },
);
