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

// #653: repairMissingParaIds stamps w14:paraId/w14:textId. When the part's
// root neither declares w14 nor lists it in mc:Ignorable, the stamps used to
// leave w14 non-ignorable, so the repaired part gained
// mce-extension-not-ignorable. The repair now declares w14 (and mc) on the
// root and lists w14 in mc:Ignorable, as Word writes it. Word 16.115 opens
// the repaired part clean (session 40, docs/plans/evidence/word-probes-2026-10-24.md).

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { NS } from "../src/lib/types";
import { parseXml } from "../src/lib/xml-helpers";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

const writeParts = async (dir: string, parts: Readonly<Record<string, string>>): Promise<void> => {
    for (const [part, xml] of Object.entries(parts)) {
        await fs.mkdir(path.dirname(path.join(dir, part)), { recursive: true });
        await fs.writeFile(path.join(dir, part), xml, "utf-8");
    }
};

/** The minimal package with a main document whose root carries `rootAttributes` and whose paragraphs have no ids. */
const withoutParaIds = (rootAttributes: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<w:document xmlns:w="${W}"${rootAttributes}><w:body>` +
        `<w:p><w:r><w:t>one</w:t></w:r></w:p>` +
        `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
});

/** Repair the main document of `parts` and return its repaired root and text. */
const repaired = async (parts: Readonly<Record<string, string>>) =>
    withTempDir(async (dir) => {
        await writeParts(dir, parts);
        const repairs = await new DOCXSchemaValidator({ unpackedDir: dir }).repairMissingParaIds();
        const text = await fs.readFile(path.join(dir, "word", "document.xml"), "utf-8");
        return { repairs, text, root: parseXml(text).documentElement! };
    });

describe("repairMissingParaIds keeps w14 ignorable (#653)", () => {
    it("declares w14 and mc on a root that has neither and lists w14 in mc:Ignorable", async () => {
        const { repairs, text, root } = await repaired(withoutParaIds(""));
        expect(repairs).toBe(6);
        expect(root.getAttribute("xmlns:w14")).toBe(W14);
        expect(root.getAttribute("xmlns:mc")).toBe(NS.MC);
        expect(root.getAttributeNS(NS.MC, "Ignorable")).toBe("w14");
        // One declaration, on the root, rather than one on each stamped element.
        expect(text.match(/xmlns:w14=/g)).toHaveLength(1);
    });

    it("appends w14 to an existing mc:Ignorable list, keeping its tokens", async () => {
        const { root } = await repaired(withoutParaIds(` xmlns:w15="${W15}" xmlns:mc="${NS.MC}" mc:Ignorable="w15"`));
        expect(root.getAttributeNS(NS.MC, "Ignorable")).toBe("w15 w14");
        expect(root.getAttribute("xmlns:w14")).toBe(W14);
    });

    it("lists a w14 the root declares but does not list", async () => {
        const { root } = await repaired(withoutParaIds(` xmlns:w14="${W14}"`));
        expect(root.getAttribute("xmlns:mc")).toBe(NS.MC);
        expect(root.getAttributeNS(NS.MC, "Ignorable")).toBe("w14");
    });

    it("reuses the prefix the root already binds to the compatibility namespace", async () => {
        const { root } = await repaired(withoutParaIds(` xmlns:compat="${NS.MC}" compat:Ignorable="w15" xmlns:w15="${W15}"`));
        expect(root.getAttributeNS(NS.MC, "Ignorable")).toBe("w15 w14");
        expect(root.hasAttribute("xmlns:mc")).toBe(false);
    });

    it.each([
        ["w14", ` xmlns:w14="urn:other"`],
        ["mc", ` xmlns:mc="urn:other"`],
    ])("stamps but leaves the root as it is when the root binds %s to another namespace", async (prefix, attributes) => {
        const { repairs, root } = await repaired(withoutParaIds(attributes));
        expect(repairs).toBe(6);
        expect(root.getAttribute(`xmlns:${prefix}`)).toBe("urn:other");
        expect(root.hasAttributeNS(NS.MC, "Ignorable")).toBe(false);
        expect(root.hasAttribute(prefix === "w14" ? "xmlns:mc" : "xmlns:w14")).toBe(false);
    });

    it("changes nothing on the root of a part it does not stamp", async () => {
        const parts = withoutParaIds("");
        const { repairs, root } = await repaired({
            ...parts,
            "word/document.xml": parts["word/document.xml"]!.replace(/<w:p>.*<\/w:tbl>/, ""),
        });
        expect(repairs).toBe(0);
        expect(root.hasAttribute("xmlns:w14")).toBe(false);
        expect(root.hasAttribute("xmlns:mc")).toBe(false);
    });

    it.each(["lenient", "strict", "word-valid"] as const)("leaves no new finding after a full %s auto-repair", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(withoutParaIds("")));
            const before = await validate(file, { profile });
            expect(before.issues.filter((issue) => issue.severity !== "info").map((issue) => issue.code)).toEqual([
                "paraid-missing-element",
                "paraid-missing-element",
            ]);
            const result = await validate(file, { profile, autoRepair: true });
            expect(result.issues.filter((issue) => issue.severity !== "info")).toEqual([]);
            expect(result.issues.filter((issue) => issue.code === "repair-plan").map((issue) => issue.message)).toEqual([
                "Before repair found 2 [paraid-missing-element] issue(s). Auto-repair cleared 2; 0 remain.",
            ]);
            expect(result.valid).toBe(true);
        });
    });
});
