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

// #620: Word reads a part by its relationship and content type, not its name.
// A main document stored as word/document.bin under the main document's
// Override opens as one, and Word refuses it truncated and offers recovery
// for its footer reference without a sidecar (session 32, s32a, in
// docs/plans/evidence/word-probes-2026-10-24.md). The XML passes therefore
// judge every part the content-types stream declares as XML, whatever its
// name, and such a main document is judged as its `.xml` twin is.

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { isXmlContentType } from "../src/scripts/office/checks/xml";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "[Content_Types].xml";
const MAIN = "word/document.xml";
const RELS = "word/_rels/document.xml.rels";
type Parts = Record<string, string>;

const replaceOnce = (text: string, from: string, to: string): string => {
    if (!text.includes(from)) throw new Error(`missing ${from}`);
    return text.replace(from, to);
};

/** Session 32's rename: the main document, its sidecar, its Override and its package relationship move to word/document.bin. */
const toBin = (parts: Parts): Parts => {
    const out: Parts = { ...parts };
    out["word/document.bin"] = out[MAIN]!;
    delete out[MAIN];
    if (out[RELS] !== undefined) {
        out["word/_rels/document.bin.rels"] = out[RELS]!;
        delete out[RELS];
    }
    out[CT] = replaceOnce(out[CT]!, 'PartName="/word/document.xml"', 'PartName="/word/document.bin"');
    out["_rels/.rels"] = replaceOnce(out["_rels/.rels"]!, 'Target="word/document.xml"', 'Target="word/document.bin"');
    return out;
};

const truncated = (): Parts => ({ ...MINIMAL_PARTS, [MAIN]: replaceOnce(MINIMAL_PARTS[MAIN]!, "</w:body></w:document>", "") });

/** The main document's sectPr references footer rId9; without `sidecar` its `.rels` is removed, otherwise it has no rId9. */
const footerReference = (sidecar: boolean): Parts => {
    const parts: Parts = { ...MINIMAL_PARTS };
    parts[MAIN] = replaceOnce(
        replaceOnce(parts[MAIN]!, "<w:document ", `<w:document xmlns:r="${R}" `),
        "<w:sectPr>",
        '<w:sectPr><w:footerReference w:type="default" r:id="rId9"/>',
    );
    if (!sidecar) delete parts[RELS];
    return parts;
};

/** The distinct `code@path` of each error a full run under `profile` reports. */
const errorsOf = async (parts: Parts, profile: Profile): Promise<string[]> =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        await fs.writeFile(file, await buildDocx(parts));
        const result = await validate(file, { profile });
        const errors = result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? "-"}`);
        expect(result.valid).toBe(errors.length === 0);
        return [...new Set(errors)].sort();
    });

const PROFILES = ["word-valid", "strict", "lenient"] as const;

describe("a main document stored without a .xml suffix is judged as its .xml twin (#620)", () => {
    it.each([
        // Word 16.115, session 32: open-error / open-error.
        ["main-truncated", truncated, ["xml-syntax@word/document.bin"]],
        // Word 16.115, session 32: repair-warning / repair-warning.
        ["footer reference without a sidecar", () => footerReference(false), ["rels-missing-sidecar@word/document.bin"]],
        // Word 16.115, session 32: repair-warning / repair-warning.
        ["footer reference to a missing Id", () => footerReference(true), []],
    ] as const)("%s", async (_name, make, wordValidErrors) => {
        for (const profile of PROFILES) {
            const twin = await errorsOf(make(), profile);
            const renamed = await errorsOf(toBin(make()), profile);
            expect(twin.length, profile).toBeGreaterThan(0);
            expect(renamed, profile).toEqual(twin.map((error) => error.replace("word/document.xml", "word/document.bin")).sort());
            if (profile === "word-valid" && wordValidErrors.length > 0) expect(renamed).toEqual(wordValidErrors);
        }
    });

    it.each(PROFILES)("keeps the clean renamed main document valid under %s", async (profile) => {
        expect(await errorsOf(toBin({ ...MINIMAL_PARTS }), profile)).toEqual([]);
    });
});

describe("the XML passes follow the content type, not the name (#620)", () => {
    it.each([
        ["application/xml", true],
        ["text/xml", true],
        ["application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", true],
        ["Application/XML; charset=utf-8", true],
        ["application/vnd.openxmlformats-officedocument.oleObject", false],
        ["application/vnd.ms-office.vbaProject", false],
        ["image/png", false],
        // Word renders an SVG through its image pipeline; malformed markup there is unobserved.
        ["image/svg+xml", false],
        ["", false],
    ] as const)("isXmlContentType(%j) is %s", (contentType, expected) => {
        expect(isXmlContentType(contentType)).toBe(expected);
    });

    it("reports a malformed part an XML Default declares, and blocks only where Word reads it", async () => {
        const parts: Parts = {
            ...MINIMAL_PARTS,
            [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", '<Default Extension="dat" ContentType="application/xml"/></Types>'),
            "word/extra.dat": "<unclosed>",
        };
        // Unreferenced: Word does not read it (#437), so word-valid only warns.
        expect(await errorsOf(parts, "word-valid")).toEqual([]);
        expect(await errorsOf(parts, "strict")).toContain("xml-syntax@word/extra.dat");
        expect(await errorsOf(parts, "lenient")).toContain("xml-syntax@word/extra.dat");
    });

    it("leaves a part with a non-XML type unparsed, whatever its bytes", async () => {
        const parts: Parts = {
            ...MINIMAL_PARTS,
            [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", '<Default Extension="svg" ContentType="image/svg+xml"/></Types>'),
            "word/media/image1.svg": "<svg",
        };
        for (const profile of PROFILES) {
            expect(
                (await errorsOf(parts, profile)).filter((error) => error.startsWith("xml-")),
                profile,
            ).toEqual([]);
        }
    });

    it("runs the xml-syntax check of a partial run over the renamed main document", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(toBin(truncated())));
            const result = await validate(file, { checks: ["xml-syntax"] });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`)).toContain(
                "xml-syntax@word/document.bin",
            );
        });
    });
});
