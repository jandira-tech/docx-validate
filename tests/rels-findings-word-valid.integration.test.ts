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

// #734: the word-valid verdict of a `rels-broken` finding reads the part the
// Target resolves to, never the raw Target spelling its message embeds. A
// target spelled `media/_rels/x.png`, `notes.rels` or `word/../customXml/x.png`
// names none of the structural parts Word was shown, so it keeps the tolerated
// broken-media warning whatever its spelling (Word opened all three clean,
// session 54); `../customXml/item1.xml` (vectors.docx, unreadable content) and
// a nested `.rels` sidecar stay blocking. No registered check emits `rels-broken` — the DOCX validator's own
// file-reference walk does — so the full run is the profile that must agree.

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const IMAGE = `${RELATIONSHIPS}/image`;
const CUSTOM_XML = `${RELATIONSHIPS}/customXml`;
const SIDECAR = "word/_rels/document.xml.rels";

/** The minimal package plus one internal relationship of `type` whose `target` names no stored part. */
const withBrokenTarget = (target: string, type = IMAGE): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace(
        "</Relationships>",
        `<Relationship Id="rId9" Type="${type}" Target="${target}"/></Relationships>`,
    ),
});

const GLOSSARY =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>`;

/** A reached glossary whose own relationships part relates a `customXml` part to an absent `target`. */
const withBrokenGlossaryTarget = (target: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml": GLOSSARY,
    "word/glossary/_rels/document.xml.rels":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${CUSTOM_XML}" Target="${target}"/></Relationships>`,
    [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace(
        "</Relationships>",
        `<Relationship Id="rIdG" Type="${RELATIONSHIPS}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>`,
    ),
});

/** The minimal package plus the given `<Relationship>` element in the main document's sidecar. */
const withRelationship = (relationship: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace("</Relationships>", `${relationship}</Relationships>`),
});

const run = async (parts: Record<string, string>, profile: Profile) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        await fs.writeFile(file, await buildDocx(parts));
        const result = await validate(file, { profile });
        return { valid: result.valid, broken: result.issues.filter((issue) => issue.code === "rels-broken") };
    });

describe("the rels-broken word-valid verdict reads the resolved part name (#734)", () => {
    it.each([
        // Spellings whose text carries `../customXml/`, `/_rels/` or `.rels`
        // while the resolved part is none of those: `word/media/_rels/x.png`,
        // `word/notes.rels` and `word/customXml/x.png`.
        ["media/_rels/x.png"],
        ["notes.rels"],
        ["word/../customXml/x.png"],
    ])("keeps %s a word-valid warning the full run tolerates", async (target) => {
        const { valid, broken } = await run(withBrokenTarget(target), "word-valid");
        expect(broken.map((issue) => issue.severity)).toEqual(["warning"]);
        expect(valid).toBe(true);
    });

    it.each([
        // vectors.docx's shape: a broken relationship to the package-root
        // customXml folder (Word offered recovery, hash-bound observation).
        ["../customXml/missing.xml"],
        // A broken target naming another relationships part.
        ["../customXml/_rels/item1.xml.rels"],
    ])("keeps %s word-blocking", async (target) => {
        const { valid, broken } = await run(withBrokenTarget(target), "word-valid");
        expect(broken.map((issue) => issue.severity)).toEqual(["error"]);
        expect(valid).toBe(false);
    });

    it.each([["media/_rels/x.png"], ["notes.rels"], ["../customXml/missing.xml"], ["../customXml/_rels/item1.xml.rels"]])(
        "keeps %s an error under strict, whatever its word-valid verdict",
        async (target) => {
            const { valid, broken } = await run(withBrokenTarget(target), "strict");
            expect(broken.map((issue) => issue.severity)).toEqual(["error"]);
            expect(valid).toBe(false);
        },
    );
});

// Session 54 (docs/plans/evidence/word-probes-2026-10-24.md, s54a/s54b): each
// row ran twice and was read twice. A customXml relationship to an absent
// package-root `customXml/item1.xml` is refused however the Target spells it,
// including `../../customXml/item1.xml`, whose extra `..` stays at the package
// root (RFC 3986 dot-segment removal); the image spellings and a reached
// glossary's `../customXml/` target (under word/) open clean.
describe("the rels-broken word-valid verdict matches Word's session 54 rows (#734)", () => {
    it.each([
        ["../customXml/item1.xml"],
        ["/customXml/item1.xml"],
        ["..\\customXml\\item1.xml"],
        ["../CUSTOMXML/%69tem1.xml"],
        ["../../customXml/item1.xml"],
    ])("blocks a customXml relationship to an absent %s", async (target) => {
        const { valid, broken } = await run(withBrokenTarget(target, CUSTOM_XML), "word-valid");
        expect(broken.map((issue) => issue.severity)).toEqual(["error"]);
        expect(valid).toBe(false);
    });

    it.each([["media/image9.png"], ["media/_rels/x.png"], ["notes.rels"], ["word/../customXml/x.png"]])(
        "tolerates an image relationship to an absent %s",
        async (target) => {
            const { valid, broken } = await run(withBrokenTarget(target), "word-valid");
            expect(broken.map((issue) => issue.severity)).toEqual(["warning"]);
            expect(valid).toBe(true);
        },
    );

    it("tolerates a reached glossary's customXml relationship to an absent ../customXml/item1.xml", async () => {
        const { broken } = await run(withBrokenGlossaryTarget("../customXml/item1.xml"), "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([["word/glossary/_rels/document.xml.rels", "warning"]]);
    });
});

describe("the rels-empty-element word-valid verdict reads the registered shape (#735)", () => {
    const shapes = async (parts: Record<string, string>) =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile: "word-valid" });
            return {
                valid: result.valid,
                empty: result.issues.filter((issue) => issue.code === "rels-empty-element").map((issue) => issue.severity),
            };
        });

    it("keeps body content a word-valid warning even when the Id spells the blocking text", async () => {
        const { valid, empty } = await shapes(
            withRelationship(
                `<Relationship Id="missing required attribute" Type="${IMAGE}" Target="media/missing.png">body</Relationship>`,
            ),
        );
        expect(empty).toEqual(["warning"]);
        expect(valid).toBe(true);
    });

    it("keeps a genuinely missing attribute word-blocking", async () => {
        const { valid, empty } = await shapes(withRelationship(`<Relationship Type="${IMAGE}" Target="media/missing.png"/>`));
        expect(empty).toEqual(["error"]);
        expect(valid).toBe(false);
    });

    it("keeps body content with an ordinary Id a word-valid warning", async () => {
        const { valid, empty } = await shapes(
            withRelationship(`<Relationship Id="rId9" Type="${IMAGE}" Target="media/missing.png">body</Relationship>`),
        );
        expect(empty).toEqual(["warning"]);
        expect(valid).toBe(true);
    });
});
