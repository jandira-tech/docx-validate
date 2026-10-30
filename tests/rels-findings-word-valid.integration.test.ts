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

const run = async (parts: Record<string, string | Uint8Array>, profile: Profile) =>
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

// Sessions 71, 72 and 73 (docs/plans/evidence/word-probes-2026-10-24.md, s71a/s71b/s72a/s72b/s73a; two runs each, all four
// readings agree): Word offered recovery for every internal relationship target naming a folder — a final `.` or `..`
// segment, a trailing `/`, or `.` — beside the stored part its path spells, whatever the relationship type and whether
// anything used it, and opened each plain target clean. The file-reference walk resolved a file-system path, which
// drops the final `/`, so it found the part and word-valid accepted every refused row.
describe("a relationship target naming a folder blocks word-valid (#975, sessions 71, 72 and 73)", () => {
    const PNG = Uint8Array.from(
        Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64"),
    );
    const CT = "[Content_Types].xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`;
    const override = (part: string, type: string) =>
        MINIMAL_PARTS[CT]!.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);

    const customXmlAt = (target: string) => ({
        ...withRelationship(`<Relationship Id="rIdX" Type="${CUSTOM_XML}" Target="${target}"/>`),
        "customXml/item1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<root xmlns="urn:s71"/>`,
    });
    const settingsAt = (target: string) => ({
        ...withRelationship(`<Relationship Id="rIdS" Type="${RELATIONSHIPS}/settings" Target="${target}"/>`),
        "word/settings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W}"><w:zoom w:percent="100"/></w:settings>`,
        [CT]: override("word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
    });
    const stylesAt = (target: string) => ({
        ...MINIMAL_PARTS,
        [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace('Target="styles.xml"', `Target="${target}"`),
    });
    const headerAt = (target: string, used: boolean) => ({
        ...withRelationship(`<Relationship Id="rIdH" Type="${RELATIONSHIPS}/header" Target="${target}"/>`),
        ...(used
            ? {
                  "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                      `xmlns:w="${W}"`,
                      `xmlns:w="${W}" xmlns:r="${RELATIONSHIPS}"`,
                  ).replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'),
              }
            : {}),
        "word/header1.xml": HEADER,
        [CT]: override("word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
    });
    const imageAt = (target: string) => ({
        ...withRelationship(`<Relationship Id="rIdI" Type="${IMAGE}" Target="${target}"/>`),
        "word/media/image1.png": PNG,
        [CT]: MINIMAL_PARTS[CT]!.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
    });
    const coreAt = (target: string) => ({
        ...MINIMAL_PARTS,
        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="${target}"/></Relationships>`,
        ),
        "docProps/core.xml":
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
            `xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>S72</dc:title></cp:coreProperties>`,
        [CT]: override("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    });

    it.each([
        ["s71a customXml", SIDECAR, customXmlAt("../customXml/item1.xml/.")],
        ["s71a customXml", SIDECAR, customXmlAt("../customXml/item1.xml/x/..")],
        ["s71b styles", SIDECAR, stylesAt("styles.xml/.")],
        ["s71b used header", SIDECAR, headerAt("header1.xml/.", true)],
        ["s72a unused image", SIDECAR, imageAt("media/image1.png/.")],
        ["s72a unused image", SIDECAR, imageAt(".")],
        ["s72a unused header", SIDECAR, headerAt("header1.xml/.", false)],
        ["s72b core properties", "_rels/.rels", coreAt("docProps/core.xml/.")],
        ["s72b styles", SIDECAR, stylesAt("styles.xml/")],
        ["s73a settings", SIDECAR, settingsAt("settings.xml/.")],
        ["s73a settings", SIDECAR, settingsAt("settings.xml/x/..")],
        ["s73a styles", SIDECAR, stylesAt("styles.xml/x/..")],
        ["s73a used header", SIDECAR, headerAt("header1.xml/x/..", true)],
    ])("blocks the %s row's folder target, as Word refused it", async (_row, relsPart, parts) => {
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([[relsPart, "error"]]);
        expect(valid).toBe(false);
    });

    it.each([
        ["s71a customXml", customXmlAt("../customXml/item1.xml")],
        ["s71b styles", stylesAt("styles.xml")],
        ["s71b used header", headerAt("header1.xml", true)],
        ["s72a unused image", imageAt("media/image1.png")],
        ["s72a unused header", headerAt("header1.xml", false)],
        ["s72b core properties", coreAt("docProps/core.xml")],
        ["s73a settings", settingsAt("settings.xml")],
    ])("accepts the %s calibration, which Word opened clean", async (_row, parts) => {
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken).toEqual([]);
        expect(valid).toBe(true);
    });
});

// #984: Word stops a relationship target's `..` segments at the package root,
// as RFC 3986's remove_dot_segments does, for every relationship type it was
// shown: it read a stored root customXml item (session 76), applied a root
// caps styles part and rendered a root header (session 75), and rendered a
// root note (session 74), each named by a target that climbs above the root,
// all two runs with every reading agreeing. Both file-reference walks
// resolved a file-system path and called such a target broken, so word-valid
// rejected the customXml rows and every profile reported the part Word reads
// as unreferenced.
describe("an escaping target names the part its climb stops at (#984, sessions 74, 75 and 76)", () => {
    const CT = "[Content_Types].xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    const FILE_REFERENCE_CODES = new Set(["rels-broken", "rels-unreferenced", "rels-target-missing"]);
    const findings = async (parts: Record<string, string>, profile: Profile) =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile });
            return {
                valid: result.valid,
                found: result.issues
                    .filter((issue) => FILE_REFERENCE_CODES.has(issue.code ?? ""))
                    .map((issue) => [issue.code, issue.message]),
            };
        });
    const appendRelationship = (relationship: string) =>
        MINIMAL_PARTS[SIDECAR]!.replace("</Relationships>", `${relationship}</Relationships>`);
    const override = (part: string, type: string) =>
        MINIMAL_PARTS[CT]!.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
    /** s76a: a stored root customXml item with its properties, named from the main document's sidecar by `target`. */
    const customXmlAt = (target: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "customXml/item1.xml": `${DECL}<root xmlns="urn:example:data"><v>1</v></root>`,
        "customXml/itemProps1.xml":
            `${DECL}<ds:datastoreItem ds:itemID="{11111111-2222-3333-4444-555555555555}" ` +
            `xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`,
        "customXml/_rels/item1.xml.rels":
            `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="${RELATIONSHIPS}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
        [SIDECAR]: appendRelationship(`<Relationship Id="rIdX" Type="${CUSTOM_XML}" Target="${target}"/>`),
        [CT]: override("customXml/itemProps1.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"),
    });
    /** s75a: the minimal styles part stored at `stored`, its Override moved there, named by `target`. */
    const stylesAt = (stored: string, target: string): Record<string, string> => {
        const parts: Record<string, string> = {
            ...MINIMAL_PARTS,
            [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace('Target="styles.xml"', `Target="${target}"`),
            [CT]: MINIMAL_PARTS[CT]!.replace('PartName="/word/styles.xml"', `PartName="/${stored}"`),
        };
        delete parts["word/styles.xml"];
        parts[stored] = MINIMAL_PARTS["word/styles.xml"]!;
        return parts;
    };
    /** s75a: a used default header stored at `stored`, named by `target`. */
    const headerAt = (stored: string, target: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            `xmlns:w="${W}"`,
            `xmlns:w="${W}" xmlns:r="${RELATIONSHIPS}"`,
        ).replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'),
        [stored]: `${DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
        [SIDECAR]: appendRelationship(`<Relationship Id="rIdH" Type="${RELATIONSHIPS}/header" Target="${target}"/>`),
        [CT]: override(stored, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
    });
    /** s74a: footnote 1, which the body references, stored at `stored` and named by `target`. */
    const footnotesAt = (stored: string, target: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            "</w:r></w:p>",
            '</w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
        ),
        [stored]: `${DECL}<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>`,
        [SIDECAR]: appendRelationship(`<Relationship Id="rIdN" Type="${RELATIONSHIPS}/footnotes" Target="${target}"/>`),
        [CT]: override(stored, "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"),
    });

    const OPENED_CLEAN: [string, Record<string, string>][] = [
        ["s76a customXml ../../customXml/item1.xml", customXmlAt("../../customXml/item1.xml")],
        ["s76a customXml ../../../customXml/item1.xml", customXmlAt("../../../customXml/item1.xml")],
        ["s76a customXml /../customXml/item1.xml", customXmlAt("/../customXml/item1.xml")],
        ["s75a styles ../../styles.xml", stylesAt("styles.xml", "../../styles.xml")],
        ["s75a styles /../styles.xml", stylesAt("styles.xml", "/../styles.xml")],
        ["s75a used header ../../header1.xml", headerAt("header1.xml", "../../header1.xml")],
        ["s74a footnotes /../footnotes.xml", footnotesAt("footnotes.xml", "/../footnotes.xml")],
        ["s74a footnotes /../word/footnotes.xml", footnotesAt("word/footnotes.xml", "/../word/footnotes.xml")],
    ];

    it.each(OPENED_CLEAN.slice(0, 3))("accepts the %s row under word-valid, as Word opened it clean", async (_row, parts) => {
        const { valid, found } = await findings(parts, "word-valid");
        expect(found).toEqual([]);
        expect(valid).toBe(true);
    });

    it.each(OPENED_CLEAN)("reports no file-reference finding for the %s row under any profile", async (_row, parts) => {
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            expect((await findings(parts, profile)).found, profile).toEqual([]);
        }
    });

    it.each([
        ["s76a customXml ../customXml/item1.xml", customXmlAt("../customXml/item1.xml")],
        ["s76a customXml /customXml/item1.xml", customXmlAt("/customXml/item1.xml")],
        ["s75a styles ../styles.xml", stylesAt("styles.xml", "../styles.xml")],
    ])("reports no file-reference finding for the %s calibration under any profile", async (_row, parts) => {
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            expect((await findings(parts, profile)).found, profile).toEqual([]);
        }
    });

    // s75a: with the styles part only in word/, `../../styles.xml` names the
    // absent `/styles.xml`; Word opened the package clean but applied none of
    // `word/styles.xml`, so the climb does not fall back to the source folder.
    it("keeps a climb whose stopped part is absent broken, naming the stopped part (s75a)", async () => {
        const { found } = await findings(stylesAt("word/styles.xml", "../../styles.xml"), "strict");
        expect(found).toEqual([
            ["rels-broken", "Broken reference to ../../styles.xml"],
            ["rels-unreferenced", "Unreferenced file: word/styles.xml"],
            ["rels-target-missing", "Relationship target '../../styles.xml' resolves to 'styles.xml' which does not exist"],
        ]);
    });
});

describe("a relative target that starts with http is resolved as a part (#988, session 79)", () => {
    const PNG = Uint8Array.from(
        Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64"),
    );
    /** s79a: an unused image stored as word/httpimage1.png, related from the main document's sidecar by `target`. */
    const httpImageAt = (target: string) => ({
        ...withRelationship(`<Relationship Id="rIdI" Type="${IMAGE}" Target="${target}"/>`),
        "word/httpimage1.png": PNG,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            '<Default Extension="png" ContentType="image/png"/></Types>',
        ),
    });

    // Word offered recovery for `httpimage1.png/.` in both runs, as for
    // every other folder target, and opened the plain `httpimage1.png`
    // clean: the `http` prefix makes no URI.
    it("blocks the folder target httpimage1.png/. under word-valid, as Word refused it", async () => {
        const { valid, broken } = await run(httpImageAt("httpimage1.png/."), "word-valid");
        expect(valid).toBe(false);
        expect(broken.map((issue) => [issue.severity, issue.message])).toEqual([
            ["error", "Broken reference to httpimage1.png/.: it names a folder, not a part"],
        ]);
    });

    it("accepts the plain httpimage1.png under word-valid, as Word opened it clean", async () => {
        const { valid, broken } = await run(httpImageAt("httpimage1.png"), "word-valid");
        expect(broken).toEqual([]);
        expect(valid).toBe(true);
    });

    it.each(["word-valid", "strict"] as const)("reports an absent http_item.xml as broken (%s)", async (profile) => {
        const { broken } = await run(withBrokenTarget("http_item.xml", CUSTOM_XML), profile);
        expect(broken.map((issue) => issue.message)).toEqual(["Broken reference to http_item.xml"]);
    });

    it.each(["word-valid", "strict", "lenient"] as const)("still skips an http URI, internal or not (%s)", async (profile) => {
        expect((await run(withBrokenTarget("http://example.com/x.png"), profile)).broken).toEqual([]);
        expect((await run(withBrokenTarget("https://example.com/x.png"), profile)).broken).toEqual([]);
    });

    // Lenient keeps its historical prefix scan, byte for byte (#665).
    it("keeps lenient's historical prefix scan", async () => {
        expect((await run(withBrokenTarget("http_item.xml", CUSTOM_XML), "lenient")).broken).toEqual([]);
        expect((await run(withBrokenTarget("httpfoo:bar"), "lenient")).broken).toEqual([]);
    });

    // Word-valid and strict skip only the `http:` and `https:` schemes: another
    // scheme that merely starts with `http` names no part, as `ftp:` does
    // (session 87, #1127). An image relationship to one blocks word-valid in
    // the main document's sidecar, as Word refused an unused image related
    // through `ftp:image1.png` (session 88, #1178).
    it.each([
        ["httpfoo:bar", "word-valid", "error"],
        ["httpfoo:bar", "strict", "error"],
        ["http+unix:/x", "word-valid", "error"],
        ["http+unix:/x", "strict", "error"],
    ] as const)("reports %s as naming no part under %s, as ftp: is", async (target, profile, severity) => {
        const { broken } = await run(withBrokenTarget(target), profile);
        expect(broken.map((issue) => [issue.severity, issue.message])).toEqual([
            [severity, `Broken reference to ${target}: it spells a URI scheme, which names no part`],
        ]);
    });

    // #1008: a URI scheme is ASCII case-insensitive (RFC 3986 §3.1), so
    // word-valid and strict skip an `HTTP:` or `MAILTO:` URI as they skip
    // its lowercase spelling, where they resolved it as a part.
    it.each(["HTTP://example.com/x.png", "Http:x", "HTTP://x/.", "HTTPS://example.com/x.png", "MAILTO:a@b.c", "Mailto:a@b.c"])(
        "skips the URI %s under word-valid and strict, as its lowercase spelling (#1008)",
        async (target) => {
            expect((await run(withBrokenTarget(target), "word-valid")).broken).toEqual([]);
            expect((await run(withBrokenTarget(target), "strict")).broken).toEqual([]);
        },
    );

    it.each(["HTTP://example.com/x.png", "Http:x", "HTTP://x/.", "HTTPS://example.com/x.png", "MAILTO:a@b.c", "Mailto:a@b.c"])(
        "keeps lenient's case-sensitive scan for %s (#665, #1008)",
        async (target) => {
            const { valid, broken } = await run(withBrokenTarget(target), "lenient");
            expect(broken.map((issue) => [issue.severity, issue.message])).toEqual([["error", `Broken reference to ${target}`]]);
            expect(valid).toBe(false);
        },
    );
});

// #987: Word checks a relationship's target only in the relationships parts
// it reads to reach the stories: `_rels/.rels` and a story root's own sidecar.
// Sessions 79 and 80 (two runs each, all readings agree) opened clean a folder
// target in the sidecar of an unreferenced part, of a used header (whose text
// Word rendered), of an unused header and of the styles part, and an absent
// escaping customXml target in the sidecar of the unreferenced part and of the
// used header, while the same folder target in the main document's sidecar got
// the recovery offer (s79a, repeating s72a). The customXml target in an unused
// header's or the styles part's sidecar is unprobed (#1006).
describe("a target in a relationships part Word does not read never blocks word-valid (#987, sessions 79 and 80)", () => {
    const PNG = Uint8Array.from(
        Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64"),
    );
    const CT = "[Content_Types].xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    const sidecarHolding = (relationship: string) =>
        `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
    const image = (target: string) => `<Relationship Id="rId1" Type="${IMAGE}" Target="${target}"/>`;
    const customXml = (target: string) => `<Relationship Id="rId1" Type="${CUSTOM_XML}" Target="${target}"/>`;
    const withPng = (types: string, extra = "") =>
        types.replace("</Types>", `<Default Extension="png" ContentType="image/png"/>${extra}</Types>`);

    /** s79a: the stored unused word/media/image1.png, and an unreferenced word/stray.xml whose sidecar holds `relationship`. */
    const strayHolding = (relationship: string) => ({
        ...MINIMAL_PARTS,
        "word/media/image1.png": PNG,
        "word/stray.xml": `${DECL}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Stray</w:t></w:r></w:p></w:body></w:document>`,
        "word/_rels/stray.xml.rels": sidecarHolding(relationship),
        [CT]: withPng(MINIMAL_PARTS[CT]!),
    });
    /**
     * s80a: the stored unused image, and word/header1.xml related from the main
     * document's sidecar, referenced by the section when `used`, whose own
     * sidecar holds `relationship`.
     */
    const headerHolding = (relationship: string, used: boolean) => ({
        ...withRelationship(`<Relationship Id="rIdH" Type="${RELATIONSHIPS}/header" Target="header1.xml"/>`),
        ...(used
            ? {
                  "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                      `xmlns:w="${W}"`,
                      `xmlns:w="${W}" xmlns:r="${RELATIONSHIPS}"`,
                  ).replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'),
              }
            : {}),
        "word/media/image1.png": PNG,
        "word/header1.xml": `${DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
        "word/_rels/header1.xml.rels": sidecarHolding(relationship),
        [CT]: withPng(
            MINIMAL_PARTS[CT]!,
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
        ),
    });
    /** s80a: the stored unused image, and the styles part's sidecar holding `relationship`. */
    const stylesHolding = (relationship: string) => ({
        ...MINIMAL_PARTS,
        "word/media/image1.png": PNG,
        "word/_rels/styles.xml.rels": sidecarHolding(relationship),
        [CT]: withPng(MINIMAL_PARTS[CT]!),
    });

    const OPENED_CLEAN: [string, string, Record<string, string | Uint8Array>][] = [
        ["s79a stray media/image1.png/.", "word/_rels/stray.xml.rels", strayHolding(image("media/image1.png/."))],
        ["s79a stray .", "word/_rels/stray.xml.rels", strayHolding(image("."))],
        ["s79a stray absent ../customXml/item1.xml", "word/_rels/stray.xml.rels", strayHolding(customXml("../customXml/item1.xml"))],
        ["s80a used header media/image1.png/.", "word/_rels/header1.xml.rels", headerHolding(image("media/image1.png/."), true)],
        [
            "s80a used header absent ../customXml/item1.xml",
            "word/_rels/header1.xml.rels",
            headerHolding(customXml("../customXml/item1.xml"), true),
        ],
        ["s80a unused header media/image1.png/.", "word/_rels/header1.xml.rels", headerHolding(image("media/image1.png/."), false)],
        ["s80a styles media/image1.png/.", "word/_rels/styles.xml.rels", stylesHolding(image("media/image1.png/."))],
    ];

    it.each(OPENED_CLEAN)("accepts the %s row with a warning, as Word opened it clean", async (_row, relsPart, parts) => {
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([[relsPart, "warning"]]);
        expect(valid).toBe(true);
    });

    it.each(OPENED_CLEAN)("keeps the %s row an error under strict", async (_row, relsPart, parts) => {
        const { valid, broken } = await run(parts, "strict");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([[relsPart, "error"]]);
        expect(valid).toBe(false);
    });

    // The same targets in the main document's sidecar: Word refused the
    // folder target (s79a, s72a) and the absent customXml target (session 54).
    it("keeps blocking media/image1.png/. in the main document's sidecar", async () => {
        const parts = {
            ...withRelationship(`<Relationship Id="rIdI" Type="${IMAGE}" Target="media/image1.png/."/>`),
            "word/media/image1.png": PNG,
            [CT]: withPng(MINIMAL_PARTS[CT]!),
        };
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([[SIDECAR, "error"]]);
        expect(valid).toBe(false);
    });

    it("keeps blocking an absent ../customXml/item1.xml in the main document's sidecar", async () => {
        const { valid, broken } = await run(withBrokenTarget("../customXml/item1.xml", CUSTOM_XML), "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([[SIDECAR, "error"]]);
        expect(valid).toBe(false);
    });

    // A reached glossary's sidecar is a story root's, which the read model
    // lists (sessions 17 and 36), so its targets keep the main document's
    // verdict; a folder target there is unprobed.
    it.each([["item1.xml/."], ["../../customXml/item1.xml"]])("keeps blocking %s in a reached glossary's sidecar", async (target) => {
        const { valid, broken } = await run(withBrokenGlossaryTarget(target), "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([["word/glossary/_rels/document.xml.rels", "error"]]);
        expect(valid).toBe(false);
    });
});

// #990: `validateFileReferences` resolved an off-convention relationships
// part's targets against the package root, while the orphan walk resolved
// them against the folder of the part the relationships part belongs to: the
// folder above a `_rels` folder, or the part's own folder outside one. The
// same Target then named two parts.
describe("both file-reference walks read an off-convention relationships part from one folder (#990)", () => {
    const withStylesRelationship = (relsPart: string, target: string) => ({
        ...MINIMAL_PARTS,
        [relsPart]:
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="${RELATIONSHIPS}/styles" Target="${target}"/></Relationships>`,
    });
    const walkFindings = async (parts: Record<string, string>) =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile: "strict" });
            return result.issues
                .filter((issue) => issue.code === "rels-broken" || issue.code === "rels-target-missing")
                .map((issue) => [issue.code, issue.path]);
        });

    it.each(["word/_rels/.rels", "word/notes.rels"])("finds %s's styles.xml in word/ in both walks", async (relsPart) => {
        expect(await walkFindings(withStylesRelationship(relsPart, "styles.xml"))).toEqual([]);
    });

    it.each(["word/_rels/.rels", "word/notes.rels"])("reports %s's word/styles.xml as broken in both walks", async (relsPart) => {
        expect(await walkFindings(withStylesRelationship(relsPart, "word/styles.xml"))).toEqual([
            ["rels-broken", relsPart],
            ["rels-target-missing", relsPart],
        ]);
    });
});

// Session 87 (s87c, s87d; two runs each, every reading agreeing, docs/plans/evidence/word-probes-2026-10-24.md): Word
// offered recovery for a used header related through `ftp:header1.xml` or `httpfoo:header1.xml`, and for the styles
// relationship through `ftp:styles.xml`, beside the stored `word/ftp:…` part their path spells, malformed or
// well-formed, and opened the plain-named control clean. A target spelling a URI scheme names no part (RFC 3986
// §4.3: it is an absolute URI, never a relative reference), as the read model already reads it (#1127); the
// file-reference walk resolved it as a relative reference and found the stored part, so word-valid and strict accepted
// every refused well-formed row.
describe("an internal target spelling a URI scheme names no part (#1127, session 87)", () => {
    const CT = "[Content_Types].xml";
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`;
    const override = (ct: string, part: string, type: string) =>
        ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
    const usedHeaderAt = (target: string) => ({
        ...withRelationship(`<Relationship Id="rIdH" Type="${RELATIONSHIPS}/header" Target="${target}"/>`),
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            `xmlns:w="${W}"`,
            `xmlns:w="${W}" xmlns:r="${RELATIONSHIPS}"`,
        ).replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'),
        [`word/${target}`]: HEADER,
        [CT]: override(MINIMAL_PARTS[CT]!, `word/${target}`, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
    });
    const stylesAt = (target: string) => {
        const parts: Record<string, string> = {
            ...MINIMAL_PARTS,
            [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace('Target="styles.xml"', `Target="${target}"`),
            [`word/${target}`]: MINIMAL_PARTS["word/styles.xml"]!,
        };
        if (target !== "styles.xml") delete parts["word/styles.xml"];
        return parts;
    };

    it.each([
        ["s87c used header", usedHeaderAt("ftp:header1.xml")],
        ["s87c used header", usedHeaderAt("httpfoo:header1.xml")],
        ["s87c used header", usedHeaderAt("HTTPFOO:header1.xml")],
        ["s87d styles", stylesAt("ftp:styles.xml")],
    ])("blocks the %s row's scheme target beside its well-formed stored part, as Word refused it", async (_row, parts) => {
        for (const profile of ["word-valid", "strict"] as const) {
            const { valid, broken } = await run(parts, profile);
            expect(broken.map((issue) => [issue.path, issue.severity, issue.message])).toEqual([
                [SIDECAR, "error", expect.stringMatching(/: it spells a URI scheme, which names no part$/)],
            ]);
            expect(valid).toBe(false);
        }
    });

    it.each([
        ["s87c used header", usedHeaderAt("header1.xml")],
        ["s87d styles", stylesAt("styles.xml")],
    ])("accepts the %s control's plain target", async (_row, parts) => {
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken).toEqual([]);
        expect(valid).toBe(true);
    });

    // Session 88 (s88a–s88c, two runs each, every reading agreeing, #1178): Word offered recovery for every other
    // placement it was shown beside the well-formed stored `ftp:…` part — a header no w:headerReference uses, a
    // `_rels/.rels` styles relationship beside the main document's own, an unused image, and the main document's
    // settings, numbering, comments, footnotes and a used and an unused footer — and opened each plain-named
    // calibration clean. The scheme target therefore blocks whatever the relationship type, in the relationships parts
    // Word reads to reach the stories, as a folder target does (#975).
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const story =
        (segment: string, xml: string, sect = "") =>
        (target: string) => ({
            ...withRelationship(`<Relationship Id="rIdX" Type="${RELATIONSHIPS}/${segment}" Target="${target}"/>`),
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                `xmlns:w="${W}"`,
                `xmlns:w="${W}" xmlns:r="${RELATIONSHIPS}"`,
            ).replace("<w:sectPr>", `<w:sectPr>${sect}`),
            [`word/${target}`]: xml,
            [CT]: override(MINIMAL_PARTS[CT]!, `word/${target}`, `${WML}.${segment}+xml`),
        });
    const FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Foot</w:t></w:r></w:p></w:ftr>`;
    const FOOTNOTES =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="${W}">` +
        `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
        `</w:footnotes>`;
    const S88: [string, (prefix: string) => Record<string, string>, string][] = [
        ["s88a unused header", (p) => story("header", HEADER)(`${p}header1.xml`), SIDECAR],
        [
            "s88a package-level styles relationship",
            (p) => ({
                ...MINIMAL_PARTS,
                "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rIdS" Type="${RELATIONSHIPS}/styles" Target="${p}styles.xml"/></Relationships>`,
                ),
                [`${p}styles.xml`]: MINIMAL_PARTS["word/styles.xml"]!,
                [CT]: override(MINIMAL_PARTS[CT]!, `${p}styles.xml`, `${WML}.styles+xml`),
            }),
            "_rels/.rels",
        ],
        [
            "s88a unused image",
            (p) => ({
                ...withRelationship(`<Relationship Id="rIdI" Type="${IMAGE}" Target="${p}image1.png"/>`),
                [`word/${p}image1.png`]: "png",
            }),
            SIDECAR,
        ],
        [
            "s88b settings",
            (p) =>
                story(
                    "settings",
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W}"/>`,
                )(`${p}settings1.xml`),
            SIDECAR,
        ],
        [
            "s88b numbering",
            (p) =>
                story(
                    "numbering",
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}"/>`,
                )(`${p}numbering1.xml`),
            SIDECAR,
        ],
        [
            "s88b comments",
            (p) =>
                story(
                    "comments",
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W}"/>`,
                )(`${p}comments1.xml`),
            SIDECAR,
        ],
        ["s88c footnotes", (p) => story("footnotes", FOOTNOTES)(`${p}footnotes1.xml`), SIDECAR],
        [
            "s88c used footer",
            (p) => story("footer", FOOTER, '<w:footerReference w:type="default" r:id="rIdX"/>')(`${p}footer1.xml`),
            SIDECAR,
        ],
        ["s88c unused footer", (p) => story("footer", FOOTER)(`${p}footer1.xml`), SIDECAR],
    ];

    it.each(S88)("blocks the %s row's ftp: target beside its well-formed stored part, as Word refused it", async (_row, make, relsPart) => {
        for (const profile of ["word-valid", "strict"] as const) {
            const { valid, broken } = await run(make("ftp:"), profile);
            expect(broken.map((issue) => [issue.path, issue.severity, issue.message])).toEqual([
                [relsPart, "error", expect.stringMatching(/: it spells a URI scheme, which names no part$/)],
            ]);
            expect(valid).toBe(false);
        }
    });

    it.each(S88)("draws no rels-broken for the %s calibration's plain target", async (_row, make) => {
        expect((await run(make(""), "word-valid")).broken).toEqual([]);
    });

    // Outside the relationships parts Word reads to reach the stories the shape is unprobed, so it keeps the
    // warning, as a folder target does there (#987): here a used header's own sidecar.
    it("keeps a scheme target in a used header's own sidecar a word-valid warning, as Word was not shown it", async () => {
        const parts = {
            ...usedHeaderAt("header1.xml"),
            "word/_rels/header1.xml.rels": withBrokenTarget("ftp:image9.png")[SIDECAR]!.replace(
                MINIMAL_PARTS[SIDECAR]!.match(/<Relationship [^>]*\/>/)![0],
                "",
            ),
        };
        const { valid, broken } = await run(parts, "word-valid");
        expect(broken.map((issue) => [issue.path, issue.severity])).toEqual([["word/_rels/header1.xml.rels", "warning"]]);
        expect(valid).toBe(true);
        expect((await run(parts, "strict")).broken.map((issue) => issue.severity)).toEqual(["error"]);
    });

    it("keeps lenient's historical scan, which resolves ftp: as a part (#665)", async () => {
        const { broken } = await run(stylesAt("ftp:styles.xml"), "lenient");
        expect(broken).toEqual([]);
    });
});
