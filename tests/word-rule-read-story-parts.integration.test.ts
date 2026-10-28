/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * #1113: the Word rules that judged the parts named like a story judge the
 * story parts Word reads, as `word-hyperlink-relationship` does since #1106.
 *
 * Word session 83 (docs/plans/evidence/word-probes-2026-10-24.md; two runs,
 * all four readings agree) put each of eleven Word-blocking shapes, from
 * session 41's fixture pairs, in two placements: a main document stored as
 * `word/main.xml`, which Word refused for all eleven, and an unrelated
 * `word/header83.xml`, typed header and named by no relationship, which Word
 * opened clean for all eleven. The clean twins renamed the same way opened
 * clean. The rows below rebuild those specimens from the same fixtures.
 *
 * The other rules in the same files share the part list. Their shapes are
 * unprobed in these placements; they follow the read model, which every
 * shape Word has been shown there obeys (sessions 51, 60 to 62, 82 and 83).
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CT = "[Content_Types].xml";
const MAIN = "word/document.xml";
const MAIN_RELS = "word/_rels/document.xml.rels";
const RENAMED = "word/main.xml";
const UNRELATED = "word/header83.xml";

type Parts = Map<string, string | Uint8Array>;

const once = (text: string, from: string | RegExp, to: string): string => {
    const count = typeof from === "string" ? text.split(from).length - 1 : (text.match(new RegExp(from.source, "g")) ?? []).length;
    if (count !== 1) throw new Error(`expected exactly one ${String(from)}, found ${count}`);
    return text.replace(from, to);
};

const textOf = (parts: Parts, name: string): string => {
    const value = parts.get(name);
    if (value === undefined) throw new Error(`no ${name}`);
    const text = typeof value === "string" ? value : new TextDecoder("utf-8").decode(value);
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
};

/** A fixture's parts in stored order. */
const load = async (fixture: string): Promise<Parts> => {
    const zip = await JSZip.loadAsync(readFileSync(path.join(FIXTURES, fixture)));
    const parts: Parts = new Map();
    for (const name of Object.keys(zip.files)) {
        if (!zip.files[name]!.dir) parts.set(name, await zip.files[name]!.async("uint8array"));
    }
    return parts;
};

/** The package with its main document stored as word/main.xml; every name that points at it follows (session 83). */
const renameMain = (parts: Parts): Parts => {
    const out: Parts = new Map();
    for (const [name, value] of parts) {
        if (name === MAIN) out.set(RENAMED, value);
        else if (name === MAIN_RELS) out.set("word/_rels/main.xml.rels", value);
        else if (name === "_rels/.rels")
            out.set(name, once(textOf(parts, name), /Target="(\/?)word\/document\.xml"/, 'Target="$1word/main.xml"'));
        else if (name === CT) out.set(name, once(textOf(parts, name), 'PartName="/word/document.xml"', 'PartName="/word/main.xml"'));
        else out.set(name, value);
    }
    return out;
};

/** `text` without its outermost `<tag>` elements, nested ones (a w:sectPr inside a w:sectPrChange) counted. */
const withoutElements = (text: string, tag: string): string => {
    const token = new RegExp(`<${tag}\\b[^>]*?(/?)>|</${tag}>`, "g");
    let out = "";
    let depth = 0;
    let cut = 0;
    for (let match = token.exec(text); match; match = token.exec(text)) {
        const closing = match[0].startsWith("</");
        const selfClosing = !closing && match[1] === "/";
        if (depth === 0 && !closing) {
            out += text.slice(cut, match.index);
            if (selfClosing) {
                cut = token.lastIndex;
                continue;
            }
            depth = 1;
        } else if (closing) {
            depth -= 1;
            if (depth === 0) cut = token.lastIndex;
        } else if (!selfClosing) {
            depth += 1;
        }
    }
    return out + text.slice(cut);
};

/**
 * `clean` plus word/header83.xml, typed header by its Override and named by
 * no relationship, holding the main document's body of `broken` under a
 * w:hdr root with the w:document root's attributes, its w:sectPr elements
 * dropped, and `broken`'s main relationships as its sidecar (session 83).
 */
const withUnrelatedHeader = (clean: Parts, broken: Parts): Parts => {
    const doc = textOf(broken, MAIN);
    const attributes = doc.match(/<w:document\b([^>]*)>/)?.[1];
    const body = doc.match(/<w:body>([\s\S]*)<\/w:body>/)?.[1];
    if (attributes === undefined || body === undefined) throw new Error("no w:document root or w:body");
    const out: Parts = new Map(clean);
    out.set(CT, once(textOf(clean, CT), "</Types>", `<Override PartName="/${UNRELATED}" ContentType="${WML}.header+xml"/></Types>`));
    out.set(
        UNRELATED,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr${attributes}>${withoutElements(body, "w:sectPr")}</w:hdr>`,
    );
    const sidecar = broken.get(MAIN_RELS);
    if (sidecar !== undefined) out.set("word/_rels/header83.xml.rels", sidecar);
    return out;
};

const run = async (parts: Parts, profile: "word-valid" | "strict", code: string) =>
    withTempDir(async (dir) => {
        const zip = new JSZip();
        for (const [name, value] of parts) zip.file(name, value);
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
        const result = await validate(file, { profile });
        return {
            valid: result.valid,
            paths: [...new Set(result.issues.filter((issue) => issue.code === code).map((issue) => issue.path))],
            errors: result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`),
        };
    });

/** Session 41's pairs session 83 probed: the rule, the fixture Word refused, and the twin it opened clean. */
const SESSION_83: readonly (readonly [string, string, string])[] = [
    ["word-row-without-cell", "broken/package.word-row-without-cell.docx", "working/package.word-row-without-cell-repaired.docx"],
    [
        "word-cell-ends-with-table",
        "broken/package.word-cell-ends-with-table.docx",
        "working/package.word-cell-ends-with-table-repaired.docx",
    ],
    ["word-drawing-arc-radius", "broken/package.drawing-arc-radius.docx", "working/package.drawing-arc-radius-integer.docx"],
    [
        "word-drawing-extent-negative",
        "broken/package.word-drawing-extent-negative.docx",
        "working/package.word-drawing-extent-positive.docx",
    ],
    ["word-empty-drawing", "broken/package.word-empty-drawing.docx", "working/package.word-empty-drawing-removed.docx"],
    ["word-anchor-missing-attribute", "broken/package.word-anchor-open.docx", "working/package.word-anchor-open-repaired.docx"],
    [
        "word-text-contains-paragraph",
        "broken/package.word-text-contains-paragraph.docx",
        "working/package.word-text-contains-paragraph-lifted.docx",
    ],
    [
        "word-textbox-deleted-text",
        "broken/package.word-textbox-deleted-text-shape.docx",
        "working/package.word-textbox-deleted-text-shape-repaired.docx",
    ],
    ["word-move-from-deleted-text", "broken/package.word-move-from-deleted-text.docx", "working/package.word-move-from-text.docx"],
    ["word-sdt-id-empty", "broken/package.word-sdt-id-empty.docx", "working/package.word-sdt-id-empty-repaired.docx"],
    [
        "word-choice-requires-undeclared",
        "broken/package.word-choice-requires-undeclared.docx",
        "working/package.word-choice-requires-declared.docx",
    ],
];

const PROFILES = ["word-valid", "strict"] as const;

describe("the Word rules judge the story parts Word reads, session 83's rows (#1113)", { timeout: 120_000 }, () => {
    it.each(SESSION_83.flatMap(([rule, broken]) => PROFILES.map((profile) => [rule, profile, broken] as const)))(
        "%s judges the shape in a main document stored as word/main.xml under %s, as Word refused it",
        async (rule, profile, broken) => {
            const found = await run(renameMain(await load(broken)), profile, rule);
            expect(found.paths).toEqual([RENAMED]);
            expect(found.valid).toBe(false);
        },
    );

    it.each(SESSION_83.flatMap(([rule, broken, clean]) => PROFILES.map((profile) => [rule, profile, broken, clean] as const)))(
        "%s leaves the shape in an unrelated word/header83.xml unjudged under %s, as Word opened it clean",
        async (rule, profile, broken, clean) => {
            const found = await run(withUnrelatedHeader(await load(clean), await load(broken)), profile, rule);
            expect(found.paths).toEqual([]);
            if (profile === "word-valid") {
                // #1116: mce-malformed still judges every part under word/,
                // so it alone rejects the undeclared Choice in the unread
                // header, which Word opened clean.
                expect(found.errors).toEqual(rule === "word-choice-requires-undeclared" ? [`mce-malformed@${UNRELATED}`] : []);
            }
        },
    );

    it.each(SESSION_83.map(([rule, , clean]) => [rule, clean] as const))(
        "the %s calibration renamed to word/main.xml passes word-valid, as Word opened it clean",
        async (rule, clean) => {
            expect(await run(renameMain(await load(clean)), "word-valid", rule)).toEqual({ valid: true, paths: [], errors: [] });
        },
    );
});

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NAMESPACES =
    `xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
    `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ` +
    `xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"`;
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const drawing = (inner: string): string => `<w:p><w:r><w:drawing>${inner}</w:drawing></w:r></w:p>`;
const picture = (children: string): string =>
    drawing(
        `<wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="p"/><a:graphic>` +
            `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>${children}</pic:pic></a:graphicData>` +
            `</a:graphic></wp:inline>`,
    );
const table = (properties: string, grid = `<w:gridCol w:w="100"/>`): string =>
    `<w:tbl>${properties}<w:tblGrid>${grid}</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>`;

/** The rules sharing the moved part list whose shapes no Word session has put in these placements, each with a shape it reports. */
const UNPROBED: readonly (readonly [string, string])[] = [
    ["word-table-properties-required", table("")],
    ["word-table-properties-misplaced", table(`<w:tblPr><w:tblLayout w:type="fixed"/><w:tblBorders/></w:tblPr>`)],
    ["word-grid-column-attributes", table("<w:tblPr/>", `<w:gridCol w:w="100" w:type="dxa"/>`)],
    [
        "word-cell-margin-order",
        table(`<w:tblPr><w:tblCellMar><w:right w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`),
    ],
    ["word-drawing-extent-missing", drawing(`<wp:inline><a:graphic/></wp:inline>`)],
    ["word-drawing-extent-misplaced", drawing(`<wp:inline><a:graphic/><wp:extent cx="1" cy="1"/></wp:inline>`)],
    [
        "word-drawing-scalar-whitespace",
        drawing(`<wp:anchor><wp:positionH relativeFrom="column"><wp:posOffset> 1 </wp:posOffset></wp:positionH></wp:anchor>`),
    ],
    [
        "word-position-offset-and-align",
        drawing(
            `<wp:anchor><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset><wp:align>left</wp:align></wp:positionH></wp:anchor>`,
        ),
    ],
    ["word-picture-nvpicpr-missing", picture("<pic:blipFill/><pic:spPr/>")],
    ["word-picture-sppr-missing", picture("<pic:nvPicPr/><pic:blipFill/>")],
    [
        "word-textbox-content-namespace",
        drawing(`<wp:wsp><wp:txbx><wne:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></wne:txbxContent></wp:txbx></wp:wsp>`),
    ],
    ["word-math-run-deletion", `<w:p><m:oMath><m:r><w:del w:id="1" w:author="A"><m:t>x</m:t></w:del></m:r></m:oMath></w:p>`],
    ["word-math-run-insertion", `<w:p><m:oMath><m:r><w:ins w:id="1" w:author="A"><m:t>x</m:t></w:ins></m:r></m:oMath></w:p>`],
    [
        "word-page-margin-attributes",
        `<w:p><w:pPr><w:sectPr><w:pgMar w:top="1" w:right="1" w:bottom="1" w:left="1"/></w:sectPr></w:pPr></w:p>`,
    ],
    ["word-shading-color", `<w:p><w:pPr><w:shd w:val="clear" w:fill=""/></w:pPr></w:p>`],
];

const minimal = (): Parts => new Map(Object.entries(MINIMAL_PARTS));
const mainHolding = (shape: string): Parts => {
    const parts = minimal();
    parts.set(MAIN, `${DECL}<w:document ${NAMESPACES}><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p>${shape}</w:body></w:document>`);
    return parts;
};
const headerHolding = (shape: string): Parts => {
    const parts = minimal();
    parts.set(CT, once(textOf(parts, CT), "</Types>", `<Override PartName="/${UNRELATED}" ContentType="${WML}.header+xml"/></Types>`));
    parts.set(UNRELATED, `${DECL}<w:hdr ${NAMESPACES}>${shape}</w:hdr>`);
    return parts;
};

describe("the rules no session probed in these placements follow the read model too (#1113)", { timeout: 120_000 }, () => {
    it.each(UNPROBED.flatMap(([rule, shape]) => PROFILES.map((profile) => [rule, profile, shape] as const)))(
        "%s judges its shape in a main document stored as word/main.xml under %s",
        async (rule, profile, shape) => {
            // The control: the shape draws the rule where the main document is conventionally named.
            expect((await run(mainHolding(shape), profile, rule)).paths).toEqual([MAIN]);
            expect((await run(renameMain(mainHolding(shape)), profile, rule)).paths).toEqual([RENAMED]);
        },
    );

    it.each(UNPROBED.flatMap(([rule, shape]) => PROFILES.map((profile) => [rule, profile, shape] as const)))(
        "%s leaves its shape in an unrelated word/header83.xml unjudged under %s",
        async (rule, profile, shape) => {
            expect((await run(headerHolding(shape), profile, rule)).paths).toEqual([]);
        },
    );
});
