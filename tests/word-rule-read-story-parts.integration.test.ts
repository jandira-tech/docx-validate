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
import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import * as wordReadParts from "../src/lib/word-read-parts";
import { validate } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

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

/** `parts` packed as dir/specimen.docx. */
const writeSpecimen = async (dir: string, parts: Parts): Promise<string> => {
    const zip = new JSZip();
    for (const [name, value] of parts) zip.file(name, value);
    const file = path.join(dir, "specimen.docx");
    writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    return file;
};

const run = async (parts: Parts, profile: "word-valid" | "strict", code: string) =>
    withTempDir(async (dir) => {
        const result = await validate(await writeSpecimen(dir, parts), { profile });
        return {
            valid: result.valid,
            paths: [...new Set(result.issues.filter((issue) => issue.code === code).map((issue) => issue.path))],
            found: result.issues.filter((issue) => issue.code === code).map((issue) => issue.path),
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
            // #1116: mce-malformed, the undeclared Choice's other finding,
            // warns in a part no relationship names.
            if (profile === "word-valid") expect(found.errors).toEqual([]);
        },
    );

    // #1123 (review 1117 F2): a rule that judged only the main document passed
    // every row above. A header the main document relates and its w:sectPr
    // references is a part Word reads; these shapes are unprobed there.
    it.each(SESSION_83.flatMap(([rule, broken, clean]) => PROFILES.map((profile) => [rule, profile, broken, clean] as const)))(
        "%s judges the shape in a header the main document uses under %s (#1123)",
        async (rule, profile, broken, clean) => {
            const found = await run(relateHeader(withUnrelatedHeader(await load(clean), await load(broken)), true), profile, rule);
            expect(found.paths).toEqual([UNRELATED]);
        },
    );

    it.each(SESSION_83.map(([rule, , clean]) => [rule, clean] as const))(
        "the %s calibration renamed to word/main.xml passes word-valid, as Word opened it clean",
        async (rule, clean) => {
            expect(await run(renameMain(await load(clean)), "word-valid", rule)).toEqual({ valid: true, paths: [], found: [], errors: [] });
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

    // #1123: the same header, related and referenced by the body's w:sectPr.
    it.each(UNPROBED.flatMap(([rule, shape]) => PROFILES.map((profile) => [rule, profile, shape] as const)))(
        "%s judges its shape in a header the main document uses under %s (#1123)",
        async (rule, profile, shape) => {
            expect((await run(relateHeader(headerHolding(shape), true), profile, rule)).paths).toEqual([UNRELATED]);
        },
    );
});

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** The minimal package plus word/hf1.xml holding `shape`, which the body's w:sectPr uses as its header and its footer. */
const headerAndFooterHolding = (shape: string): Parts => {
    const parts = minimal();
    parts.set(CT, once(textOf(parts, CT), "</Types>", `<Override PartName="/word/hf1.xml" ContentType="${WML}.header+xml"/></Types>`));
    parts.set(
        MAIN_RELS,
        once(
            textOf(parts, MAIN_RELS),
            "</Relationships>",
            `<Relationship Id="rIdH" Type="${R}/header" Target="hf1.xml"/><Relationship Id="rIdF" Type="${R}/footer" Target="hf1.xml"/></Relationships>`,
        ),
    );
    parts.set(
        MAIN,
        once(
            textOf(parts, MAIN),
            "<w:sectPr>",
            `<w:sectPr><w:headerReference xmlns:r="${R}" w:type="default" r:id="rIdH"/><w:footerReference xmlns:r="${R}" w:type="default" r:id="rIdF"/>`,
        ),
    );
    parts.set("word/hf1.xml", `${DECL}<w:hdr ${NAMESPACES}>${shape}</w:hdr>`);
    return parts;
};

// #1118 (review 1115 F1): the read model lists a part once per segment that
// reaches it, so a part the body uses as its header and its footer arrives
// twice; every rule walking wordReadStoryParts judges it once.
describe("a part read through two story segments is judged once (#1118)", { timeout: 120_000 }, () => {
    it.each(PROFILES)("word-sdt-id-empty reports one empty id once in a part used as header and footer under %s", async (profile) => {
        const shape = `<w:sdt><w:sdtPr><w:id w:val=""/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt><w:p/>`;
        expect((await run(headerAndFooterHolding(shape), profile, "word-sdt-id-empty")).found).toEqual(["word/hf1.xml"]);
    });
});

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const CHOICE_BROKEN = "broken/package.word-choice-requires-undeclared.docx";
const CHOICE_CLEAN = "working/package.word-choice-requires-declared.docx";

/** The run's verdict and each markup-compatibility finding as `code:severity@path`, from a full run or a `--checks` run. */
const mceRun = async (parts: Parts, profile: "word-valid" | "strict", checks?: readonly string[]) =>
    withTempDir(async (dir) => {
        const file = await writeSpecimen(dir, parts);
        const result = await validate(file, checks === undefined ? { profile } : { profile, checks });
        return {
            valid: result.valid,
            mce: result.issues
                .filter((issue) => issue.code?.startsWith("mce-"))
                .map((issue) => `${issue.code}:${issue.severity}@${issue.path}`),
        };
    });

/** `parts` with the main document relating word/header83.xml as a header, and its last w:sectPr referencing it when `used`. */
const relateHeader = (parts: Parts, used: boolean): Parts => {
    const out: Parts = new Map(parts);
    out.set(
        MAIN_RELS,
        once(
            textOf(parts, MAIN_RELS),
            "</Relationships>",
            `<Relationship Id="rId83" Type="${R}/header" Target="header83.xml"/></Relationships>`,
        ),
    );
    if (used) {
        const main = textOf(parts, MAIN);
        const at = main.lastIndexOf("<w:sectPr");
        const open = main.indexOf(">", at) + 1;
        if (at < 0 || main[open - 2] === "/") throw new Error("no open w:sectPr");
        out.set(MAIN, `${main.slice(0, open)}<w:headerReference xmlns:r="${R}" w:type="default" r:id="rId83"/>${main.slice(open)}`);
    }
    return out;
};

/** The minimal package plus word/header83.xml, typed header and named by no relationship, whose root requires `mc:MustUnderstand="foo"`. */
const mustUnderstandHeader = (): Parts => {
    const parts = minimal();
    parts.set(CT, once(textOf(parts, CT), "</Types>", `<Override PartName="/${UNRELATED}" ContentType="${WML}.header+xml"/></Types>`));
    parts.set(UNRELATED, `${DECL}<w:hdr xmlns:w="${W}" xmlns:mc="${MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"><w:p/></w:hdr>`);
    return parts;
};

const RUNS = [
    ["a full", undefined],
    ["a --checks", ["word-markup-compatibility"]],
] as const;

// #1116: Word opened clean the undeclared-Requires Choice (session 83 s83d) and
// an unknown MustUnderstand (session 84 s84a) in a word/header83.xml no
// relationship names, and refused each where it reads it. #1125: session 84
// opened both clean in a header the main document relates but does not use.
describe("markup compatibility in a word/ part Word does not read warns under word-valid (#1116, #1125)", { timeout: 120_000 }, () => {
    it.each(RUNS)("%s run passes s83d's unrelated header, its mce-malformed a warning, as Word opened it clean", async (_label, checks) => {
        const parts = withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN));
        expect(await mceRun(parts, "word-valid", checks)).toEqual({ valid: true, mce: [`mce-malformed:warning@${UNRELATED}`] });
    });

    it.each(RUNS)(
        "%s run keeps mce-malformed an error in the main document stored as word/main.xml, as Word refused it",
        async (_label, checks) => {
            expect(await mceRun(renameMain(await load(CHOICE_BROKEN)), "word-valid", checks)).toEqual({
                valid: false,
                mce: [`mce-malformed:error@${RENAMED}`],
            });
        },
    );

    it("strict keeps mce-malformed an error in the unrelated header", async () => {
        const parts = withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN));
        expect((await mceRun(parts, "strict")).mce).toContain(`mce-malformed:error@${UNRELATED}`);
    });

    it.each(
        RUNS.flatMap(
            ([label, checks]) =>
                [
                    ["a used header the main document relates", label, checks, (parts: Parts) => relateHeader(parts, true)],
                    [
                        "a part only a relationships part with no stored source names",
                        label,
                        checks,
                        (parts: Parts) => {
                            const out: Parts = new Map(parts);
                            out.set(
                                "word/_rels/stray.xml.rels",
                                `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header83.xml"/></Relationships>`,
                            );
                            return out;
                        },
                    ],
                ] as const,
        ),
    )("keeps mce-malformed an error in %s (%s run), a part a relationship names", async (_placement, _label, checks, place) => {
        const parts = place(withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN)));
        expect((await mceRun(parts, "word-valid", checks)).mce).toContain(`mce-malformed:error@${UNRELATED}`);
    });

    // A packed file with a refused item name judges the rest through the
    // archive (#629); its --checks run reads the findings there (#454).
    it.each(RUNS)("%s run of a package with a refused item name warns there too", async (_label, checks) => {
        const judge = async (parts: Parts) =>
            withTempDir(async (dir) => {
                const file = path.join(dir, "refused.docx");
                writeFileSync(
                    file,
                    buildRawZip([...[...parts].map(([name, content]) => ({ name, content })), { name: "/abs/x.xml", content: "<x/>" }]),
                );
                const result = await validate(file, checks === undefined ? { profile: "word-valid" } : { profile: "word-valid", checks });
                return result.issues
                    .filter((issue) => issue.code?.startsWith("mce-"))
                    .map((issue) => `${issue.code}:${issue.severity}@${issue.path}`);
            });
        expect(await judge(withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN)))).toEqual([
            `mce-malformed:warning@${UNRELATED}`,
        ]);
        expect(await judge(renameMain(await load(CHOICE_BROKEN)))).toEqual([`mce-malformed:error@${RENAMED}`]);
    });

    // Session 84 (s84a): Word opened the unrelated header clean and refused the used one.
    it.each(RUNS)("%s run warns for mce-must-understand in the unrelated header too, as Word opened it clean", async (_label, checks) => {
        expect(await mceRun(mustUnderstandHeader(), "word-valid", checks)).toEqual({
            valid: true,
            mce: [`mce-must-understand:warning@${UNRELATED}`],
        });
        expect((await mceRun(relateHeader(mustUnderstandHeader(), true), "word-valid", checks)).mce).toEqual([
            `mce-must-understand:error@${UNRELATED}`,
        ]);
    });

    // Session 84 (s84a mu-unused-header, choice-unused-header): Word opened both
    // clean in a header the main document relates but no section references.
    it.each(RUNS)(
        "%s run warns in a header the main document relates but does not use, as Word opened it clean (#1125)",
        async (_label, checks) => {
            expect(await mceRun(relateHeader(mustUnderstandHeader(), false), "word-valid", checks)).toEqual({
                valid: true,
                mce: [`mce-must-understand:warning@${UNRELATED}`],
            });
            const choice = relateHeader(withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN)), false);
            expect(await mceRun(choice, "word-valid", checks)).toEqual({ valid: true, mce: [`mce-malformed:warning@${UNRELATED}`] });
        },
    );

    // #1135 (review 1126 F7): a header relationship the read model cannot see
    // still names the header. Word's reading of such a relationships part is
    // unprobed, so the list errs wide and the used header's finding blocks.
    it.each(
        RUNS.flatMap(
            ([label, checks]) =>
                [
                    [
                        "in no namespace",
                        label,
                        checks,
                        (rels: string) => once(rels, 'Target="header83.xml"/>', 'Target="header83.xml" xmlns=""/>'),
                    ],
                    [
                        "under a root outside the package-relationships namespace",
                        label,
                        checks,
                        (rels: string) =>
                            once(
                                rels,
                                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                                '<Relationships xmlns="urn:not-pr">',
                            ),
                    ],
                ] as const,
        ),
    )(
        "keeps mce-must-understand an error in a used header a relationship %s names (%s run, #1135)",
        async (_shape, _label, checks, spell) => {
            const parts = relateHeader(mustUnderstandHeader(), true);
            parts.set(MAIN_RELS, spell(textOf(parts, MAIN_RELS)));
            expect((await mceRun(parts, "word-valid", checks)).mce).toEqual([`mce-must-understand:error@${UNRELATED}`]);
        },
    );

    // #1134 (review 1126 F8): the blocking half keys the finding's part
    // canonically. A used header stored as word/Hdr7.XML and named hdr7.xml
    // is the part the relationship names, so its finding blocks.
    it.each(RUNS)(
        "%s run keeps mce-must-understand an error in a used header stored under an OPC-equivalent spelling (#1134)",
        async (_label, checks) => {
            const STORED = "word/Hdr7.XML";
            const parts = relateHeader(mustUnderstandHeader(), true);
            parts.set(STORED, textOf(parts, UNRELATED));
            parts.delete(UNRELATED);
            parts.set(CT, once(textOf(parts, CT), `PartName="/${UNRELATED}"`, 'PartName="/word/hdr7.xml"'));
            parts.set(MAIN_RELS, once(textOf(parts, MAIN_RELS), 'Target="header83.xml"', 'Target="hdr7.xml"'));
            expect((await mceRun(parts, "word-valid", checks)).mce).toEqual([`mce-must-understand:error@${STORED}`]);
        },
    );

    // #1145 (review 1138 F1, F8): the read model decides whether a root uses
    // its comments or notes part from the root alone, so a part whose only
    // reference sits in a used header was dropped and its finding warned.
    // Whether Word reads it then is unprobed, so the part stays named. An
    // unused header Word does not read (session 84 s84a) names nothing, and
    // nor does a footnote: session 16 (s16a) opened clean a malformed
    // endnotes part only a footnote referenced.
    const NOTE_KINDS = [
        ["footnotes", '<w:footnote w:id="1"><w:p/></w:footnote>', '<w:footnoteReference w:id="1"/>'],
        ["endnotes", '<w:endnote w:id="1"><w:p/></w:endnote>', '<w:endnoteReference w:id="1"/>'],
        ["comments", '<w:comment w:id="0" w:author="A"><w:p/></w:comment>', '<w:commentReference w:id="0"/>'],
    ] as const;
    /** The minimal package plus a related `word/<kind>.xml` requiring `mc:MustUnderstand="foo"`, and word/header83.xml holding `reference`. */
    const notesWithHeaderReference = (kind: string, body: string, reference: string, used: boolean): Parts => {
        const parts = relateHeader(minimal(), used);
        parts.set(UNRELATED, `${DECL}<w:hdr xmlns:w="${W}"><w:p><w:r>${reference}</w:r></w:p></w:hdr>`);
        parts.set(
            `word/${kind}.xml`,
            `${DECL}<w:${kind} xmlns:w="${W}" xmlns:mc="${MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo">${body}</w:${kind}>`,
        );
        parts.set(
            CT,
            once(
                textOf(parts, CT),
                "</Types>",
                `<Override PartName="/${UNRELATED}" ContentType="${WML}.header+xml"/><Override PartName="/word/${kind}.xml" ContentType="${WML}.${kind}+xml"/></Types>`,
            ),
        );
        parts.set(
            MAIN_RELS,
            once(
                textOf(parts, MAIN_RELS),
                "</Relationships>",
                `<Relationship Id="rIdN" Type="${R}/${kind}" Target="${kind}.xml"/></Relationships>`,
            ),
        );
        return parts;
    };

    it.each(
        RUNS.flatMap(([label, checks]) => NOTE_KINDS.map(([kind, body, reference]) => [kind, label, checks, body, reference] as const)),
    )(
        "keeps mce-must-understand an error in a related %s part whose only reference is in a used header (%s run, #1145)",
        async (kind, _label, checks, body, reference) => {
            expect((await mceRun(notesWithHeaderReference(kind, body, reference, true), "word-valid", checks)).mce).toEqual([
                `mce-must-understand:error@word/${kind}.xml`,
            ]);
        },
    );

    it.each(
        RUNS.flatMap(([label, checks]) => NOTE_KINDS.map(([kind, body, reference]) => [kind, label, checks, body, reference] as const)),
    )(
        "warns for mce-must-understand in a related %s part only an unused header references (%s run, #1145)",
        async (kind, _label, checks, body, reference) => {
            expect((await mceRun(notesWithHeaderReference(kind, body, reference, false), "word-valid", checks)).mce).toEqual([
                `mce-must-understand:warning@word/${kind}.xml`,
            ]);
        },
    );

    it.each(RUNS)(
        "%s run warns for mce-must-understand in a related endnotes part only a footnote references (session 16, #1145)",
        async (_label, checks) => {
            const parts = minimal();
            parts.set(
                MAIN,
                once(textOf(parts, MAIN), "<w:t>Hello</w:t></w:r>", '<w:t>Hello</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>'),
            );
            parts.set(
                "word/footnotes.xml",
                `${DECL}<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:footnote></w:footnotes>`,
            );
            parts.set(
                "word/endnotes.xml",
                `${DECL}<w:endnotes xmlns:w="${W}" xmlns:mc="${MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"><w:endnote w:id="1"><w:p/></w:endnote></w:endnotes>`,
            );
            parts.set(
                CT,
                once(
                    textOf(parts, CT),
                    "</Types>",
                    `<Override PartName="/word/footnotes.xml" ContentType="${WML}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="${WML}.endnotes+xml"/></Types>`,
                ),
            );
            parts.set(
                MAIN_RELS,
                once(
                    textOf(parts, MAIN_RELS),
                    "</Relationships>",
                    `<Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdE" Type="${R}/endnotes" Target="endnotes.xml"/></Relationships>`,
                ),
            );
            expect((await mceRun(parts, "word-valid", checks)).mce).toEqual(["mce-must-understand:warning@word/endnotes.xml"]);
        },
    );

    // #1147 (review 1138 F6): the relationships scan runs only for a run that
    // holds one of the two findings, on every path that applies the profile:
    // the full run, a --checks run, and a packed file with a refused item
    // name, whose rest is judged through the archive (#629, #454).
    describe("computes the possibly-read list only for a run with an mce finding (#1147)", () => {
        afterEach(() => {
            vi.restoreAllMocks();
        });
        const scans = async (parts: Parts, checks: readonly string[] | undefined, refused: boolean): Promise<number> =>
            withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                const entries = [...parts].map(([name, content]) => ({ name, content }));
                writeFileSync(file, buildRawZip(refused ? [...entries, { name: "/abs/x.xml", content: "<x/>" }] : entries));
                const spy = vi.spyOn(wordReadParts, "wordPossiblyReadParts");
                await validate(file, checks === undefined ? { profile: "word-valid" } : { profile: "word-valid", checks });
                return spy.mock.calls.length;
            });

        it.each(RUNS.flatMap(([label, checks]) => [false, true].map((refused) => [label, refused, checks] as const)))(
            "%s run (refused item name: %s) scans for a finding and not without one",
            async (_label, refused, checks) => {
                expect(await scans(minimal(), checks, refused)).toBe(0);
                expect(await scans(mustUnderstandHeader(), checks, refused)).toBeGreaterThan(0);
            },
        );
    });

    it("strict keeps mce-malformed an error in a header the main document relates but does not use", async () => {
        const choice = relateHeader(withUnrelatedHeader(await load(CHOICE_CLEAN), await load(CHOICE_BROKEN)), false);
        expect((await mceRun(choice, "strict")).mce).toContain(`mce-malformed:error@${UNRELATED}`);
    });
});
