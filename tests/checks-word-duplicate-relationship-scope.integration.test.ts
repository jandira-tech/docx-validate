/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Session 49 (s49c) and session 50 (s50a/s50b) specimens for #717: Word
 * 16.115 offers recovery for two relationships of one kind only in the
 * relationships part of a story root — the main document, or a glossary
 * document a main-document `glossaryDocument` relationship reaches. The
 * same pair in `_rels/.rels` or in a referenced header's sidecar opens
 * clean, and the real styles relationship beside one or two External
 * styles relationships is still refused. The fontTable and webSettings
 * rows in `_rels/.rels` and the webSettings and theme rows in a glossary
 * sidecar were not shown to Word; they follow the same rule.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { R } from "./support/word-rule-test-xml";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CT = "[Content_Types].xml";
const SIDECAR = "word/_rels/document.xml.rels";
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const P = MINIMAL_PARTS as Record<string, string>;

type Kind = "styles" | "numbering" | "settings" | "fontTable" | "webSettings" | "theme";

/** The part body and content type of each kind (the bodies the session builders used). */
const KINDS: Record<Kind, { root: string; ct: string }> = {
    styles: { root: `${XML_DECL}<w:styles xmlns:w="${W}"/>`, ct: `${WML}.styles+xml` },
    numbering: { root: `${XML_DECL}<w:numbering xmlns:w="${W}"/>`, ct: `${WML}.numbering+xml` },
    settings: { root: `${XML_DECL}<w:settings xmlns:w="${W}"/>`, ct: `${WML}.settings+xml` },
    fontTable: { root: `${XML_DECL}<w:fonts xmlns:w="${W}"/>`, ct: `${WML}.fontTable+xml` },
    webSettings: { root: `${XML_DECL}<w:webSettings xmlns:w="${W}"/>`, ct: `${WML}.webSettings+xml` },
    theme: {
        root: `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
        ct: "application/vnd.openxmlformats-officedocument.theme+xml",
    },
};
const CHECK_OF: Record<Kind, string> = {
    styles: "word-duplicate-styles-relationship",
    numbering: "word-duplicate-numbering-relationship",
    settings: "word-duplicate-settings-relationship",
    fontTable: "word-duplicate-font-table-relationship",
    webSettings: "word-duplicate-web-settings-relationship",
    theme: "word-duplicate-theme-relationship",
};

/** Where a kind's parts live: `theme/theme1.xml` for the theme, `<kind>.xml` otherwise. */
const fileOf = (kind: Kind, n: 1 | 2): string => (kind === "theme" ? `theme/theme${n}.xml` : `${kind}${n === 1 ? "" : "2"}.xml`);

const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};
const relationship = (id: string, segment: string, target: string, external = false): string =>
    `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`;
const relsPart = (body: string): string =>
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const appendRels = (relsXml: string, extra: string): string => replaceOnce(relsXml, "</Relationships>", `${extra}</Relationships>`);
const overrides = (ct: string, ...entries: [string, string][]): string =>
    replaceOnce(
        ct,
        "</Types>",
        `${entries.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join("")}</Types>`,
    );

/** The minimal package plus `word/<file of kind>` related from the main document's sidecar as rId2. Styles is already related (rId1). */
const withRelated = (kind: Kind): Record<string, string> =>
    kind === "styles"
        ? { ...P }
        : {
              ...P,
              [`word/${fileOf(kind, 1)}`]: KINDS[kind].root,
              [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rId2", kind, fileOf(kind, 1))),
              [CT]: overrides(P[CT]!, [`word/${fileOf(kind, 1)}`, KINDS[kind].ct]),
          };

/** Session 45's refused shape: two relationships of `kind` to two parts in the main document's sidecar. */
const dupInMainRels = (kind: Kind): Promise<Buffer> => {
    const parts = withRelated(kind);
    return buildDocx({
        ...parts,
        [`word/${fileOf(kind, 2)}`]: KINDS[kind].root,
        [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rId3", kind, fileOf(kind, 2))),
        [CT]: overrides(parts[CT]!, [`word/${fileOf(kind, 2)}`, KINDS[kind].ct]),
    });
};
/** The main document's own `kind` relationship, and two more of `kind` in `_rels/.rels`. */
const dupInPackageRels = (kind: Kind): Promise<Buffer> => {
    const parts = withRelated(kind);
    return buildDocx({
        ...parts,
        [`word/${fileOf(kind, 2)}`]: KINDS[kind].root,
        "_rels/.rels": appendRels(
            P["_rels/.rels"]!,
            `${relationship("rId8", kind, `word/${fileOf(kind, 1)}`)}${relationship("rId9", kind, `word/${fileOf(kind, 2)}`)}`,
        ),
        [CT]: overrides(parts[CT]!, [`word/${fileOf(kind, 2)}`, KINDS[kind].ct]),
    });
};
const HEADER = `${XML_DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
/** The main document's own `kind` relationship, and a referenced header whose sidecar holds two of `kind`. */
const dupInHeaderRels = (kind: Kind): Promise<Buffer> => {
    const parts = withRelated(kind);
    return buildDocx({
        ...parts,
        [`word/${fileOf(kind, 2)}`]: KINDS[kind].root,
        "word/document.xml": replaceOnce(
            replaceOnce(P["word/document.xml"]!, `xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`),
            "<w:sectPr>",
            `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`,
        ),
        "word/header9.xml": HEADER,
        "word/_rels/header9.xml.rels": relsPart(
            `${relationship("rId1", kind, fileOf(kind, 1))}${relationship("rId2", kind, fileOf(kind, 2))}`,
        ),
        [SIDECAR]: appendRels(parts[SIDECAR]!, relationship("rIdH", "header", "header9.xml")),
        [CT]: overrides(parts[CT]!, ["word/header9.xml", `${WML}.header+xml`], [`word/${fileOf(kind, 2)}`, KINDS[kind].ct]),
    });
};
const GLOSSARY =
    `${XML_DECL}<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
/** A reached glossary whose own sidecar holds two relationships of `kind` to two parts in its folder. */
const dupInGlossaryRels = (kind: Kind): Promise<Buffer> =>
    buildDocx({
        ...P,
        "word/glossary/document.xml": GLOSSARY,
        [`word/glossary/${fileOf(kind, 1)}`]: KINDS[kind].root,
        [`word/glossary/${fileOf(kind, 2)}`]: KINDS[kind].root,
        "word/glossary/_rels/document.xml.rels": relsPart(
            `${relationship("rId1", kind, fileOf(kind, 1))}${relationship("rId2", kind, fileOf(kind, 2))}`,
        ),
        [SIDECAR]: appendRels(P[SIDECAR]!, relationship("rIdG", "glossaryDocument", "glossary/document.xml")),
        [CT]: overrides(
            P[CT]!,
            ["word/glossary/document.xml", `${WML}.document.glossary+xml`],
            [`word/glossary/${fileOf(kind, 1)}`, KINDS[kind].ct],
            [`word/glossary/${fileOf(kind, 2)}`, KINDS[kind].ct],
        ),
    });
/** The main document's real styles relationship beside `count` External ones. */
const withExternalStyles = (count: number): Promise<Buffer> =>
    buildDocx({
        ...P,
        [SIDECAR]: appendRels(
            P[SIDECAR]!,
            Array.from({ length: count }, (_, i) =>
                relationship(`rIdX${i + 1}`, "styles", `http://example.com/styles${i + 1}.xml`, true),
            ).join(""),
        ),
    });

const ALL_KINDS: readonly Kind[] = ["styles", "numbering", "settings", "fontTable", "webSettings", "theme"];
const PROBED_PACKAGE_KINDS: readonly Kind[] = ["styles", "numbering", "settings", "theme"];
const PROBED_GLOSSARY_KINDS: readonly Kind[] = ["styles", "numbering", "settings", "fontTable"];

const run = async (
    dir: string,
    name: string,
    make: () => Promise<Buffer>,
): Promise<{ wordValid: Awaited<ReturnType<typeof validate>>; strict: Awaited<ReturnType<typeof validate>> }> => {
    const file = path.join(dir, `${name}.docx`);
    await fs.writeFile(file, await make());
    return { wordValid: await validate(file, { profile: "word-valid" }), strict: await validate(file, { profile: "strict" }) };
};
const hasDuplicateFinding = (result: Awaited<ReturnType<typeof validate>>, kind?: Kind): boolean =>
    result.issues.some((issue) => (kind === undefined ? issue.code?.startsWith("word-duplicate-") : issue.code === CHECK_OF[kind]));

describe("duplicate-relationship scope (#717; sessions 49 and 50)", { timeout: 60000 }, () => {
    it("opens the minimal control clean", async () => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "control-minimal", () => buildDocx({ ...P }));
            expect(wordValid.valid).toBe(true);
            expect(hasDuplicateFinding(wordValid)).toBe(false);
            expect(hasDuplicateFinding(strict)).toBe(false);
        });
    });

    it.each(ALL_KINDS)("reports two %s relationships in the main document's sidecar", async (kind) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, `dup-${kind}-main`, () => dupInMainRels(kind));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
        });
    });

    it.each(PROBED_PACKAGE_KINDS)("does not report two %s relationships in _rels/.rels (session probe)", async (kind) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, `dup-${kind}-package-rels`, () => dupInPackageRels(kind));
            expect(hasDuplicateFinding(wordValid)).toBe(false);
            expect(wordValid.valid).toBe(true);
            expect(hasDuplicateFinding(strict)).toBe(false);
        });
    });

    it.each(["fontTable", "webSettings"] as const)(
        "does not report two %s relationships in _rels/.rels (unshown, same rule)",
        async (kind) => {
            await withTempDir(async (dir) => {
                const { wordValid, strict } = await run(dir, `dup-${kind}-package-rels`, () => dupInPackageRels(kind));
                expect(hasDuplicateFinding(wordValid)).toBe(false);
                expect(wordValid.valid).toBe(true);
                expect(hasDuplicateFinding(strict)).toBe(false);
            });
        },
    );

    it.each(ALL_KINDS)("does not report two %s relationships in a referenced header's sidecar", async (kind) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, `dup-${kind}-header-sidecar`, () => dupInHeaderRels(kind));
            expect(hasDuplicateFinding(wordValid)).toBe(false);
            expect(wordValid.valid).toBe(true);
            expect(hasDuplicateFinding(strict)).toBe(false);
        });
    });

    it.each(PROBED_GLOSSARY_KINDS)("reports two %s relationships in a reached glossary's sidecar (session probe)", async (kind) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, `dup-${kind}-glossary-sidecar`, () => dupInGlossaryRels(kind));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
        });
    });

    it.each(["webSettings", "theme"] as const)(
        "reports two %s relationships in a reached glossary's sidecar (unshown, same rule)",
        async (kind) => {
            await withTempDir(async (dir) => {
                const { wordValid, strict } = await run(dir, `dup-${kind}-glossary-sidecar`, () => dupInGlossaryRels(kind));
                expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
                expect(wordValid.valid).toBe(false);
                expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF[kind])).toBe(true);
            });
        },
    );

    it.each([1, 2])("reports the real styles relationship beside %i External styles relationship(s)", async (count) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, `styles-plus-${count}-external`, () => withExternalStyles(count));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF.styles)).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CHECK_OF.styles)).toBe(true);
        });
    });
});
