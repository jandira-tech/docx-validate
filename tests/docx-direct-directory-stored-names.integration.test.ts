/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Direct directory validators resolve stored spellings as the packed package does (#700). */
import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { canonicalPartName } from "../src/lib/content-types";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { normalizePartName } from "../src/lib/validation-context";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const IDS_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const CEX_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
const relationships = (...rels: string[]): string =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
const typedRelationship = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
const relationship = (id: string, type: string, target: string): string => typedRelationship(id, `${REL}/${type}`, target);
const paragraph = (style: string): string => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
const styles = (style: string): string =>
    MINIMAL_PARTS["word/styles.xml"]!.replace(
        "</w:styles>",
        `<w:style w:type="paragraph" w:styleId="${style}"><w:name w:val="${style}"/></w:style></w:styles>`,
    );
const GLOSSARY = `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartBody>${paragraph("GlossaryStyle")}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
const PARTS: Readonly<Record<string, string>> = {
    ...MINIMAL_PARTS,
    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${paragraph("Heading1")}<w:sectPr>`),
    "word/styles.xml": styles("Heading1"),
    "word/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
    ),
    "word/zglossary/document.xml": GLOSSARY,
    "word/zglossary/styles.xml": styles("GlossaryStyle"),
    "word/zglossary/_rels/document.xml.rels": relationships(relationship("rId1", "styles", "styles.xml")),
};

const renamed = (parts: Readonly<Record<string, string>>, from: string, to: string): Record<string, string> =>
    Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name === from ? to : name, xml]));

const writeDirectory = async (dir: string, parts: Readonly<Record<string, string>>): Promise<void> => {
    for (const [name, xml] of Object.entries(parts)) {
        const file = path.join(dir, name);
        await fs.mkdir(path.dirname(file), { recursive: true });
        await fs.writeFile(file, xml);
    }
};

const findings = (issues: readonly ValidationIssue[], codes: readonly string[]): string[] =>
    issues
        .filter((issue) => codes.includes(issue.code ?? ""))
        .map((issue) => `${issue.code}: ${canonicalPartName(normalizePartName(issue.path ?? ""))}`)
        .sort();

const MAIN_STORED_STYLES = renamed(
    {
        ...PARTS,
        "word/_rels/document.xml.rels": relationships(
            relationship("rId1", "styles", "sub/styles.xml"),
            relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
        ),
    },
    "word/styles.xml",
    "word/sub\\styles.xml",
);
const MAIN_FALLBACK = renamed(
    {
        ...MINIMAL_PARTS,
        "_rels/.rels": relationships(relationship("rId1", "officeDocument", "word/sub/document.xml")),
        "word/document.xml": PARTS["word/document.xml"]!,
        "word/styles.xml": styles("Heading1"),
        "word/_rels/document.xml.rels": relationships(),
    },
    "word/document.xml",
    "word/sub/document.xml",
);
delete MAIN_FALLBACK["word/_rels/document.xml.rels"];
const FALLBACK_STYLES = renamed(MAIN_FALLBACK, "word/styles.xml", "word/sub\\styles.xml");
const GLOSSARY_COMMENTS = {
    ...PARTS,
    "word/zglossary/_rels/document.xml.rels": relationships(
        relationship("rId1", "styles", "styles.xml"),
        relationship("rId2", "comments", "sub/comments.xml"),
    ),
    "word/zglossary/sub\\comments.xml": `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A">${paragraph("GlossaryStyle")}</w:comment></w:comments>`,
};
const SIDECAR = renamed(
    {
        ...MINIMAL_PARTS,
        "_rels/.rels": relationships(relationship("rId1", "officeDocument", "word/sub/document.xml")),
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            "<w:sectPr>",
            `<w:p><w:hyperlink xmlns:r="${REL}" r:id="rId2"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:sectPr>`,
        ),
        "word/sub/_rels/document.xml.rels": relationships(
            relationship("rId1", "styles", "../styles.xml"),
            `<Relationship Id="rId2" Type="${REL}/hyperlink" Target="https://example.com" TargetMode="External"/>`,
        ),
    },
    "word/document.xml",
    "word/sub\\document.xml",
);

const ROWS = [
    {
        name: "main related styles",
        parts: MAIN_STORED_STYLES,
        codes: ["style-reference-undefined", "latent-styles-missing"],
        expected: ["latent-styles-missing: word/sub/styles.xml"],
    },
    {
        name: "exact normalized related styles before an equivalent alias",
        parts: { ...MAIN_STORED_STYLES, "word/s%75b/styles.xml": styles("OtherStyle") },
        codes: ["style-reference-undefined", "latent-styles-missing"],
        expected: ["latent-styles-missing: word/sub/styles.xml"],
    },
    {
        name: "main related font table",
        parts: {
            ...PARTS,
            "word/_rels/document.xml.rels": relationships(
                relationship("rId1", "styles", "styles.xml"),
                relationship("rId2", "fontTable", "sub/fonts.xml"),
            ),
            "word/sub\\fonts.xml": `<w:fonts xmlns:w="${W}"/>`,
        },
        codes: ["font-table-empty"],
        expected: ["font-table-empty: word/sub/fonts.xml"],
    },
    {
        name: "main folder fallback",
        parts: FALLBACK_STYLES,
        codes: ["style-reference-undefined", "latent-styles-missing"],
        expected: ["latent-styles-missing: word/sub/styles.xml"],
    },
    {
        name: "OPC-equivalent main folder fallback",
        parts: renamed(FALLBACK_STYLES, "word/sub\\styles.xml", "word/s%75b/styles.xml"),
        codes: ["style-reference-undefined", "latent-styles-missing"],
        expected: ["latent-styles-missing: word/sub/styles.xml"],
    },
    {
        name: "glossary root",
        parts: renamed(PARTS, "word/zglossary/document.xml", "word/zglossary\\document.xml"),
        codes: ["style-reference-undefined"],
        expected: [],
    },
    { name: "glossary-only story", parts: GLOSSARY_COMMENTS, codes: ["style-reference-undefined"], expected: [] },
    {
        name: "glossary related styles",
        parts: renamed(
            { ...PARTS, "word/zglossary/document.xml": GLOSSARY.replace("GlossaryStyle", "Heading1") },
            "word/zglossary/styles.xml",
            "word/zglossary\\styles.xml",
        ),
        codes: ["style-reference-undefined"],
        expected: ["style-reference-undefined: word/zglossary/document.xml"],
    },
    {
        name: "stored document sidecar",
        parts: renamed(MINIMAL_PARTS, "word/_rels/document.xml.rels", "word\\_rels\\document.xml.rels"),
        codes: ["rels-broken", "rels-target-missing"],
        expected: [],
    },
    {
        name: "stored package sidecar",
        parts: renamed(MINIMAL_PARTS, "_rels/.rels", "_rels\\.rels"),
        codes: ["rels-broken", "rels-target-missing"],
        expected: [],
    },
    { name: "owner part sidecar", parts: SIDECAR, codes: ["rels-missing-sidecar", "rels-id-missing"], expected: [] },
];

describe("direct directory stored names (#700)", () => {
    it.each(ROWS)("$name agrees with the packed package", async ({ parts, codes, expected }) => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await writeDirectory(unpackedDir, parts);
            const file = path.join(dir, "specimen.docx");
            await fs.writeFile(file, await buildDocx(parts));
            const packed = findings((await validate(file, { profile: "word-valid" })).issues, codes);
            const direct = findings((await new DOCXSchemaValidator({ unpackedDir }).validate()).issues, codes);
            expect(packed).toEqual(expected);
            expect(direct).toEqual(packed);
        });
    });

    it.each(["comments", "ids", "extensible", "folder", "equivalent-folder"])(
        "repairs the main comment thread through %s stored names",
        async (row) => {
            await withTempDir(async (dir) => {
                const comments = `<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="0" w:author="A"><w:p w14:paraId="0000AAAA"><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>`;
                const ids = `<w16cid:commentsIds xmlns:w16cid="${W16CID}"><w16cid:commentId w16cid:paraId="0000BBBB" w16cid:durableId="0000BBBB"/></w16cid:commentsIds>`;
                const extensible = `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}"><w16cex:commentExtensible w16cex:durableId="0000BBBB"/></w16cex:commentsExtensible>`;
                const fallback = row === "folder" || row === "equivalent-folder";
                const parts = {
                    ...MINIMAL_PARTS,
                    "word/_rels/document.xml.rels": relationships(
                        relationship("rId1", "styles", "styles.xml"),
                        relationship("rId2", "comments", "sub/comments.xml"),
                        ...(fallback
                            ? []
                            : [
                                  typedRelationship("rId3", IDS_TYPE, "sub/ids.xml"),
                                  typedRelationship("rId4", CEX_TYPE, "sub/extensible.xml"),
                              ]),
                    ),
                    [row === "comments" || fallback ? "word/sub\\comments.xml" : "word/sub/comments.xml"]: comments,
                    [row === "ids"
                        ? "word/sub\\ids.xml"
                        : row === "equivalent-folder"
                          ? "WORD/s%75b/commentsIds.xml"
                          : fallback
                            ? "word/sub/commentsIds.xml"
                            : "word/sub/ids.xml"]: ids,
                    [row === "extensible"
                        ? "word/sub\\extensible.xml"
                        : row === "equivalent-folder"
                          ? "WORD/s%75b/commentsExtensible.xml"
                          : fallback
                            ? "word/sub/commentsExtensible.xml"
                            : "word/sub/extensible.xml"]: extensible,
                };
                await writeDirectory(dir, parts);
                const validator = new DOCXSchemaValidator({ unpackedDir: dir });
                expect(await validator.repairCommentThreading()).toBe(3);
                for (const [name, xml] of Object.entries(parts)) {
                    if (xml !== ids && xml !== extensible) continue;
                    const repaired = await fs.readFile(path.join(dir, name), "utf-8");
                    expect(repaired).toContain("0000AAAA");
                    expect(repaired).not.toContain("0000BBBB");
                }
            });
        },
    );
});
