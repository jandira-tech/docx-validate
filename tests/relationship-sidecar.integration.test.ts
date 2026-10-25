/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { STORY_FOOTER_XML, glossaryStoryReferenceParts, storyContentType } from "./support/story-content-type-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const ESCAPED_NAME = "opc-part-name-escaped-unreserved";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const CT = "[Content_Types].xml";
const HEADER1 = "word/header1.xml";
const HEADER1_SIDECAR = "word/_rels/header1.xml.rels";
const GLOSSARY = "word/glossary/document.xml";
const GLOSSARY_SIDECAR = "word/glossary/_rels/document.xml.rels";
/** `text` with the one occurrence of `from` replaced; throws when `from` is absent, so a shape cannot silently stay unchanged. */
const replaceOnce = (text: string, from: string, to: string): string => {
    if (!text.includes(from)) throw new Error(`missing ${from}`);
    return text.replace(from, to);
};
const TEXT = "<w:r><w:t>Hello</w:t></w:r>";
const docH = MINIMAL_PARTS[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
    TEXT,
    `<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`,
);
const relsH = MINIMAL_PARTS[SIDECAR]!.replace(
    "</Relationships>",
    `<Relationship Id="rId2" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
);
const hyperlinkParts = (extension: string): Record<string, string> => ({
    ...Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== SIDECAR)),
    [MAIN]: docH,
    [`word/_rels/document.xml.${extension}`]: relsH,
});
const picture = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const imageParts = (extension: string): Record<string, string> => ({
    ...hyperlinkParts(extension),
    [MAIN]: docH.replace(`<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`, picture),
    [`word/_rels/document.xml.${extension}`]: relsH.replace(
        `Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"`,
        `Type="${R}/image" Target="media/image.svg"`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        '<Default Extension="svg" ContentType="image/svg+xml"/></Types>',
    ),
    "word/media/image.svg": '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1"/></svg>',
});
const errors = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`);

describe("relationship references without a sidecar (#624)", () => {
    // Session 33, s33a/s33c: rewrite only document/styles WML, preserving
    // MINIMAL_PARTS' package relationships and content types as observed.
    const noSidecar = (strict: boolean, reference = ""): Record<string, string> => {
        const parts = Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== SIDECAR));
        if (strict) {
            for (const part of [MAIN, "word/styles.xml"]) parts[part] = parts[part]!.replaceAll(W, WS);
        }
        if (reference) {
            parts[MAIN] = parts[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${strict ? RS : R}" `).replace(
                "<w:sectPr>",
                `<w:sectPr>${reference}`,
            );
        }
        return parts;
    };
    const footerReference = '<w:footerReference w:type="default" r:id="rId9"/>';

    it.each([
        { strict: true, label: "s33a ISO Strict" },
        { strict: false, label: "session 32 Transitional twin" },
    ])("rejects $label with exactly one missing-sidecar finding", async ({ strict }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "footer-reference.docx");
            writeFileSync(file, await buildDocx(noSidecar(strict, footerReference)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errors(result.issues)).toEqual([`rels-missing-sidecar@${MAIN}`]);
            expect(result.issues.filter((issue) => issue.code === "rels-missing-sidecar")).toEqual([
                expect.objectContaining({ path: MAIN, severity: "error" }),
            ]);
            // No sidecar belongs to the legacy walk alone, including s33a.
            expect(result.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
        });
    });

    it("keeps s33c ISO Strict without a sidecar or reference valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "no-reference.docx");
            writeFileSync(file, await buildDocx(noSidecar(true)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "rels-missing-sidecar")).toEqual([]);
        });
    });

    // Same-footing namespace extension, not Word-observed: a Strict embed/link
    // alone on a descendant of a Transitional part still needs a sidecar.
    it.each(["embed", "link"] as const)("reports an ISO Strict r:%s in a Transitional part", async (attribute) => {
        await withTempDir(async (dir) => {
            mkdirSync(path.join(dir, "word"));
            writeFileSync(
                path.join(dir, MAIN),
                `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:drawing>` +
                    `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rs="${RS}" rs:${attribute}="rId9"/>` +
                    "</w:drawing></w:r></w:p></w:body></w:document>",
            );
            const result = await new BaseSchemaValidator({ unpackedDir: dir }).validateAllRelationshipIds();
            expect(result.valid).toBe(false);
            expect(result.issues).toEqual([expect.objectContaining({ code: "rels-missing-sidecar", path: MAIN, severity: "error" })]);
        });
    });
});

// Session 33's s33b shapes (#625), built as the evidence builder built them
// (docs/plans/evidence/word-probes-2026-10-24.md, session 33's builder).
const FOOTER_RID9 = '<w:footerReference w:type="default" r:id="rId9"/>';
/** The main document's body sectPr gets the reference, with `xmlns:r` declared. */
const withFooterReference = (parts: Record<string, string>): Record<string, string> => ({
    ...parts,
    [MAIN]: replaceOnce(replaceOnce(parts[MAIN]!, "<w:document ", `<w:document xmlns:r="${R}" `), "<w:sectPr>", `<w:sectPr>${FOOTER_RID9}`),
});
/** `parts` without the relationships part `rels`. */
const withoutRels = (parts: Record<string, string>, rels: string): Record<string, string> => {
    const out: Record<string, string> = { ...parts };
    if (out[rels] === undefined) throw new Error(`missing ${rels}`);
    delete out[rels];
    return out;
};
/** Review BE S2 (#625): the main document stored as word/header1.xml, with its `.rels`, Override and package relationship renamed with it. */
const mainAsHeader1 = (parts: Record<string, string>): Record<string, string> => {
    const out: Record<string, string> = { ...parts };
    out[HEADER1] = out[MAIN]!;
    delete out[MAIN];
    if (out[SIDECAR] !== undefined) {
        out[HEADER1_SIDECAR] = out[SIDECAR]!;
        delete out[SIDECAR];
    }
    out[CT] = replaceOnce(out[CT]!, 'PartName="/word/document.xml"', 'PartName="/word/header1.xml"');
    out["_rels/.rels"] = replaceOnce(out["_rels/.rels"]!, 'Target="word/document.xml"', 'Target="word/header1.xml"');
    return out;
};
/** Session 26's clean reached-glossary footer shape. */
const glossary = (): Record<string, string> =>
    glossaryStoryReferenceParts("footerReference", "footer", "footer9.xml", STORY_FOOTER_XML, storyContentType("footer"));
/** Review BE S2 (#625): the reached glossary stored as word/header1.xml, with its `.rels` (Target adjusted), Override and relationship renamed with it. */
const glossaryAsHeader1 = (parts: Record<string, string>): Record<string, string> => {
    const out: Record<string, string> = { ...parts };
    out[HEADER1] = out[GLOSSARY]!;
    delete out[GLOSSARY];
    if (out[GLOSSARY_SIDECAR] !== undefined) {
        out[HEADER1_SIDECAR] = replaceOnce(out[GLOSSARY_SIDECAR]!, 'Target="footer9.xml"', 'Target="glossary/footer9.xml"');
        delete out[GLOSSARY_SIDECAR];
    }
    out[CT] = replaceOnce(out[CT]!, 'PartName="/word/glossary/document.xml"', 'PartName="/word/header1.xml"');
    out[SIDECAR] = replaceOnce(out[SIDECAR]!, 'Target="glossary/document.xml"', 'Target="header1.xml"');
    return out;
};
/** A genuine header at word/header1.xml: the main document's sectPr references it, and its own picture names `rId2` with no sidecar beside it. */
const genuineHeaderNoSidecar = (): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [MAIN]: MINIMAL_PARTS[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
        "<w:sectPr>",
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>',
    ),
    [HEADER1]: `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}"><w:p>${picture}</w:p></w:hdr>`,
    [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace(
        "</Relationships>",
        `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/></Relationships>`,
    ),
    [CT]: MINIMAL_PARTS[CT]!.replace(
        "</Types>",
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
});
interface StorySpecimen {
    name: string;
    parts: () => Record<string, string>;
    wordValid: boolean;
    /** The single expected `rels-missing-sidecar` finding and its word-valid severity, if any. */
    sidecar?: { path: string; severity: "error" | "warning" };
}
const STORY_SPECIMENS: readonly StorySpecimen[] = [
    { name: "header1-main-control", parts: () => mainAsHeader1({ ...MINIMAL_PARTS }), wordValid: true },
    {
        name: "header1-main-footerref-no-sidecar",
        parts: () => withoutRels(mainAsHeader1(withFooterReference({ ...MINIMAL_PARTS })), HEADER1_SIDECAR),
        wordValid: false,
        sidecar: { path: HEADER1, severity: "error" },
    },
    { name: "header1-glossary-control", parts: () => glossaryAsHeader1(glossary()), wordValid: true },
    {
        name: "header1-glossary-footerref-no-sidecar",
        parts: () => withoutRels(glossaryAsHeader1(glossary()), HEADER1_SIDECAR),
        wordValid: false,
        sidecar: { path: HEADER1, severity: "error" },
    },
    {
        name: "glossary-footerref-no-sidecar",
        parts: () => withoutRels(glossary(), GLOSSARY_SIDECAR),
        wordValid: false,
        sidecar: { path: GLOSSARY, severity: "error" },
    },
    {
        name: "genuine-header-no-sidecar",
        parts: genuineHeaderNoSidecar,
        wordValid: true,
        sidecar: { path: HEADER1, severity: "warning" },
    },
];

describe.each(["packed", "directory"] as const)("OPC sidecar lookup, %s (#402)", { timeout: 120_000 }, (input) => {
    // Filesystem and ZIP operations are integration-only. Snapshot every file,
    // including stored names, to detect accidental writes without auto-repair.
    const snapshot = (target: string): unknown =>
        input === "packed"
            ? readFileSync(target)
            : readdirSync(target, { recursive: true, withFileTypes: true })
                  .filter((entry) => entry.isFile())
                  .map((entry) => {
                      const file = path.join(entry.parentPath, entry.name);
                      return [path.relative(target, file), readFileSync(file)];
                  });
    const writePackage = async (dir: string, label: string, parts: Record<string, string>): Promise<string> => {
        const target = path.join(dir, `${label}.docx`);
        if (input === "packed") writeFileSync(target, await buildDocx(parts));
        else
            for (const [name, text] of Object.entries(parts)) {
                const file = path.join(target, name);
                mkdirSync(path.dirname(file), { recursive: true });
                writeFileSync(file, text);
            }
        return target;
    };

    it("resolves the hyperlink triple in full, named and quick runs without changing input", async () => {
        await withTempDir(async (dir) => {
            let strictControl: ValidationIssue[] = [];
            for (const [index, extension] of ["rels", "RELS", "%72els"].entries()) {
                const file = await writePackage(dir, `variant-${index}`, hyperlinkParts(extension));
                const before = snapshot(file);
                const full = await validate(file, { profile: "word-valid" });
                const strict = await validate(file, { profile: "strict" });
                for (const result of [full, strict]) {
                    expect.soft(result.valid, extension).toBe(extension !== "%72els");
                    expect
                        .soft(
                            result.issues.filter((issue) =>
                                ["rels-missing-sidecar", "word-hyperlink-relationship"].includes(issue.code ?? ""),
                            ),
                            extension,
                        )
                        .toEqual([]);
                    if (extension === "%72els") {
                        expect.soft(errors(result.issues), extension).toEqual([`${ESCAPED_NAME}@word/_rels/document.xml.%72els`]);
                    }
                }
                if (extension === "rels") strictControl = strict.issues;
                if (extension === "RELS") {
                    expect.soft(strict.issues).toEqual(
                        strictControl.map((issue) => ({
                            ...issue,
                            path: issue.path?.replace(SIDECAR, "word/_rels/document.xml.RELS"),
                            message: issue.message.replaceAll(SIDECAR, "word/_rels/document.xml.RELS"),
                        })),
                    );
                }
                for (const selection of [
                    { checks: ["word-hyperlink-relationship"] },
                    { checks: ["word-hyperlink-relationship", ESCAPED_NAME] },
                    { quick: true },
                ]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect
                        .soft(
                            partial.issues.filter((issue) => issue.code === "word-hyperlink-relationship"),
                            extension,
                        )
                        .toEqual([]);
                    if (full.valid) expect.soft(partial.valid, extension).toBe(true);
                    expect.soft(errors(full.issues), extension).toEqual(expect.arrayContaining(errors(partial.issues)));
                    if (extension === "%72els") {
                        if ("quick" in selection || selection.checks?.includes(ESCAPED_NAME)) {
                            expect.soft(partial.valid).toBe(false);
                            expect.soft(errors(partial.issues)).toContain(`${ESCAPED_NAME}@word/_rels/document.xml.%72els`);
                        } else {
                            expect.soft(partial.valid).toBe(true);
                            expect.soft(partial.metadata?.omittedChecks).toContain(ESCAPED_NAME);
                        }
                    }
                }
                expect(snapshot(file)).toEqual(before);
            }
        });
    });

    it("resolves image sidecars and preserves missing-sidecar controls", async () => {
        await withTempDir(async (dir) => {
            for (const [index, extension] of ["rels", "RELS", "%72els"].entries()) {
                const file = await writePackage(dir, `variant-${index}`, imageParts(extension));
                const before = snapshot(file);
                for (const profile of ["word-valid", "strict"] as const) {
                    const full = await validate(file, { profile });
                    expect
                        .soft(
                            full.issues.filter((issue) => issue.code === "rels-missing-sidecar"),
                            `${profile}/${extension}`,
                        )
                        .toEqual([]);
                }
                expect(snapshot(file)).toEqual(before);
            }
            // Negative controls retain the existing document error and the
            // Word-tolerated header/footer warning, in full and partial runs.
            for (const story of ["document", "header1", "footer1"]) {
                const part = `word/${story}.xml`;
                const root = story === "header1" ? "hdr" : "ftr";
                const parts =
                    story === "document"
                        ? Object.fromEntries(Object.entries(hyperlinkParts("rels")).filter(([name]) => name !== SIDECAR))
                        : {
                              ...MINIMAL_PARTS,
                              [part]: docH
                                  .replace(`<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`, picture)
                                  .replaceAll("w:document", `w:${root}`)
                                  .replace("<w:body>", "")
                                  .replace("</w:body>", "")
                                  .replace(/<w:sectPr>.*<\/w:sectPr>/, ""),
                          };
                const file = await writePackage(dir, story, parts);
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues).toContainEqual(
                    expect.objectContaining({
                        code: "rels-missing-sidecar",
                        path: part,
                        severity: story === "document" ? "error" : "warning",
                    }),
                );
                expect(result.valid).toBe(story !== "document");
                for (const selection of [{ checks: ["word-hyperlink-relationship"] }, { quick: true }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    if (result.valid) expect(partial.valid).toBe(true);
                    expect(errors(result.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                }
            }
        });
    });

    it("keeps a story root's missing sidecar blocking under a header's name (#625)", async () => {
        // Session 33 (s33b): Word judges a story root by the relationship that
        // reaches it, not by its file name. A main document, and a reached
        // glossary, each stored as word/header1.xml with no .rels beside their
        // story reference get the recovery offer, and the renamed controls
        // open clean; the ordinary-name glossary calibration already blocks,
        // and a genuine header keeps its Word-tolerated warning.
        await withTempDir(async (dir) => {
            for (const specimen of STORY_SPECIMENS) {
                const file = await writePackage(dir, specimen.name, specimen.parts());
                const before = snapshot(file);
                const full = await validate(file, { profile: "word-valid" });
                expect.soft(full.valid, specimen.name).toBe(specimen.wordValid);
                expect
                    .soft(
                        full.issues.filter((issue) => issue.code === "rels-missing-sidecar"),
                        specimen.name,
                    )
                    .toEqual(
                        specimen.sidecar === undefined
                            ? []
                            : [expect.objectContaining({ code: "rels-missing-sidecar", ...specimen.sidecar })],
                    );
                // The legacy ID walk that produces the finding stays outside
                // named and quick selections, so a partial run agrees by
                // reporting nothing here either.
                for (const selection of [{ checks: ["word-hyperlink-relationship"] }, { quick: true }] as const) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect
                        .soft(
                            partial.issues.filter((issue) => issue.code === "rels-missing-sidecar"),
                            `${specimen.name}: ${JSON.stringify(selection)}`,
                        )
                        .toEqual([]);
                    if (full.valid) expect.soft(partial.valid, specimen.name).toBe(true);
                }
                expect(snapshot(file)).toEqual(before);
            }
        });
    });
});
