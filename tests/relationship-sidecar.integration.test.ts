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
import {
    barChartXml,
    chartSpecimenParts,
    CHART_DRAWING,
    equivalentTargetChartSpecimenParts,
    fallbackChartSpecimenParts,
    PICTURE_DRAWING,
    reachedChartSpecimenParts,
    referencedChartSpecimenParts,
    storyChartSpecimenParts,
    twoChartRelationshipsSpecimenParts,
} from "./support/chart-specimens";

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

// Session 35, chunks s35a/s35b (#638). These reproduce the supplied builders;
// no Word automation runs here. The comments calibration is an unobserved
// extension that checks Word-read non-story parts use the same predicate.

// Session 36, chunks s36b/s36d (#645), rebuilt from tests/support/chart-specimens:
// a chart part whose `c:externalData r:id` names nothing because the chart has
// no `.rels` beside it, related from the main document's sidecar as `rId9`
// (`…/relationships/chart`). Word 16.115 opened a related-but-unreferenced
// chart clean — inside `word/` and outside it — and offered recovery for the
// same chart an inline drawing's `c:chart r:id` references. No Word automation
// runs here.
const CHART_INSIDE = "word/charts/chart1.xml";
const CHART_OUTSIDE = "charts/chart1.xml";
const session36Specimens = [
    {
        name: "chart-word-related-no-sidecar",
        part: CHART_INSIDE,
        blocking: false,
        parts: () => chartSpecimenParts(CHART_INSIDE, "charts/chart1.xml"),
    },
    {
        name: "chart-outside-word-related-no-sidecar",
        part: CHART_OUTSIDE,
        blocking: false,
        parts: () => chartSpecimenParts(CHART_OUTSIDE, `../${CHART_OUTSIDE}`),
    },
    {
        name: "chart-word-referenced-no-sidecar",
        part: CHART_INSIDE,
        blocking: true,
        parts: () => referencedChartSpecimenParts(CHART_INSIDE, "charts/chart1.xml", true),
    },
    {
        name: "chart-outside-word-referenced-no-sidecar",
        part: CHART_OUTSIDE,
        blocking: true,
        parts: () => referencedChartSpecimenParts(CHART_OUTSIDE, `../${CHART_OUTSIDE}`, true),
    },
    {
        name: "chart-word-referenced-no-externaldata-no-sidecar",
        part: CHART_INSIDE,
        blocking: undefined,
        parts: () => referencedChartSpecimenParts(CHART_INSIDE, "charts/chart1.xml", false),
    },
    // Session 37 (s37b) showed Word both boundary shapes. A reference inside
    // an `mc:Choice` Word never selects cannot load the chart: it opened
    // clean, and word-valid warns. A chart an image relationship also reaches
    // opened clean as well (s37b, and again in both runs of s44a), and
    // word-valid warns since #647.
    {
        name: "chart-outside-word-mixed-reach-no-sidecar",
        part: CHART_OUTSIDE,
        blocking: false,
        parts: () => {
            const parts = chartSpecimenParts(CHART_OUTSIDE, `../${CHART_OUTSIDE}`);
            return {
                ...parts,
                [SIDECAR]: replaceOnce(
                    parts[SIDECAR]!,
                    "</Relationships>",
                    `<Relationship Id="rId8" Type="${R}/image" Target="../${CHART_OUTSIDE}"/></Relationships>`,
                ),
            };
        },
    },
    {
        name: "chart-word-unselected-reference-no-sidecar",
        part: CHART_INSIDE,
        blocking: false,
        parts: () => {
            const parts = referencedChartSpecimenParts(CHART_INSIDE, "charts/chart1.xml", true);
            return {
                ...parts,
                [MAIN]: replaceOnce(
                    replaceOnce(parts[MAIN]!, "<w:document ", `<w:document xmlns:foo="urn:foo" `),
                    CHART_DRAWING,
                    `<mc:AlternateContent><mc:Choice Requires="foo">${CHART_DRAWING}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
                ),
            };
        },
    },
];

// Lenient and strict expectations captured at aed11e4e, computed from that
// commit's sources independently of the fix; both profiles produced the
// identical rows for these specimens.
const session36Lenient: Record<string, { valid: boolean; issues: string[] }> = {
    "chart-word-related-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@word/charts/chart1.xml: Part 'word/charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/charts/chart1.xml",
            "error xsd-error@word/charts/chart1.xml",
        ],
    },
    "chart-outside-word-related-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@charts/chart1.xml: Part 'charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@charts/chart1.xml",
            "error xsd-error@charts/chart1.xml",
        ],
    },
    "chart-word-referenced-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@word/charts/chart1.xml: Part 'word/charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "chart-outside-word-referenced-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@charts/chart1.xml: Part 'charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "chart-word-referenced-no-externaldata-no-sidecar": {
        valid: true,
        issues: ["info latent-styles-missing@word/styles.xml"],
    },
    "chart-outside-word-mixed-reach-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@charts/chart1.xml: Part 'charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@charts/chart1.xml",
            "error xsd-error@charts/chart1.xml",
        ],
    },
    "chart-word-unselected-reference-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@word/charts/chart1.xml: Part 'word/charts/chart1.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
};

describe("chart relationship reach without a reference (#645)", { timeout: 120_000 }, () => {
    it.each(session36Specimens)("judges $name without changing strict or lenient output", async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${specimen.name}.docx`);
            const bytes = await buildDocx(specimen.parts());
            writeFileSync(file, bytes);
            for (const profile of ["lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    {
                        valid: result.valid,
                        issues: result.issues.map(
                            (finding) =>
                                `${finding.severity} ${finding.code}@${finding.path ?? ""}` +
                                (finding.code?.startsWith("rels-") ? `: ${finding.message}` : ""),
                        ),
                    },
                    profile,
                ).toEqual(session36Lenient[specimen.name]);
            }
            const full = await validate(file, { profile: "word-valid" });
            expect.soft(full.valid).toBe(specimen.blocking !== true);
            expect.soft(full.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual(
                specimen.blocking === undefined
                    ? []
                    : [
                          {
                              code: "rels-missing-sidecar",
                              path: specimen.part,
                              severity: specimen.blocking ? "error" : "warning",
                              message: `Part '${specimen.part}' uses relationship references but has no .rels sidecar file`,
                          },
                      ],
            );
            // The legacy walk that produces the finding stays outside named
            // and quick selections, so a partial run reports nothing here and
            // stays valid where the full run is.
            for (const selection of [{ quick: true }, { checks: ["word-header-footer-relationship", "word-hyperlink-relationship"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect(partial.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([]);
                expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                if (full.valid) expect(partial.valid).toBe(true);
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});

// Session 39 (s39a/s39c, #661): the chart sources #645's rule reaches beyond
// sessions 36-37. Word 16.115 opened each unreferenced shape clean and offered
// recovery for each referenced twin, so the relationship's use, not the
// relating part, decides. The helpers rebuild the stored specimens' parts.
const session39ChartSpecimens = [
    ...(["header", "footnote", "glossary"] as const).flatMap((source) => [
        { name: `chart-${source}-unreferenced`, blocking: false, parts: () => storyChartSpecimenParts(source, false) },
        { name: `chart-${source}-referenced`, blocking: true, parts: () => storyChartSpecimenParts(source, true) },
    ]),
    { name: "chart-unselected-fallback", blocking: false, parts: () => fallbackChartSpecimenParts(false) },
    { name: "chart-selected-fallback", blocking: true, parts: () => fallbackChartSpecimenParts(true) },
    { name: "chart-two-relationships-no-references", blocking: false, parts: () => twoChartRelationshipsSpecimenParts(false) },
    { name: "chart-two-relationships-one-reference", blocking: true, parts: () => twoChartRelationshipsSpecimenParts(true) },
    { name: "chart-equivalent-target-unreferenced", blocking: false, parts: () => equivalentTargetChartSpecimenParts(false) },
    { name: "chart-equivalent-target-referenced", blocking: true, parts: () => equivalentTargetChartSpecimenParts(true) },
];

describe("chart sources Word was shown in session 39 (#661)", { timeout: 120_000 }, () => {
    it.each(session39ChartSpecimens)("$name follows the Word observation", async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${specimen.name}.docx`);
            writeFileSync(file, await buildDocx(specimen.parts()));
            const missingSidecar = (severity: "error" | "warning") => ({
                code: "rels-missing-sidecar",
                path: CHART_INSIDE,
                severity,
                message: `Part '${CHART_INSIDE}' uses relationship references but has no .rels sidecar file`,
            });
            const full = await validate(file, { profile: "word-valid" });
            expect(full.valid).toBe(!specimen.blocking);
            expect(full.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([
                missingSidecar(specimen.blocking ? "error" : "warning"),
            ]);
            // Strict keeps the finding an error whichever way Word went.
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(strict.issues).toContainEqual(missingSidecar("error"));
            // The legacy missing-sidecar walk is not a registered check, so
            // partial runs omit it and keep only the #381 one-way contract (#663).
            for (const selection of [{ quick: true }, { checks: ["relationships"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect(partial.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([]);
                expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                if (full.valid) expect(partial.valid).toBe(true);
            }
        });
    });
});

// Session 44 (s44a–s44c, #647): which relationship types load a part's own
// relationships. The bare chart, with no sidecar, is reached from the main
// document's sidecar by the listed relationships, inside `word/` and outside
// it. Word 16.115 opened clean, in two runs each, every shape an image
// relationship reaches, even when a picture's `a:blip r:embed` names it, and
// every unreferenced diagramData, oleObject or hdphoto reach; it refused the
// chart a `c:chart r:id` names. The last three rows are unobserved and keep
// the reach-based verdict: a reference to a diagramData or oleObject
// relationship, and an image reach beside an unobserved type.
const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const O = "urn:schemas-microsoft-com:office:office";
const HDPHOTO = "http://schemas.microsoft.com/office/2007/relationships/hdphoto";
const reaching = (id: string, segment: string) => ({ id, type: segment === HDPHOTO ? HDPHOTO : `${R}/${segment}` });
const IMAGE = reaching("rId8", "image");
const CHART = reaching("rId9", "chart");
const DIAGRAM_DRAWING =
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Diagram 1"/><a:graphic><a:graphicData uri="${DGM}">` +
    `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rId8" r:lo="rId8" r:qs="rId8" r:cs="rId8"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const OLE_OBJECT =
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:object>' +
    `<o:OLEObject xmlns:o="${O}" Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="rId8"/></w:object></w:r></w:p>`;
// A drawing whose `c:chart r:id` names the image relationship `rId8`.
const IMAGE_CHART_DRAWING = CHART_DRAWING.replace('r:id="rId9"', 'r:id="rId8"');
const session44Specimens = [
    { name: "chart-outside-word-image-only-reach-no-sidecar", part: CHART_OUTSIDE, reach: [IMAGE], blocking: false },
    {
        name: "chart-outside-word-image-only-blip-reference-no-sidecar",
        part: CHART_OUTSIDE,
        reach: [IMAGE],
        drawing: PICTURE_DRAWING,
        blocking: false,
    },
    {
        name: "chart-outside-word-mixed-reach-blip-reference-no-sidecar",
        part: CHART_OUTSIDE,
        reach: [IMAGE, CHART],
        drawing: PICTURE_DRAWING,
        blocking: false,
    },
    {
        name: "chart-outside-word-mixed-reach-chart-reference-no-sidecar",
        part: CHART_OUTSIDE,
        reach: [IMAGE, CHART],
        drawing: CHART_DRAWING,
        blocking: true,
    },
    {
        name: "chart-outside-word-image-only-blip-reference-no-externaldata",
        part: CHART_OUTSIDE,
        reach: [IMAGE],
        drawing: PICTURE_DRAWING,
        externalData: false,
        blocking: undefined,
    },
    ...(["diagramData", "oleObject", HDPHOTO] as const).flatMap((segment) => {
        const label = segment === HDPHOTO ? "hdphoto" : segment.toLowerCase();
        return [
            {
                name: `chart-outside-word-${label}-reach-no-sidecar`,
                part: CHART_OUTSIDE,
                reach: [reaching("rId8", segment)],
                blocking: false,
            },
            { name: `chart-word-${label}-reach-no-sidecar`, part: CHART_INSIDE, reach: [reaching("rId8", segment)], blocking: false },
        ];
    }),
    { name: "chart-word-image-only-reach-no-sidecar", part: CHART_INSIDE, reach: [IMAGE], blocking: false },
    {
        name: "chart-word-image-only-blip-reference-no-sidecar",
        part: CHART_INSIDE,
        reach: [IMAGE],
        drawing: PICTURE_DRAWING,
        blocking: false,
    },
    // Unobserved: keep the reach-based verdict.
    {
        name: "chart-word-diagramdata-relids-reference",
        part: CHART_INSIDE,
        reach: [reaching("rId8", "diagramData")],
        drawing: DIAGRAM_DRAWING,
        blocking: true,
    },
    {
        name: "chart-word-oleobject-reference",
        part: CHART_INSIDE,
        reach: [reaching("rId8", "oleObject")],
        drawing: OLE_OBJECT,
        blocking: true,
    },
    { name: "chart-word-image-and-audio-reach", part: CHART_INSIDE, reach: [IMAGE, reaching("rId7", "audio")], blocking: true },
    // Refused for the reference's Type, whatever the target (sessions 45
    // s45c and 46, #705): `word-part-reference-relationship` blocks these
    // rows, and #706 takes #701's extension back out, so the missing
    // sidecar alone warns.
    {
        name: "chart-word-image-reach-chart-reference",
        part: CHART_INSIDE,
        reach: [IMAGE],
        drawing: IMAGE_CHART_DRAWING,
        blocking: true,
        referenceRefused: true,
    },
    {
        name: "chart-word-strict-image-reach-chart-reference",
        part: CHART_INSIDE,
        reach: [{ id: "rId8", type: `${RS}/image` }],
        drawing: IMAGE_CHART_DRAWING,
        blocking: true,
        referenceRefused: true,
    },
    {
        name: "chart-word-image-reach-relids-reference",
        part: CHART_INSIDE,
        reach: [IMAGE],
        drawing: DIAGRAM_DRAWING,
        blocking: true,
        referenceRefused: true,
    },
    {
        name: "chart-outside-word-image-reach-oleobject-reference",
        part: CHART_OUTSIDE,
        reach: [IMAGE],
        drawing: OLE_OBJECT,
        blocking: true,
        referenceRefused: true,
    },
    {
        name: "chart-word-image-reach-blip-and-chart-reference",
        part: CHART_INSIDE,
        reach: [IMAGE],
        drawing: `${PICTURE_DRAWING}${IMAGE_CHART_DRAWING}`,
        blocking: true,
        referenceRefused: true,
    },
];

describe("relationship types Word was shown in session 44 (#647)", { timeout: 120_000 }, () => {
    it.each(session44Specimens)("$name follows the Word observation", async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${specimen.name}.docx`);
            writeFileSync(
                file,
                await buildDocx(reachedChartSpecimenParts(specimen.part, specimen.reach, specimen.drawing, specimen.externalData)),
            );
            const missingSidecar = (severity: "error" | "warning") => ({
                code: "rels-missing-sidecar",
                path: specimen.part,
                severity,
                message: `Part '${specimen.part}' uses relationship references but has no .rels sidecar file`,
            });
            const full = await validate(file, { profile: "word-valid" });
            expect.soft(full.valid).toBe(specimen.blocking !== true);
            expect
                .soft(full.issues.filter((finding) => finding.code === "rels-missing-sidecar"))
                .toEqual(
                    specimen.blocking === undefined
                        ? []
                        : [missingSidecar(specimen.blocking && !specimen.referenceRefused ? "error" : "warning")],
                );
            // #706: a row Word refuses for the reference's Type stays invalid
            // through #705's rule; the missing sidecar alone warns.
            if (specimen.referenceRefused) {
                expect
                    .soft(full.issues.filter((finding) => finding.code === "word-part-reference-relationship"))
                    .toEqual([expect.objectContaining({ path: MAIN, severity: "error" })]);
            }
            // Strict and lenient keep the finding an error whichever way Word went.
            for (const profile of ["strict", "lenient"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(false);
                if (specimen.blocking !== undefined) expect(result.issues, profile).toContainEqual(missingSidecar("error"));
            }
            // The legacy missing-sidecar walk is not a registered check, so
            // partial runs omit it and keep only the #381 one-way contract (#663).
            for (const selection of [{ quick: true }, { checks: ["relationships"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect(partial.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([]);
                expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                if (full.valid) expect(partial.valid).toBe(true);
            }
        });
    });
});

// #706, with #702's S1 rows: the references #701 made blocking that Word
// 16.115 opened clean, in two runs each. The review of layers 140–142
// (findings 3–5) built the first four shapes on the bare chart with no
// sidecar, reached by an `image` relationship: a `w:name r:id` — the
// attribute session 35 showed Word never reads as a reference — a VML
// picture's `v:imagedata r:id`, the same picture in a selected
// `mc:Fallback`, and a linked picture's `a:blip r:link`. Session 46 (s46d)
// showed Word the VML picture and the linked picture naming an `image`
// relationship to session 36d's bar chart, with the `a:blip r:embed`
// control; session 45 (s45b/s45d, #702 S1) showed it the same picture
// naming a `diagramData` or `oleObject` relationship to the bare chart,
// under `word/` and outside it. An `image` reach never loads the part, and
// the one reference to a `diagramData`, `oleObject` or `hdphoto`
// relationship Word was shown is a picture's `a:blip r:embed`, so
// `rels-missing-sidecar` warns on every row here; the shapes Word refuses
// (`c:chart r:id`, `dgm:relIds`, `o:OLEObject r:id` naming an `image`
// relationship) are #705's `word-part-reference-relationship`, covered by
// the session 44 rows above.
const VML_PICTURE_RID =
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:pict>' +
    `<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="s1" style="width:10pt;height:10pt">` +
    `<v:imagedata r:id="rId8" o:title=""/></v:shape></w:pict></w:r></w:p>`;
const NAME_REFERENCE = '<w:p w14:paraId="00000002" w14:textId="00000002"><w:name r:id="rId8"/><w:r><w:t>x</w:t></w:r></w:p>';
const LINKED_PICTURE_DRAWING = replaceOnce(PICTURE_DRAWING, 'r:embed="rId8"', 'r:link="rId8"');
/** The `drawing` inside a selected `mc:Fallback`: an unselected `foo` Choice before it, so the Fallback is read. */
const inSelectedFallback = (parts: Record<string, string>, drawing: string): Record<string, string> => ({
    ...parts,
    [MAIN]: replaceOnce(
        replaceOnce(parts[MAIN]!, "<w:document ", '<w:document xmlns:foo="urn:foo" '),
        drawing,
        `<mc:AlternateContent><mc:Choice Requires="foo"/><mc:Fallback>${drawing}</mc:Fallback></mc:AlternateContent>`,
    ),
});
/** Session 46's bar chart (s46d) in place of the bare chart, still with no sidecar and an image reach. */
const barImageReachedChartParts = (drawing: string): Record<string, string> => ({
    ...reachedChartSpecimenParts(CHART_INSIDE, [IMAGE], drawing),
    [CHART_INSIDE]: barChartXml(true),
});
const session45And46Specimens = [
    {
        name: "chart-word-image-reach-wname-reference",
        part: CHART_INSIDE,
        parts: () => reachedChartSpecimenParts(CHART_INSIDE, [IMAGE], NAME_REFERENCE),
    },
    {
        name: "chart-word-image-reach-vml-imagedata-reference",
        part: CHART_INSIDE,
        parts: () => reachedChartSpecimenParts(CHART_INSIDE, [IMAGE], VML_PICTURE_RID),
    },
    {
        name: "chart-word-image-reach-vml-imagedata-fallback",
        part: CHART_INSIDE,
        parts: () => inSelectedFallback(reachedChartSpecimenParts(CHART_INSIDE, [IMAGE], VML_PICTURE_RID), VML_PICTURE_RID),
    },
    {
        name: "chart-word-image-reach-blip-link-reference",
        part: CHART_INSIDE,
        parts: () => reachedChartSpecimenParts(CHART_INSIDE, [IMAGE], LINKED_PICTURE_DRAWING),
    },
    { name: "bar-image-reach-vml-imagedata-rid", part: CHART_INSIDE, parts: () => barImageReachedChartParts(VML_PICTURE_RID) },
    { name: "bar-image-reach-blip-link", part: CHART_INSIDE, parts: () => barImageReachedChartParts(LINKED_PICTURE_DRAWING) },
    { name: "bar-image-reach-blip-embed", part: CHART_INSIDE, parts: () => barImageReachedChartParts(PICTURE_DRAWING) },
    {
        name: "chart-word-diagramdata-reach-blip-reference",
        part: CHART_INSIDE,
        parts: () => reachedChartSpecimenParts(CHART_INSIDE, [reaching("rId8", "diagramData")], PICTURE_DRAWING),
    },
    {
        name: "chart-word-oleobject-reach-blip-reference",
        part: CHART_INSIDE,
        parts: () => reachedChartSpecimenParts(CHART_INSIDE, [reaching("rId8", "oleObject")], PICTURE_DRAWING),
    },
    {
        name: "chart-outside-word-diagramdata-reach-blip-reference",
        part: CHART_OUTSIDE,
        parts: () => reachedChartSpecimenParts(CHART_OUTSIDE, [reaching("rId8", "diagramData")], PICTURE_DRAWING),
    },
];

describe("references Word opened clean to a reference-loaded part (#706, #702)", { timeout: 120_000 }, () => {
    it.each(session45And46Specimens)("$name leaves the missing sidecar a warning", async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${specimen.name}.docx`);
            writeFileSync(file, await buildDocx(specimen.parts()));
            const missingSidecar = (severity: "error" | "warning") => ({
                code: "rels-missing-sidecar",
                path: specimen.part,
                severity,
                message: `Part '${specimen.part}' uses relationship references but has no .rels sidecar file`,
            });
            const full = await validate(file, { profile: "word-valid" });
            expect.soft(full.valid).toBe(true);
            expect.soft(full.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([missingSidecar("warning")]);
            // Strict and lenient keep the finding an error whichever way Word went.
            for (const profile of ["strict", "lenient"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(false);
                expect(result.issues, profile).toContainEqual(missingSidecar("error"));
            }
            // The legacy missing-sidecar walk is not a registered check, so
            // partial runs omit it and keep only the #381 one-way contract (#663).
            for (const selection of [{ quick: true }, { checks: ["relationships"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect(partial.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([]);
                expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                if (full.valid) expect(partial.valid).toBe(true);
            }
        });
    });
});

const session35Main = (declarations: string, from: string, to: string): string =>
    replaceOnce(replaceOnce(MINIMAL_PARTS[MAIN]!, "<w:document ", `<w:document ${declarations} `), from, to);
const session35Override = (part: string, type: string): string =>
    replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
const session35Specimens = [
    ...[RS, R].map((namespace, index) => ({
        name: `unread-xtra-${index === 0 ? "rs" : "r"}-id-no-sidecar`,
        part: "word/xtra.xml",
        blocking: false,
        parts: () => ({
            ...MINIMAL_PARTS,
            "word/xtra.xml": `<?xml version="1.0"?><x:root xmlns:x="urn:x" xmlns:r="${namespace}"><x:e r:id="rId9"/></x:root>`,
        }),
    })),
    {
        name: "unread-app-props-rs-id-no-sidecar",
        part: "docProps/app.xml",
        blocking: false,
        parts: () => ({
            ...MINIMAL_PARTS,
            "docProps/app.xml": `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:rs="${RS}"><Application rs:id="rId9">x</Application></Properties>`,
            [CT]: session35Override("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
        }),
    },
    {
        name: "unreached-glossary-rs-id-no-sidecar",
        part: GLOSSARY,
        blocking: false,
        parts: () => ({
            ...MINIMAL_PARTS,
            [GLOSSARY]: `<?xml version="1.0"?><w:glossaryDocument xmlns:w="${WS}" xmlns:rs="${RS}"><w:docParts><w:docPart><w:docPartBody><w:sectPr><w:footerReference w:type="default" rs:id="rId9"/></w:sectPr></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`,
            [CT]: session35Override(GLOSSARY, storyContentType("document.glossary")),
        }),
    },
    {
        name: "styles-rs-id-no-sidecar",
        part: "word/styles.xml",
        blocking: false,
        parts: () => ({
            ...MINIMAL_PARTS,
            "word/styles.xml": replaceOnce(
                replaceOnce(MINIMAL_PARTS["word/styles.xml"]!, "<w:styles ", `<w:styles xmlns:rs="${RS}" `),
                "<w:name ",
                '<w:name rs:id="rId9" ',
            ),
        }),
    },
    ...[RS, R].map((namespace, index) => ({
        name: `mce-unselected-${index === 0 ? "rs" : "r"}-embed-no-sidecar`,
        part: MAIN,
        blocking: false,
        parts: () =>
            withoutRels(
                {
                    ...MINIMAL_PARTS,
                    [MAIN]: session35Main(
                        `xmlns:r="${namespace}" xmlns:foo="urn:foo" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`,
                        "<w:sectPr>",
                        '<mc:AlternateContent><mc:Choice Requires="foo"><w:p><w:r><w:drawing><a:blip r:embed="rId9"/></w:drawing></w:r></w:p></mc:Choice><mc:Fallback/></mc:AlternateContent><w:sectPr>',
                    ),
                },
                SIDECAR,
            ),
    })),
    {
        name: "sectprchange-rs-id-no-sidecar",
        part: MAIN,
        blocking: true,
        parts: () =>
            withoutRels(
                {
                    ...MINIMAL_PARTS,
                    [MAIN]: session35Main(
                        `xmlns:rs="${RS}"`,
                        "</w:sectPr>",
                        '<w:sectPrChange w:id="5" w:author="a" w:date="2026-01-01T00:00:00Z"><w:sectPr><w:footerReference w:type="default" rs:id="rId9"/></w:sectPr></w:sectPrChange></w:sectPr>',
                    ),
                },
                SIDECAR,
            ),
    },
    {
        name: "read-comments-reference-no-sidecar",
        part: "word/comments.xml",
        blocking: true,
        parts: () => ({
            ...MINIMAL_PARTS,
            "word/comments.xml": `<w:comments xmlns:w="${W}" xmlns:r="${R}"><w:comment w:id="0" w:author="a"><w:p><w:hyperlink r:id="rId9"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p></w:comment></w:comments>`,
            [MAIN]: MINIMAL_PARTS[MAIN]!.replace(
                "<w:r><w:t>Hello",
                '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t>Hello',
            ),
            [SIDECAR]: MINIMAL_PARTS[SIDECAR]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdC" Type="${R}/comments" Target="comments.xml"/></Relationships>`,
            ),
            [CT]: session35Override("word/comments.xml", storyContentType("comments")),
        }),
    },
];

// Lenient expectations from the original RED-time results at 404b9804.
const session35Lenient: Record<string, { valid: boolean; issues: string[] }> = {
    "mce-unselected-r-embed-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/styles.xml: Unreferenced file: word/styles.xml",
            "error rels-missing-sidecar@word/document.xml: Part 'word/document.xml' uses relationship references but has no .rels sidecar file",
            "warning paraid-missing-element@word/document.xml",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "mce-unselected-rs-embed-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/styles.xml: Unreferenced file: word/styles.xml",
            "error rels-missing-sidecar@word/document.xml: Part 'word/document.xml' uses relationship references but has no .rels sidecar file",
            "warning paraid-missing-element@word/document.xml",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "read-comments-reference-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@word/comments.xml: Part 'word/comments.xml' uses relationship references but has no .rels sidecar file",
            "warning paraid-missing-element@word/comments.xml",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "sectprchange-rs-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/styles.xml: Unreferenced file: word/styles.xml",
            "error rels-missing-sidecar@word/document.xml: Part 'word/document.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/document.xml",
            "error xsd-error@word/document.xml",
        ],
    },
    "styles-rs-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-missing-sidecar@word/styles.xml: Part 'word/styles.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/styles.xml",
            "error xsd-error@word/styles.xml",
        ],
    },
    "unreached-glossary-rs-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/glossary/document.xml: Unreferenced file: word/glossary/document.xml",
            "error rels-missing-sidecar@word/glossary/document.xml: Part 'word/glossary/document.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "info xsd-strict-skipped@",
        ],
    },
    "unread-app-props-rs-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@docProps/app.xml: Unreferenced file: docProps/app.xml",
            "error rels-missing-sidecar@docProps/app.xml: Part 'docProps/app.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@docProps/app.xml",
            "error xsd-error@docProps/app.xml",
        ],
    },
    "unread-xtra-r-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/xtra.xml: Unreferenced file: word/xtra.xml",
            "error rels-missing-sidecar@word/xtra.xml: Part 'word/xtra.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/xtra.xml",
            "error xsd-error@word/xtra.xml",
        ],
    },
    "unread-xtra-rs-id-no-sidecar": {
        valid: false,
        issues: [
            "error rels-unreferenced@word/xtra.xml: Unreferenced file: word/xtra.xml",
            "error rels-missing-sidecar@word/xtra.xml: Part 'word/xtra.xml' uses relationship references but has no .rels sidecar file",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/xtra.xml",
            "error xsd-error@word/xtra.xml",
        ],
    },
};

describe("Word-read processed missing-sidecar references (#638)", { timeout: 120_000 }, () => {
    it.each(session35Specimens)("judges $name without changing strict or lenient output", async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${specimen.name}.docx`);
            const bytes = await buildDocx(specimen.parts());
            writeFileSync(file, bytes);
            const lenient = await validate(file, { profile: "lenient" });
            // Pin findings rather than the registry-dependent result metadata.
            expect({
                valid: lenient.valid,
                issues: lenient.issues.map(
                    (finding) =>
                        `${finding.severity} ${finding.code}@${finding.path ?? ""}` +
                        (finding.code?.startsWith("rels-") ? `: ${finding.message}` : ""),
                ),
            }).toEqual(session35Lenient[specimen.name]);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            const expected = {
                code: "rels-missing-sidecar",
                path: specimen.part,
                severity: "error",
                message: `Part '${specimen.part}' uses relationship references but has no .rels sidecar file`,
            };
            expect(strict.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([expected]);
            const full = await validate(file, { profile: "word-valid" });
            expect.soft(full.valid).toBe(!specimen.blocking);
            expect
                .soft(full.issues.filter((finding) => finding.code === "rels-missing-sidecar"))
                .toEqual([{ ...expected, severity: specimen.blocking ? "error" : "warning" }]);
            for (const selection of [{ quick: true }, { checks: ["word-header-footer-relationship", "word-hyperlink-relationship"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect(partial.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([]);
                expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                if (full.valid) expect(partial.valid).toBe(true);
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});

// Unobserved shapes retain the pre-#638 blocking decision without Word evidence.
describe("conservative missing-sidecar classification (#638 rework)", { timeout: 120_000 }, () => {
    it.each([
        [
            "chart",
            'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
            '<w:r><w:drawing><c:chart r:id="rId9"/></w:drawing></w:r>',
        ],
        [
            "svg",
            'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"',
            '<w:r><w:drawing><asvg:svgBlip r:embed="rId9"/></w:drawing></w:r>',
        ],
        ["unknown", 'xmlns:x="urn:x"', '<x:e r:id="rId9"/>'],
        ["subdoc", "", '<w:subDoc r:id="rId9"/>'],
    ])("keeps %s blocking in the main story", async (_name, declaration, content) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "reference.docx");
            writeFileSync(
                file,
                await buildDocx(
                    withoutRels(
                        {
                            ...MINIMAL_PARTS,
                            [MAIN]: session35Main(`xmlns:r="${R}" ${declaration}`, "<w:sectPr>", `${content}<w:sectPr>`),
                        },
                        SIDECAR,
                    ),
                ),
            );
            // Computed against 404b9804 validator copies, independently of the fix.
            const lenient = await validate(file, { profile: "lenient" });
            expect({
                valid: lenient.valid,
                issues: lenient.issues.map(
                    (finding) =>
                        `${finding.severity} ${finding.code}@${finding.path ?? ""}` +
                        (finding.code?.startsWith("rels-") ? `: ${finding.message}` : ""),
                ),
            }).toEqual({
                valid: false,
                issues: [
                    "error rels-unreferenced@word/styles.xml: Unreferenced file: word/styles.xml",
                    "error rels-missing-sidecar@word/document.xml: Part 'word/document.xml' uses relationship references but has no .rels sidecar file",
                    "info latent-styles-missing@word/styles.xml",
                    "error xsd-summary@word/document.xml",
                    "error xsd-error@word/document.xml",
                ],
            });
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(strict.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([
                {
                    severity: "error",
                    code: "rels-missing-sidecar",
                    path: MAIN,
                    message: `Part '${MAIN}' uses relationship references but has no .rels sidecar file`,
                },
            ]);
            const result = await validate(file, { profile: "word-valid" });
            expect.soft(result.valid).toBe(false);
            expect(result.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([
                {
                    severity: "error",
                    code: "rels-missing-sidecar",
                    path: MAIN,
                    message: `Part '${MAIN}' uses relationship references but has no .rels sidecar file`,
                },
            ]);
        });
    });

    // A part an unobserved Office relationship type names outside word/ was
    // never shown to Word; it blocked before #638 and keeps blocking. The same
    // part named by no relationship is unread, as s35a's word/xtra.xml and
    // docProps/app.xml are, and warns. The chart type is the one exception:
    // session 36 observed it (#645), so a related-but-unreferenced chart warns
    // alongside the unrelated shape — the rows below pin both.
    const CHART = "charts/chart1.xml";
    const SIDECAR_MESSAGE = `Part '${CHART}' uses relationship references but has no .rels sidecar file`;
    it.each([
        { related: true, blocking: false },
        { related: false, blocking: false },
    ])("judges a chart outside word/ with no sidecar, related: $related", async ({ related, blocking }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "chart.docx");
            writeFileSync(file, await buildDocx(chartSpecimenParts(CHART, related ? `../${CHART}` : undefined)));
            // Computed at bcea63f6 (404b9804's validators), independently of the fix.
            const rows = [
                ...(related ? [] : [`error rels-unreferenced@${CHART}: Unreferenced file: ${CHART}`]),
                `error rels-missing-sidecar@${CHART}: ${SIDECAR_MESSAGE}`,
                "info latent-styles-missing@word/styles.xml",
                `error xsd-summary@${CHART}`,
                `error xsd-error@${CHART}`,
            ];
            for (const profile of ["lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    {
                        valid: result.valid,
                        issues: result.issues.map(
                            (finding) =>
                                `${finding.severity} ${finding.code}@${finding.path ?? ""}` +
                                (finding.code?.startsWith("rels-") ? `: ${finding.message}` : ""),
                        ),
                    },
                    profile,
                ).toEqual({ valid: false, issues: rows });
            }
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(!blocking);
            expect(result.issues.filter((finding) => finding.code === "rels-missing-sidecar")).toEqual([
                { severity: blocking ? "error" : "warning", code: "rels-missing-sidecar", path: CHART, message: SIDECAR_MESSAGE },
            ]);
        });
    });
});
