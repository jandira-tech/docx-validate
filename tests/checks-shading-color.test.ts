/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for the shading colour value-space check
 * (`src/scripts/office/checks/shading-color.ts`, Task 8 family 8c,
 * lane issues #210, #215, #220, #221, #225, #229, #238, #242).
 *
 * Every defect case is paired with a nearby legal control that stays
 * unreported. Matching is namespace-based, so the strict-namespace case uses
 * the purl.oclc.org vocabulary with no `w:` prefix at all.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext, type PartContent } from "../src/lib/validation-context";
import { wordShadingColorCheck } from "../src/scripts/office/checks/shading-color";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const documentXml = (body: string, ns: string = W): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`;

const fromParts = (parts: Readonly<Record<string, PartContent>>): ValidationContext => ValidationContext.fromEntries(parts);

const findings = async (parts: Readonly<Record<string, PartContent>>): Promise<readonly ValidationIssue[]> =>
    wordShadingColorCheck.run(fromParts(parts));

describe("shading-color check", () => {
    it('collapses whitespace around hexBinary colours as the schema does, but not around "auto" or theme names', async () => {
        // ST_HexColorRGB and ST_UcharHexNumber restrict xsd:hexBinary, whose
        // whiteSpace facet is fixed to collapse; ST_HexColorAuto and
        // ST_ThemeColor restrict xsd:string, which preserves whitespace.
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:p><w:pPr><w:shd w:val="clear" w:fill=" DDDDDD " w:color="&#9;0a0B0c&#10;" w:themeFillTint=" 66"/></w:pPr></w:p>` +
                    `<w:p><w:pPr><w:shd w:val="clear" w:fill=" auto" w:themeFill="accent1 "/></w:pPr></w:p>`,
            ),
        });
        expect(issues.map((issue) => issue.message.match(/carries (w:\w+)=/)?.[1])).toEqual(["w:fill", "w:themeFill"]);
    });

    it("reports every w:shd colour attribute outside the declared value space, once per occurrence", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                // Two identical empty fills must not collapse into one finding.
                `<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
                    `<w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill=""/></w:tcPr><w:p/></w:tc>` +
                    `<w:tc><w:tcPr><w:shd w:val="clear" w:fill=""/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>` +
                    `<w:p><w:pPr><w:shd w:val="clear" w:color="DDDDD"/></w:pPr></w:p>` +
                    `<w:p><w:r><w:rPr><w:shd w:val="clear" w:themeColor="sparkle"/></w:rPr><w:t>x</w:t></w:r></w:p>` +
                    `<w:p><w:r><w:rPr><w:shd w:val="clear" w:themeTint="0"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
            ),
        });
        expect(issues).toHaveLength(5);
        for (const issue of issues) {
            expect(issue.severity).toBe("warning");
            expect(issue.code).toBe("word-shading-color");
            expect(issue.path).toBe("word/document.xml");
        }
        expect(issues[0]!.message).toContain("w:fill");
        expect(issues[0]!.message).toContain(`""`);
        expect(issues[2]!.message).toContain("w:color");
        expect(issues[2]!.message).toContain(`"DDDDD"`);
        expect(issues[3]!.message).toContain("w:themeColor");
        expect(issues[3]!.message).toContain(`"sparkle"`);
        expect(issues[4]!.message).toContain("w:themeTint");
        expect(issues[4]!.message).toContain(`"0"`);
    });

    it("keeps every nearby legal shading control unreported", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:p><w:pPr><w:shd w:val="clear" w:fill="DDDDDD"/></w:pPr></w:p>` + // 6-hex RGB
                    `<w:p><w:pPr><w:shd w:val="clear" w:fill="dddddd"/></w:pPr></w:p>` + // lowercase hex is still hexBinary
                    `<w:p><w:pPr><w:shd w:val="clear"/></w:pPr></w:p>` + // optional attribute omitted
                    `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F7F7F7"/></w:pPr></w:p>` + // ST_HexColorAuto
                    `<w:p><w:pPr><w:shd w:val="clear" w:themeFill="accent1" w:themeFillTint="66" w:themeFillShade="B2"/></w:pPr></w:p>` +
                    `<w:p><w:pPr><w:shd w:val="clear" w:themeColor="background1" w:themeShade="9C"/></w:pPr></w:p>` +
                    `<w:p><w:pPr><w:shd w:val="bogus-pattern" w:fill="auto"/></w:pPr></w:p>` + // w:val is not colour-typed
                    // Attributes outside the element's namespace are not the
                    // schema-declared ones, whatever their value: a w14: colour
                    // and an unprefixed colour are extension/qualification
                    // concerns, not colour value-space defects.
                    `<w:p><w:pPr><w:shd xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w:val="clear" w:fill="DDDDDD" w14:fill=""/></w:pPr></w:p>` +
                    `<w:p><w:pPr><w:shd w:val="clear" w:fill="DDDDDD" fill=""/></w:pPr></w:p>`,
            ),
            // Default-namespace document: the shd element still matches by
            // namespace, and its unprefixed attributes stay out of scope.
            "word/footer1.xml":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
                `<document xmlns="${W}"><ftr><p><pPr><shd val="clear" fill="auto"/></pPr></p></ftr></document>`,
        });
        expect(issues).toEqual([]);
    });

    it("scans headers, footers, footnotes, endnotes, comments, and styles.xml", async () => {
        const defect = `<w:shd w:val="clear" w:fill=""/>`;
        const parts: Record<string, PartContent> = {
            "word/document.xml": documentXml(`<w:p><w:r><w:rPr>${defect}</w:rPr><w:t>x</w:t></w:r></w:p>`),
            "word/header1.xml": documentXml(`<w:p><w:pPr>${defect}</w:pPr></w:p>`),
            "word/footer1.xml": documentXml(`<w:p><w:pPr>${defect}</w:pPr></w:p>`),
            "word/footnotes.xml": documentXml(`<w:footnote w:id="1"><w:p><w:pPr>${defect}</w:pPr></w:p></w:footnote>`),
            "word/endnotes.xml": documentXml(`<w:endnote w:id="1"><w:p><w:pPr>${defect}</w:pPr></w:p></w:endnote>`),
            "word/comments.xml": documentXml(`<w:comment w:id="1"><w:p><w:pPr>${defect}</w:pPr></w:p></w:comment>`),
            "word/styles.xml":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
                `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="P"><w:pPr>${defect}</w:pPr></w:style></w:styles>`,
            // numbering.xml is scanned with the story parts: CT_Lvl's
            // pPr/rPr admit CT_Shd (issue #303), so the defect is reported
            // there like anywhere else.
            "word/numbering.xml": documentXml(`<w:p><w:pPr>${defect}</w:pPr></w:p>`),
        };
        const paths = (await findings(parts)).map((issue) => issue.path);
        expect(paths).toEqual([
            "word/comments.xml",
            "word/document.xml",
            "word/endnotes.xml",
            "word/footer1.xml",
            "word/footnotes.xml",
            "word/header1.xml",
            "word/numbering.xml",
            "word/styles.xml",
        ]);
    });

    it("matches the strict WordprocessingML namespace whatever prefix carries it", async () => {
        const strictDoc =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<word:document xmlns:word="${STRICT}"><word:body><word:p><word:pPr>` +
            `<word:shd word:val="clear" word:fill=""/></word:pPr></word:p></word:body></word:document>`;
        const issues = await findings({ "word/document.xml": strictDoc });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.code).toBe("word-shading-color");
        expect(issues[0]!.path).toBe("word/document.xml");
        expect(issues[0]!.message).toContain("fill");
    });

    it("checks shading in numbering properties", async () => {
        // Issue #303: numbering's CT_Lvl admits pPr/rPr, which admit CT_Shd,
        // so an empty w:fill there is the same value-space defect wml.xsd
        // rejects — the part cannot be excluded from the scan.
        const numbering =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">` +
            `<w:lvl w:ilvl="0"><w:rPr><w:shd w:val="clear" w:fill=""/></w:rPr></w:lvl>` +
            `</w:abstractNum></w:numbering>`;
        const issues = await findings({ "word/numbering.xml": numbering });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ code: "word-shading-color", path: "word/numbering.xml", severity: "warning" });
    });

    it("skips a part that fails to parse without throwing, and still checks the others", async () => {
        const issues = await wordShadingColorCheck.run(
            fromParts({
                "word/document.xml": `<w:document xmlns:w="${W}"><w:body>`, // unterminated
                "word/footer1.xml": documentXml(`<w:p><w:pPr><w:shd w:val="clear" w:fill=""/></w:pPr></w:p>`),
            }),
        );
        expect(issues).toHaveLength(1);
        expect(issues[0]!.path).toBe("word/footer1.xml");
    });

    // #1122: since #1113 the story list holds the main document wherever it
    // is stored, so one stored under a styles or numbering name is in both
    // lists the check walks.
    describe("a main document stored under a styles or numbering name (#1122)", () => {
        const DEFECT = documentXml(`<w:p><w:pPr><w:shd w:val="clear" w:fill=""/></w:pPr></w:p>`);
        /** A package whose officeDocument relationship names `main`, holding one empty w:fill. */
        const mainAt = (main: string, extra: Readonly<Record<string, PartContent>> = {}): Record<string, PartContent> => ({
            "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', `Target="${main}"`),
            [main]: DEFECT,
            ...extra,
        });

        it.each(["word/styles.xml", "word/numbering.xml"])("reports the main document stored as %s once", async (main) => {
            expect((await findings(mainAt(main))).map((issue) => issue.path)).toEqual([main]);
        });

        // Two stored entries with OPC-equivalent names are two parts with
        // their own bytes, so each is walked: the walk keys on the stored
        // name, not the canonical one.
        it("walks each stored part once, a second spelling of the main document's name included", async () => {
            const styles =
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
                `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="P"><w:pPr><w:shd w:val="clear" w:fill=""/></w:pPr></w:style></w:styles>`;
            const issues = await findings(mainAt("word/Styles.xml", { "word/styles.xml": styles }));
            expect(issues.map((issue) => issue.path)).toEqual(["word/Styles.xml", "word/styles.xml"]);
        });
    });
});
