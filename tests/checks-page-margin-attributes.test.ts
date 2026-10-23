/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for `pageMarginAttributesCheck` (issue #214, jubarte lane
 * family k10). CT_PageMar (wml.xsd) requires all seven of top, right,
 * bottom, left, header, footer and gutter; Word silently defaults the
 * missing ones, so every finding is an advisory, one per missing
 * attribute.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { pageMarginAttributesCheck } from "../src/scripts/office/checks/page-margin-attributes";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const COMPLETE_MARGINS = 'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"';

const documentWith = (body: string, ns = W): string =>
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`;

const run = (entries: Readonly<Record<string, string>>) => pageMarginAttributesCheck.run(ValidationContext.fromEntries(entries));

describe("page-margin-attributes check", () => {
    it("names each missing required attribute of a body-level w:pgMar", async () => {
        const xml = documentWith(
            `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>`,
        );
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(3);
        for (const issue of issues) {
            expect(issue).toMatchObject({ severity: "warning", code: "page-margin-attributes", path: "word/document.xml" });
        }
        expect(issues.map((issue) => issue.message)).toEqual([
            expect.stringContaining("w:header"),
            expect.stringContaining("w:footer"),
            expect.stringContaining("w:gutter"),
        ]);
        expect(issues[0]!.message).toContain("CT_PageMar");
    });

    it("leaves a complete w:pgMar unreported", async () => {
        const xml = documentWith(`<w:sectPr><w:pgMar ${COMPLETE_MARGINS}/></w:sectPr>`);
        expect(await run({ "word/document.xml": xml })).toEqual([]);
    });

    it("leaves a w:sectPr without w:pgMar unreported", async () => {
        const xml = documentWith(`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`);
        expect(await run({ "word/document.xml": xml })).toEqual([]);
    });

    it("reports a paragraph-level w:pPr/w:sectPr and the w:sectPr inside w:sectPrChange", async () => {
        const xml = documentWith(
            `<w:p><w:pPr><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:pPr></w:p>` +
                `<w:sectPr><w:pgMar ${COMPLETE_MARGINS}/>` +
                `<w:sectPrChange w:author="a" w:id="1" w:date="1970-01-01T00:00:00Z">` +
                `<w:sectPr><w:pgMar w:top="1060" w:right="1000" w:bottom="1060" w:left="1000" w:footer="300"/><w:titlePg/></w:sectPr>` +
                `</w:sectPrChange></w:sectPr>`,
        );
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(3);
        expect(issues.map((issue) => issue.message)).toEqual([
            expect.stringContaining("w:gutter"),
            expect.stringContaining("w:header"),
            expect.stringContaining("w:gutter"),
        ]);
        expect(issues[0]!.message).toContain("w:pPr");
        expect(issues[1]!.message).toContain("w:sectPrChange");
    });

    it("scans headers, footers, footnotes, endnotes and comments, not only document.xml", async () => {
        const defective = `<w:sectPr><w:pgMar w:top="1440"/></w:sectPr>`;
        const parts = {
            "word/document.xml": documentWith(`<w:sectPr><w:pgMar ${COMPLETE_MARGINS}/></w:sectPr>`),
            "word/header1.xml": `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:pPr>${defective}</w:pPr></w:p></w:hdr>`,
            "word/footer2.xml": `<?xml version="1.0"?><w:ftr xmlns:w="${W}"><w:p><w:pPr>${defective}</w:pPr></w:p></w:ftr>`,
            "word/footnotes.xml": `<?xml version="1.0"?><w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:pPr>${defective}</w:pPr></w:p></w:footnote></w:footnotes>`,
            "word/endnotes.xml": `<?xml version="1.0"?><w:endnotes xmlns:w="${W}"><w:endnote w:id="1"><w:p><w:pPr>${defective}</w:pPr></w:p></w:endnote></w:endnotes>`,
            "word/comments.xml": `<?xml version="1.0"?><w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="a" w:date="1970-01-01T00:00:00Z"><w:p><w:pPr>${defective}</w:pPr></w:p></w:comment></w:comments>`,
            "word/numbering.xml": `<?xml version="1.0"?><w:numbering xmlns:w="${W}">${defective}</w:numbering>`,
        };
        const issues = await run(parts);
        expect([...new Set(issues.map((issue) => issue.path))].sort()).toEqual([
            "word/comments.xml",
            "word/endnotes.xml",
            "word/footer2.xml",
            "word/footnotes.xml",
            "word/header1.xml",
        ]);
        // Six missing attributes per part; numbering.xml is not a story part.
        expect(issues).toHaveLength(30);
    });

    it("checks ISO Strict namespace page margins the same way", async () => {
        const xml = documentWith(`<w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`, W_STRICT);
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(3);
        expect(issues.map((issue) => issue.code)).toEqual(["page-margin-attributes", "page-margin-attributes", "page-margin-attributes"]);
    });

    it("skips a part that does not parse instead of throwing", async () => {
        const issues = await run({
            "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:sectPr>`,
        });
        expect(issues).toEqual([]);
    });
});
