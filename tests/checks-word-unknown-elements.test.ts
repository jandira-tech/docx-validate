/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { WML_ELEMENT_NAMES } from "../src/scripts/office/checks/wml-vocabulary";
import {
    wordLinkWithoutNamespaceCheck as linkCheck,
    wordUnknownElementCheck as check,
} from "../src/scripts/office/checks/word-unknown-elements";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const document = (body: string, extra = "") =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:m="${M}" xmlns:r="${R}" ${extra}><w:body>${body}</w:body></w:document>`;

describe("word-unknown-element", () => {
    it("reports a w:p2 element at body level, as in UnknownElement.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:p2 w:rsidR="000402C2" w:rsidRDefault="000402C2"><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>Video provides…</w:t></w:r></w:p2>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:p2 sits directly inside w:body; wml\.xsd declares no w:p2/),
            }),
        ]);
    });

    it("tolerates unknown names inside property containers, as the Word-clean sdtPr and pPr specimens carry them", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:sdt><w:sdtPr><w:alias w:val="name"/><w:fieldType w:val="NAMETEXTINPUT"/><w:fieldTypeShort w:val="ti"/></w:sdtPr>` +
                    `<w:sdtContent><w:p><w:r><w:t>a</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
            ),
            "word/styles.xml": `<?xml version="1.0"?><w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:textFlow w:val="lrTb"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("reports unknown elements in the other block-level slots", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p2/></w:tc></w:tr></w:tbl>` +
                    `<w:sdt><w:sdtPr/><w:sdtContent><w:p3/></w:sdtContent></w:sdt>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ message: expect.stringMatching(/w:p2 sits directly inside w:tc;/) }),
            expect.objectContaining({ message: expect.stringMatching(/w:p3 sits directly inside w:sdtContent;/) }),
        ]);
    });

    it("accepts declared WordprocessingML elements wherever they occur", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:p><w:pPr><w:rPr/></w:pPr><w:r><w:rPr/><w:t>text</w:t></w:r></w:p><w:sectPr><w:pgSz/></w:sectPr>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
        for (const name of ["p", "pPr", "rPr", "tblBorders", "sectPr", "hyperlink", "sdtContent"]) {
            expect(WML_ELEMENT_NAMES.has(name)).toBe(true);
        }
        expect(WML_ELEMENT_NAMES.has("p2")).toBe(false);
    });

    it("accepts extension-namespace elements and skips non-WordprocessingML and unparseable parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><m:oMath><m:r><m:t>1</m:t></m:r></m:oMath></w:p>`, `xmlns:w14="${W14}"`),
            "customXml/item1.xml": `<b:Sources xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>`,
            "word/strict.xml": `<?xml version="1.0"?><w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body><w:p><w:p2/></w:p></w:body></w:document>`,
            "word/broken.xml": "<w:p2",
        });
        expect(ctx.xml("word/broken.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});

describe("word-link-without-namespace", () => {
    it("reports a namespace-less link inside run properties, as in hyperlink-exported.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:hyperlink w:history="0" r:id="rId9"><w:r><w:rPr><link /></w:rPr><w:t>Click</w:t></w:r></w:hyperlink>`,
            ),
        });
        expect(await linkCheck.run(ctx)).toEqual([
            expect.objectContaining({
                code: linkCheck.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/link .* without an XML namespace, directly inside w:rPr;/),
            }),
        ]);
    });

    it("accepts the namespaced w:link style child and foreign link elements, and skips broken parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/styles.xml": `<?xml version="1.0"?><w:styles xmlns:w="${W}"><w:style w:type="character"><w:link w:val="Strong"/></w:style></w:styles>`,
            "word/document.xml": document(`<w:p><x:link xmlns:x="urn:x"/></w:p>`),
            "word/broken.xml": "<link",
        });
        expect(await linkCheck.run(ctx)).toEqual([]);
    });
});
