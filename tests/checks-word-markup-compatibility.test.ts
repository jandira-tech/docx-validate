/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { wordMarkupCompatibilityCheck as check } from "../src/scripts/office/checks/word-markup-compatibility";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const ROOT = `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" xmlns:e="urn:future"`;
const document = (body: string, attributes = "") => `<w:document ${ROOT}${attributes}><w:body>${body}</w:body></w:document>`;
const FALLBACK_ONLY = "<mc:AlternateContent><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>";

describe("markup compatibility as Word processes it", () => {
    it("rejects an AlternateContent without a Choice and a namespace Word must understand but does not", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(FALLBACK_ONLY),
            "word/footnotes.xml": `<w:footnotes ${ROOT} mc:MustUnderstand="e"/>`,
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
            expect.objectContaining({
                code: "mce-must-understand",
                severity: "error",
                path: "word/footnotes.xml",
                message: expect.stringContaining("e"),
            }),
        ]);
    });

    it("accepts what Word understands: a w14 requirement and a Choice with a Fallback", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<mc:AlternateContent><mc:Choice Requires="e"><w:tbl/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>`,
                ` mc:MustUnderstand="w14"`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("ignores content in a branch Word does not select", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback>${FALLBACK_ONLY}</mc:Fallback></mc:AlternateContent>`,
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("reports errors only, and judges only the XML parts under word/", async () => {
        const ctx = ValidationContext.fromEntries({
            // An Office extension attribute outside mc:Ignorable is a warning
            // of the schema view, not a Word rejection.
            "word/settings.xml": `<w:settings ${ROOT} w14:foo="1"/>`,
            "docProps/custom.xml": document(FALLBACK_ONLY),
            "word/_rels/document.xml.rels": document(FALLBACK_ONLY),
            "word/styles.xml": document(FALLBACK_ONLY).replace("</w:document>", ""),
        });
        expect(await check.run(ctx)).toEqual([]);
    });
});
