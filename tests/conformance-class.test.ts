/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { citesWordSession, STRICT_OOXML_NS_PREFIX } from "../src/lib/conformance-class";
import { ValidationContext } from "../src/lib/validation-context";
import { xmlPartNames } from "../src/scripts/office/checks/xml";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const P = MINIMAL_PARTS as Record<string, string>;

const cites = (parts: Record<string, string>, shapeParts?: readonly string[]): boolean => {
    const ctx = ValidationContext.fromEntries(parts);
    return citesWordSession(
        ctx,
        xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
        shapeParts,
    );
};

describe("citesWordSession (#1016)", () => {
    it("names the Strict namespace prefix", () => {
        expect(STRICT_OOXML_NS_PREFIX).toBe("http://purl.oclc.org/ooxml/");
    });

    it("cites a session for a Transitional package", () => {
        expect(cites(P)).toBe(true);
    });

    it("withholds the citation for a Strict root outside the shape", () => {
        const parts = { ...P, "word/document.xml": P["word/document.xml"]!.replaceAll(W, STRICT_W) };
        expect(cites(parts)).toBe(false);
        expect(
            cites({ ...P, "docProps/app.xml": `<Properties xmlns="${STRICT_OOXML_NS_PREFIX}officeDocument/extendedProperties"/>` }),
        ).toBe(false);
    });

    it("cites a session when the only Strict part is one the shape names, under any equivalent spelling", () => {
        const parts = { ...P, "word/comments.xml": `<w:comments xmlns:w="${STRICT_W}"/>` };
        expect(cites(parts, ["word/comments.xml"])).toBe(true);
        expect(cites(parts, ["/word/Comments.xml"])).toBe(true);
        expect(cites(parts, ["word/styles.xml"])).toBe(false);
    });

    it("reads a relationships part's class from its relationship types", () => {
        const parts = { ...P, "_rels/.rels": P["_rels/.rels"]!.replaceAll(R, STRICT_R) };
        expect(cites(parts)).toBe(false);
    });

    it("reads a root markup compatibility selects", () => {
        const parts = {
            ...P,
            "word/comments.xml":
                `<mc:AlternateContent xmlns:mc="${MC}" xmlns:w14="${W14}"><mc:Choice Requires="w14">` +
                `<w:comments xmlns:w="${STRICT_W}"/></mc:Choice><mc:Fallback><w:comments xmlns:w="${W}"/></mc:Fallback></mc:AlternateContent>`,
        };
        expect(cites(parts)).toBe(false);
    });

    it("reads no class from a part that does not parse", () => {
        expect(cites({ ...P, "word/comments.xml": `<w:comments xmlns:w="${STRICT_W}">` })).toBe(true);
    });

    it("parses each part once per context, however often it is asked", () => {
        const ctx = ValidationContext.fromEntries(P);
        const xmlParts = xmlPartNames(ctx.partNames, (name) => ctx.bytes(name));
        expect(citesWordSession(ctx, xmlParts)).toBe(true);
        expect(citesWordSession(ctx, xmlParts, ["word/document.xml"])).toBe(true);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });
});
