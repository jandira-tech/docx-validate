/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { propertiesParts } from "../src/lib/properties-parts";
import { NS } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";

const TYPES = [
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties",
] as const;
const relationship = (target: string, type: string = TYPES[0], attributes = ""): string =>
    `<Relationship Id="properties" Type="${type}" Target="${target}" ${attributes}/>`;
const relationships = (xml: string): string => `<Relationships xmlns="${NS.PR}">${xml}</Relationships>`;
const context = (xml: string, name = "_rels/.rels"): ValidationContext => ValidationContext.fromEntries({ [name]: xml });

describe("propertiesParts", () => {
    it("does not read properties through nested package relationships (#434)", async () => {
        expect(
            await propertiesParts(context(relationships(`<x:Wrap xmlns:x="urn:x">${relationship("docProps/core.xml")}</x:Wrap>`))),
        ).toEqual(new Set());
    });

    it.each(TYPES)("recognizes %s", async (type) => {
        expect(await propertiesParts(context(relationships(relationship("docProps/core.xml", type))))).toEqual(
            new Set(["docprops/core.xml"]),
        );
    });

    it.each([
        ["/docProps/core.xml", "docprops/core.xml"],
        ["docProps/../docProps/core.xml", "docprops/core.xml"],
        ["props/meta.xml", "props/meta.xml"],
        ["DOCPROPS/%63ORE.XML", "docprops/core.xml"],
    ])("resolves root-relative Target %s", async (target, expected) => {
        expect(await propertiesParts(context(relationships(relationship(target))))).toEqual(new Set([expected]));
    });

    it.each(["_RELS/.RELS", "_rels/.%72els"])("reads OPC-equivalent package relationships at %s", async (name) => {
        expect(await propertiesParts(context(relationships(relationship("docProps/core.xml")), name))).toEqual(
            new Set(["docprops/core.xml"]),
        );
    });

    it("collects all properties kinds once, excluding external and unrelated relationships", async () => {
        const xml = relationships(
            relationship("docProps/core.xml") +
                relationship("docProps/app.xml", TYPES[2]) +
                relationship("docProps/custom.xml", TYPES[4]) +
                relationship("docProps/core.xml", TYPES[1]) +
                relationship("external.xml", TYPES[0], 'TargetMode="External"') +
                relationship("other.xml", "urn:other") +
                `<Relationship Id="no-type" Target="no-type.xml"/>` +
                `<Relationship Id="no-target" Type="${TYPES[0]}"/>` +
                relationship("../outside.xml") +
                relationship("/"),
        );
        expect(await propertiesParts(context(xml))).toEqual(new Set(["docprops/core.xml", "docprops/app.xml", "docprops/custom.xml"]));
    });

    it("ignores an external properties target", async () => {
        expect(await propertiesParts(context(relationships(relationship("docProps/core.xml", TYPES[0], 'TargetMode="External"'))))).toEqual(
            new Set(),
        );
    });

    it("returns no properties when package relationships are missing", async () => {
        const ctx = ValidationContext.fromEntries({
            "docProps/core.xml": "<core/>",
            "word/_rels/document.xml.rels": relationships(relationship("docProps/core.xml")),
        });
        expect(await propertiesParts(ctx)).toEqual(new Set());
    });

    it.each([
        "",
        `<Relationships xmlns="${NS.PR}">${relationship("docProps/core.xml")}`,
        `<Relationships xmlns="${NS.PR}">${relationship("docProps/core.xml")}<bad:child/></Relationships>`,
        `<?xml version="1.0" encoding="X-UNKNOWN-ENCODING"?>${relationships(relationship("docProps/core.xml"))}`,
    ])("returns no properties for unreadable package relationships: %j", async (xml) => {
        expect(await propertiesParts(context(xml))).toEqual(new Set());
    });
});
