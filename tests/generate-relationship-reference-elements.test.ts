/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
    collectRelationshipElements,
    readBundledSchemas,
    renderRelationshipElements,
} from "../scripts/generate-relationship-reference-elements";

const XSD = "http://www.w3.org/2001/XMLSchema";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const schema = (body: string, ns = "urn:test") =>
    `<xs:schema xmlns:xs="${XSD}" xmlns="${ns}" xmlns:r="${R}" targetNamespace="${ns}" elementFormDefault="qualified">${body}</xs:schema>`;

describe("relationship schema declaration resolver", () => {
    it("follows element refs, anonymous types, attribute refs/groups and extension/restriction bases", () => {
        const result = collectRelationshipElements([
            { name: "relationships.xsd", xml: schema('<xs:attribute name="id" type="xs:string"/>', R) },
            {
                name: "test.xsd",
                xml: schema(`
                <xs:attributeGroup name="inner"><xs:attribute ref="r:id"/></xs:attributeGroup>
                <xs:attributeGroup name="outer"><xs:attributeGroup ref="inner"/></xs:attributeGroup>
                <xs:complexType name="base"><xs:attributeGroup ref="outer"/></xs:complexType>
                <xs:complexType name="extension"><xs:complexContent><xs:extension base="base"/></xs:complexContent></xs:complexType>
                <xs:complexType name="restriction"><xs:complexContent><xs:restriction base="extension"/></xs:complexContent></xs:complexType>
                <xs:element name="global" type="restriction"/>
                <xs:element name="inherited" substitutionGroup="global"/>
                <xs:element name="root"><xs:complexType><xs:sequence>
                    <xs:element ref="global"/>
                    <xs:element name="local" type="extension"/>
                    <xs:element name="inline"><xs:complexType><xs:attribute ref="r:id"/></xs:complexType></xs:element>
                    <xs:element name="plain" type="xs:string"/>
                    <xs:element name="global" type="xs:string"/>
                </xs:sequence></xs:complexType></xs:element>
            `),
            },
        ]);
        for (const name of ["global", "local", "inline", "inherited"]) {
            expect(result.references.get(`{urn:test}${name}`)).toEqual(new Set(["id"]));
        }
        expect(result.declared.has("{urn:test}plain")).toBe(true);
        expect(result.references.has("{urn:test}plain")).toBe(false);
        expect(result.references.has("{urn:test}root")).toBe(false);
        expect(result.uncertain.size).toBe(0);
    });

    it("honors prohibited attributes, local form, and keeps unresolved types and anyType conservative", () => {
        const result = collectRelationshipElements([
            {
                name: "test.xsd",
                xml: schema(`
            <xs:complexType name="base"><xs:attribute ref="r:id"/></xs:complexType>
            <xs:complexType name="restriction"><xs:complexContent><xs:restriction base="base"><xs:attribute ref="r:id" use="prohibited"/></xs:restriction></xs:complexContent></xs:complexType>
            <xs:element name="removed" type="restriction"/>
            <xs:element name="closed"><xs:complexType><xs:complexContent><xs:restriction base="xs:anyType"/></xs:complexContent></xs:complexType></xs:element>
            <xs:element name="unknown" type="missing"/>
            <xs:element name="any"/>
            <xs:element name="root"><xs:complexType><xs:sequence><xs:element name="unqualified" form="unqualified" type="xs:string"/></xs:sequence></xs:complexType></xs:element>
        `),
            },
        ]);
        expect(result.references.has("{urn:test}removed")).toBe(false);
        expect(result.uncertain.has("{urn:test}closed")).toBe(false);
        expect(result.uncertain).toEqual(new Set(["{urn:test}unknown", "{urn:test}any"]));
        expect(result.declared.has("{}unqualified")).toBe(true);
    });

    it("keeps unresolved element refs, cycles and attribute wildcards uncertain", () => {
        const data = collectRelationshipElements([
            {
                name: "test.xsd",
                xml: schema(`
            <xs:complexType name="cycle"><xs:complexContent><xs:extension base="cycle"/></xs:complexContent></xs:complexType>
            <xs:element name="cyclic" type="cycle"/>
            <xs:element name="wild"><xs:complexType><xs:anyAttribute/></xs:complexType></xs:element>
            <xs:element name="root"><xs:complexType><xs:sequence><xs:element ref="missing"/></xs:sequence></xs:complexType></xs:element>
        `),
            },
        ]);
        expect(data.uncertain).toEqual(new Set(["{urn:test}cyclic", "{urn:test}wild", "{urn:test}missing"]));
        expect(data.uncertain.has("{urn:test}root")).toBe(false);
    });

    it("unions ambiguous declarations and Transitional/Strict counterparts by schema filename", () => {
        const result = collectRelationshipElements([
            {
                name: "ISO-IEC29500-4_2016/test.xsd",
                xml: schema(
                    '<xs:element name="same" type="xs:string"/><xs:element name="root"><xs:complexType><xs:sequence><xs:element name="same"><xs:complexType><xs:attribute ref="r:id"/></xs:complexType></xs:element></xs:sequence></xs:complexType></xs:element>',
                ),
            },
            { name: "iso-strict/test.xsd", xml: schema('<xs:element name="same" type="xs:string"/>', "urn:strict") },
        ]);
        expect(result.references.get("{urn:test}same")).toEqual(new Set(["id"]));
        expect(result.references.get("{urn:strict}same")).toEqual(new Set(["id"]));
        expect(result.declared.has("{urn:strict}root")).toBe(true);
    });
});

describe("bundled relationship declaration generation (integration)", () => {
    it("regenerates the checked-in module in memory without drift", async () => {
        const data = collectRelationshipElements(readBundledSchemas());
        expect(await renderRelationshipElements(data)).toBe(
            readFileSync(new URL("../src/lib/relationship-reference-elements.ts", import.meta.url), "utf8"),
        );
        expect(data.declared.has("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")).toBe(true);
        expect(data.references.has("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")).toBe(false);
        expect(data.references.get("{http://schemas.openxmlformats.org/drawingml/2006/chart}chart")).toEqual(new Set(["id"]));
        expect(data.references.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top")).toContain("id");
        expect(data.declared.has("{http://schemas.microsoft.com/office/word/2010/wordml}checkbox")).toBe(true);
    });
});
