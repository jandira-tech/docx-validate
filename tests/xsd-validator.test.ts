/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import { describe, it, expect, afterEach, beforeAll, vi } from "vitest";
import path from "node:path";
import { writeFileSync } from "node:fs";
import {
    _resetSchemaResourceRoots,
    _resetXsdValidatorMemo,
    allowSchemaResourceRoot,
    createXsdValidator,
    schemaResourceRoots,
    xsdValidateStats,
    type XsdValidator,
} from "../src/lib/xsd-validator";
import { withTempDir } from "../src/lib/run-cli";

// Self-contained XSD with no imports — proves the validator core works.
// The bundled OOXML schemas have unresolved imports (CLAUDE.md note 4); they
// exercise the graceful-degradation path below.
const SELF_CONTAINED_XSD = `<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="root">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="child" type="xs:string"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>`;

const VALID_XML = `<?xml version="1.0"?><root><child>hello</child></root>`;
const SCHEMA_INVALID_XML = `<?xml version="1.0"?><root><wrong/></root>`;
const NOT_XML = "this is not xml at all";

/**
 * A schema in its own namespace that an importing schema pulls in. The
 * importing schema's root element is typed with a type that only exists in the
 * import, so the import is load-bearing: if it does not resolve, compilation
 * fails instead of silently degrading.
 */
const IMPORTED_XSD = `<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="urn:imported"
  xmlns:i="urn:imported" elementFormDefault="qualified">
  <xs:complexType name="shared">
    <xs:simpleContent><xs:extension base="xs:string"/></xs:simpleContent>
  </xs:complexType>
</xs:schema>`;

const importingXsd = (schemaLocation: string): string =>
    `<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:i="urn:imported">` +
    `<xs:import namespace="urn:imported" schemaLocation="${schemaLocation}"/>` +
    `<xs:element name="root" type="i:shared"/></xs:schema>`;

const IMPORTED_DOC = `<?xml version="1.0"?><root>text</root>`;

const OOXML_WML_SCHEMA = path.resolve(__dirname, "..", "src", "scripts", "office", "schemas", "ISO-IEC29500-4_2016", "wml.xsd");

describe("xsd-validator", () => {
    let validator: XsdValidator;

    beforeAll(async () => {
        validator = await createXsdValidator();
    });

    it("createXsdValidator returns an object with an async validate()", () => {
        expect(typeof validator.validate).toBe("function");
    });

    it("memoises the validator (same instance across calls)", async () => {
        const a = await createXsdValidator();
        const b = await createXsdValidator();
        expect(a).toBe(b);
    });

    it("validates a known-good document with zero issues (self-contained XSD)", async () => {
        await withTempDir(async (tempDir) => {
            const inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
            writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
            const issues = await validator.validate(VALID_XML, inlineSchemaPath);
            expect(issues).toEqual([]);
        });
    });

    it("reports a structured ValidationIssue on schema-invalid input", async () => {
        await withTempDir(async (tempDir) => {
            const inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
            writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
            const issues = await validator.validate(SCHEMA_INVALID_XML, inlineSchemaPath);
            expect(issues.length).toBeGreaterThan(0);
            const first = issues[0]!;
            expect(first.severity).toBe("error");
            expect(first.code).toBe("xsd-validation-failed");
            expect(typeof first.message).toBe("string");
            expect(first.message.length).toBeGreaterThan(0);
        });
    });

    it("falls back to public validation when private WASM pointers are unavailable", async () => {
        const { XsdValidator: WasmValidator } = await import("libxml2-wasm");
        const fromDoc = WasmValidator.fromDoc.bind(WasmValidator);
        const factory = vi.spyOn(WasmValidator, "fromDoc").mockImplementation((...args) => {
            const instance = fromDoc(...args);
            return new Proxy(instance, {
                get(target, property, receiver) {
                    if (property === "_ptr") return undefined;
                    const value: unknown = Reflect.get(target, property, receiver);
                    return typeof value === "function" ? value.bind(target) : value;
                },
            });
        });
        _resetXsdValidatorMemo();
        try {
            const fallback = await createXsdValidator();
            await withTempDir(async (dir) => {
                const schema = path.join(dir, "fallback.xsd");
                writeFileSync(schema, SELF_CONTAINED_XSD, "utf-8");
                expect(await fallback.validate(VALID_XML, schema)).toEqual([]);
                const issues = await fallback.validate(SCHEMA_INVALID_XML, schema);
                expect(issues[0]?.code).toBe("xsd-validation-failed");
                expect(issues[0]?.message).toContain("wrong");
            });
        } finally {
            factory.mockRestore();
            _resetXsdValidatorMemo();
        }
    });

    it("reports xml-parse-error on malformed XML (not XSD failure)", async () => {
        await withTempDir(async (tempDir) => {
            const inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
            writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
            const issues = await validator.validate(NOT_XML, inlineSchemaPath);
            expect(issues.length).toBe(1);
            expect(issues[0]!.code).toBe("xml-parse-error");
            expect(issues[0]!.severity).toBe("error");
        });
    });

    it("loads the bundled OOXML wml.xsd cleanly (fs input providers resolve imports)", async () => {
        // The bundled OOXML schemas have relative <xs:import schemaLocation=
        // ".../sharedTypes.xsd"/> references. The wasm fs input providers +
        // the document base URL resolve them — proving the new engine is
        // strictly stronger than libxmljs2's silent-degradation behaviour
        // documented in CLAUDE.md note 4.
        //
        // The XML below isn't valid against wml.xsd (wrong root element), so
        // expect xsd-validation-failed — that proves schema LOAD succeeded.
        const issues = await validator.validate(VALID_XML, OOXML_WML_SCHEMA);
        expect(issues.length).toBeGreaterThan(0);
        expect(issues[0]!.code).toBe("xsd-validation-failed");
    });

    it("surfaces a non-existent schema path as an error (never silently valid)", async () => {
        const issues = await validator.validate(VALID_XML, "/tmp/this-schema-does-not-exist.xsd");
        expect(issues.length).toBe(1);
        expect(issues[0]!.code).toBe("xsd-schema-load-failed");
        expect(issues[0]!.severity).toBe("error");
    });

    describe("external resources", () => {
        afterEach(() => {
            _resetSchemaResourceRoots();
        });

        it("never reads a file named by an external entity in the document's DTD", async () => {
            await withTempDir(async (dir) => {
                const secret = path.join(dir, "secret.txt");
                writeFileSync(secret, "TOP-SECRET-HOSTS-MARKER", "utf-8");
                const inlineSchemaPath = path.join(dir, "self-contained.xsd");
                writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
                const document = `<?xml version="1.0"?><!DOCTYPE root [<!ENTITY e SYSTEM "file://${secret}">]><root><child>&e;</child></root>`;

                const issues = await validator.validate(document, inlineSchemaPath);

                expect(JSON.stringify(issues)).not.toContain("TOP-SECRET-HOSTS-MARKER");
                expect(issues.map((i) => i.code)).not.toContain("xml-parse-error");
            });
        });

        it("refuses to resolve a schema import outside the allowed schema roots", async () => {
            await withTempDir(async (dir) => {
                const imported = path.join(dir, "imported.xsd");
                writeFileSync(imported, IMPORTED_XSD, "utf-8");
                const importingSchema = path.join(dir, "importing.xsd");
                writeFileSync(importingSchema, importingXsd(`file://${imported}`), "utf-8");

                const issues = await validator.validate(IMPORTED_DOC, importingSchema);

                expect(issues.map((i) => i.code)).toEqual(["xsd-schema-load-failed"]);
                expect(issues[0]!.severity).toBe("error");
                expect(issues[0]!.message).toContain("imported.xsd");
            });
        });

        it("resolves a schema import once its directory is an allowed root", async () => {
            await withTempDir(async (dir) => {
                const imported = path.join(dir, "imported.xsd");
                writeFileSync(imported, IMPORTED_XSD, "utf-8");
                const importingSchema = path.join(dir, "importing.xsd");
                writeFileSync(importingSchema, importingXsd(`file://${imported}`), "utf-8");

                allowSchemaResourceRoot(dir);

                expect(schemaResourceRoots()).toContain(dir);
                const issues = await validator.validate(IMPORTED_DOC, importingSchema);
                expect(issues).toEqual([]);
            });
        });

        it("resolves the bundled OOXML schema import closure from the default root", async () => {
            expect(schemaResourceRoots()).toContain(path.resolve(__dirname, "..", "src", "scripts", "office", "schemas"));
            const issues = await validator.validate(VALID_XML, OOXML_WML_SCHEMA);
            expect(issues.map((i) => i.code)).toEqual(["xsd-validation-failed"]);
        });

        it("_resetSchemaResourceRoots drops explicitly allowed roots but keeps the bundled tree", async () => {
            await withTempDir(async (dir) => {
                allowSchemaResourceRoot(dir);
                expect(schemaResourceRoots()).toContain(dir);
                _resetSchemaResourceRoots();
                expect(schemaResourceRoots()).not.toContain(dir);
                expect(schemaResourceRoots()).toContain(path.resolve(__dirname, "..", "src", "scripts", "office", "schemas"));
            });
        });
    });

    it("xsdValidateStats counts status-only passes and message collections on the real engine path", async () => {
        await withTempDir(async (tempDir) => {
            const inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
            writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
            const passesBefore = xsdValidateStats.statusOnlyPasses;
            const messagesBefore = xsdValidateStats.messageCollections;
            expect(await validator.validate(VALID_XML, inlineSchemaPath)).toEqual([]);
            expect(xsdValidateStats.statusOnlyPasses).toBeGreaterThan(passesBefore);
            expect(await validator.validate(SCHEMA_INVALID_XML, inlineSchemaPath)).not.toEqual([]);
            expect(xsdValidateStats.messageCollections).toBeGreaterThan(messagesBefore);
        });
    });

    it("maps XmlValidateError details onto ValidationIssue line numbers", async () => {
        await withTempDir(async (tempDir) => {
            const inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
            writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
            const issues = await validator.validate(SCHEMA_INVALID_XML, inlineSchemaPath);
            expect(issues.length).toBeGreaterThan(0);
            expect(issues.some((issue) => typeof issue.line === "number")).toBe(true);
        });
    });
});
