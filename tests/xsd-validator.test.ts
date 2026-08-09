/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import { describe, it, expect, beforeAll } from "vitest";
import path from "node:path";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { createXsdValidator, type XsdValidator } from "../src/lib/xsd-validator";

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

const OOXML_WML_SCHEMA = path.resolve(
  __dirname,
  "..",
  "src",
  "scripts",
  "office",
  "schemas",
  "ISO-IEC29500-4_2016",
  "wml.xsd",
);

describe("xsd-validator", () => {
  let validator: XsdValidator;
  let tempDir: string;
  let inlineSchemaPath: string;

  beforeAll(async () => {
    validator = await createXsdValidator();
    tempDir = mkdtempSync(path.join(tmpdir(), "xsd-validator-test-"));
    inlineSchemaPath = path.join(tempDir, "self-contained.xsd");
    writeFileSync(inlineSchemaPath, SELF_CONTAINED_XSD, "utf-8");
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
    const issues = await validator.validate(VALID_XML, inlineSchemaPath);
    expect(issues).toEqual([]);
  });

  it("reports a structured ValidationIssue on schema-invalid input", async () => {
    const issues = await validator.validate(SCHEMA_INVALID_XML, inlineSchemaPath);
    expect(issues.length).toBeGreaterThan(0);
    const first = issues[0]!;
    expect(first.severity).toBe("error");
    expect(first.code).toBe("xsd-validation-failed");
    expect(typeof first.message).toBe("string");
    expect(first.message.length).toBeGreaterThan(0);
  });

  it("reports xml-parse-error on malformed XML (not XSD failure)", async () => {
    const issues = await validator.validate(NOT_XML, inlineSchemaPath);
    expect(issues.length).toBe(1);
    expect(issues[0]!.code).toBe("xml-parse-error");
    expect(issues[0]!.severity).toBe("error");
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

  it("gracefully degrades on a non-existent schema path", async () => {
    const issues = await validator.validate(VALID_XML, "/tmp/this-schema-does-not-exist.xsd");
    expect(issues.length).toBe(1);
    expect(issues[0]!.code).toBe("xsd-schema-load-skipped");
    expect(issues[0]!.severity).toBe("info");
  });

  it("cleanup", () => {
    rmSync(tempDir, { recursive: true, force: true });
  });
});
