/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * PR B Task B.1: assert BaseSchemaValidator accepts an injected XsdValidator
 * and delegates XSD checks to it. The default code path (no injection) goes
 * through the wasm validator from PR A.
 */

import { describe, it, expect } from "vitest";
import { mkdtempSync, rmSync, writeFileSync, mkdirSync } from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import type { XsdValidator } from "../src/lib/xsd-validator";
import type { ValidationIssue } from "../src/lib/types";

const tinyUnpackedDir = (): string => {
  const dir = mkdtempSync(path.join(tmpdir(), "base-injection-"));
  mkdirSync(path.join(dir, "word"), { recursive: true });
  writeFileSync(
    path.join(dir, "word", "document.xml"),
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`,
    "utf-8",
  );
  return dir;
};

describe("BaseSchemaValidator XsdValidator injection", () => {
  it("accepts xsdValidator in constructor opts", () => {
    const dir = tinyUnpackedDir();
    try {
      const v = new BaseSchemaValidator({
        unpackedDir: dir,
        xsdValidator: { validate: async () => [] },
      });
      expect(v).toBeInstanceOf(BaseSchemaValidator);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("delegates _validateSingleFileXsd to the injected validator", async () => {
    const dir = tinyUnpackedDir();
    try {
      const calls: { xml: string; schemaPath: string }[] = [];
      const fakeValidator: XsdValidator = {
        async validate(xml: string, schemaPath: string): Promise<ValidationIssue[]> {
          calls.push({ xml, schemaPath });
          return [];
        },
      };

      const v = new BaseSchemaValidator({
        unpackedDir: dir,
        xsdValidator: fakeValidator,
      });

      const outcome = await v.validateFileAgainstXsd(path.join(dir, "word", "document.xml"));

      expect(calls.length).toBe(1);
      expect(calls[0]!.schemaPath).toContain("wml.xsd");
      expect(outcome.valid).toBe(true);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("translates a returned error-severity issue into XsdValidationOutcome.errors", async () => {
    const dir = tinyUnpackedDir();
    try {
      const fakeValidator: XsdValidator = {
        async validate(): Promise<ValidationIssue[]> {
          return [{ severity: "error", code: "xsd-validation-failed", message: "fake error" }];
        },
      };
      const v = new BaseSchemaValidator({
        unpackedDir: dir,
        xsdValidator: fakeValidator,
      });

      const outcome = await v.validateFileAgainstXsd(path.join(dir, "word", "document.xml"));
      expect(outcome.valid).toBe(false);
      expect([...outcome.errors]).toContain("fake error");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("treats info-severity issues as non-fatal (CLAUDE.md note 4 spirit)", async () => {
    const dir = tinyUnpackedDir();
    try {
      const fakeValidator: XsdValidator = {
        async validate(): Promise<ValidationIssue[]> {
          return [
            { severity: "info", code: "xsd-schema-load-skipped", message: "schema not found" },
          ];
        },
      };
      const v = new BaseSchemaValidator({
        unpackedDir: dir,
        xsdValidator: fakeValidator,
      });

      const outcome = await v.validateFileAgainstXsd(path.join(dir, "word", "document.xml"));
      expect(outcome.valid).toBe(true);
      expect(outcome.errors.size).toBe(0);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
