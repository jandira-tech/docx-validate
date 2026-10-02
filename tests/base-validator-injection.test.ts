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
import { writeFileSync, mkdirSync } from "node:fs";
import path from "node:path";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import type { XsdValidator } from "../src/lib/xsd-validator";
import type { ValidationIssue } from "../src/lib/types";
import { withTempDir } from "../src/lib/run-cli";

const runWithTinyUnpackedDir = async <T>(fn: (dir: string) => Promise<T> | T): Promise<T> => {
    return withTempDir(async (dir) => {
        mkdirSync(path.join(dir, "word"), { recursive: true });
        writeFileSync(
            path.join(dir, "word", "document.xml"),
            `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`,
            "utf-8",
        );
        return await fn(dir);
    });
};

describe("BaseSchemaValidator XsdValidator injection", () => {
    it("accepts xsdValidator in constructor opts", async () => {
        await runWithTinyUnpackedDir((dir) => {
            const v = new BaseSchemaValidator({
                unpackedDir: dir,
                xsdValidator: { validate: async () => [] },
            });
            expect(v).toBeInstanceOf(BaseSchemaValidator);
        });
    });

    it("delegates _validateSingleFileXsd to the injected validator", async () => {
        await runWithTinyUnpackedDir(async (dir) => {
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
        });
    });

    it("translates a returned error-severity issue into XsdValidationOutcome.errors", async () => {
        await runWithTinyUnpackedDir(async (dir) => {
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
        });
    });

    it("treats info-severity issues as non-fatal (CLAUDE.md note 4 spirit)", async () => {
        await runWithTinyUnpackedDir(async (dir) => {
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
        });
    });
});
