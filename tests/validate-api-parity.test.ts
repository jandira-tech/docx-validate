/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * API parity (docs/plans/2026-10-23-fixing-validator.md, task 11).
 *
 * Path-based `validate()`, the `Validate` class and the `jubarte` helper must
 * forward the same applicable options and return the same execution metadata,
 * so a caller gets the same answer whichever entrypoint it uses.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { jubarte, Validate } from "../src/index";
import { withTempDir } from "../src/lib/run-cli";
import type { XsdValidator } from "../src/lib/xsd-validator";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

/** Reports every part as failing to load its schema, so a run is unmistakably invalid. */
const schemaLoadFailure: XsdValidator = {
    validate: async (_xml, schemaPath) => [
        { severity: "error", code: "xsd-schema-load-failed", message: `Schema load failed (${schemaPath}): injected` },
    ],
};

const controlBytes = async (dir: string): Promise<{ file: string; bytes: Buffer }> => {
    const bytes = await buildDocx();
    const file = path.join(dir, "control.docx");
    writeFileSync(file, bytes);
    return { file, bytes };
};

describe("an injected XsdValidator reaches every entrypoint", { timeout: 120_000 }, () => {
    it("reaches the path-taking validate()", async () => {
        const result = await withTempDir(async (dir) =>
            validate((await controlBytes(dir)).file, { profile: "strict", xsdValidator: schemaLoadFailure }),
        );

        expect(result.valid).toBe(false);
        expect(result.metadata?.schemaCompleted).toBe(false);
        expect(result.issues.some((issue) => issue.code === "xsd-summary")).toBe(true);
    });

    it("reaches the Validate class", async () => {
        const result = await withTempDir(async (dir) => {
            const { bytes } = await controlBytes(dir);
            return new Validate({ profile: "strict", xsdValidator: schemaLoadFailure }).run(bytes);
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.schemaCompleted).toBe(false);
        expect(result.issues.some((issue) => issue.code === "xsd-summary")).toBe(true);
    });

    it("reaches the jubarte helper", async () => {
        const result = await withTempDir(async (dir) => {
            const { bytes } = await controlBytes(dir);
            return jubarte.validate(bytes, { profile: "strict", xsdValidator: schemaLoadFailure });
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.schemaCompleted).toBe(false);
        expect(result.issues.some((issue) => issue.code === "xsd-summary")).toBe(true);
    });
});

describe("equivalent bytes and path inputs agree", { timeout: 120_000 }, () => {
    it("Validate.run(bytes) and validate(path) return the same issues and metadata", async () => {
        await withTempDir(async (dir) => {
            const { file, bytes } = await controlBytes(dir);
            const fromPath = await validate(file);
            const fromBytes = await new Validate().run(bytes);

            expect(fromBytes.valid).toBe(fromPath.valid);
            expect(fromBytes.issues).toEqual(fromPath.issues);
            expect(fromBytes.metadata).toEqual(fromPath.metadata);
        });
    });

    it("jubarte.validate(bytes) and validate(path) return the same issues and metadata", async () => {
        await withTempDir(async (dir) => {
            const { file, bytes } = await controlBytes(dir);
            const fromPath = await validate(file);
            const fromBytes = await jubarte.validate(bytes);

            expect(fromBytes.valid).toBe(fromPath.valid);
            expect(fromBytes.issues).toEqual(fromPath.issues);
            expect(fromBytes.metadata).toEqual(fromPath.metadata);
        });
    });
});
