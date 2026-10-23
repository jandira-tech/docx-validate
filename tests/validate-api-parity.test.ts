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

import { mkdirSync, symlinkSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { jubarte, ModeConflictError, resolveMode, Validate } from "../src/index";
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

describe("the public barrel exposes the mode policy", () => {
    it("re-exports the resolver and the typed conflict error", async () => {
        expect(resolveMode("docx", { strict: true })).toEqual({ profile: "strict", select: undefined });
        expect(new ModeConflictError("boom")).toBeInstanceOf(Error);
        // A consumer can catch the library's typed error by class.
        await withTempDir(async (dir) => {
            const { file } = await controlBytes(dir);
            await expect(validate(file, { strict: true, quick: true })).rejects.toBeInstanceOf(ModeConflictError);
        });
    });
});

describe("target validation errors stay structured and specific", () => {
    it("rejects an --original whose suffix is not an Office file", async () => {
        await withTempDir(async (dir) => {
            const notOffice = path.join(dir, "notes.txt");
            writeFileSync(notOffice, "hello");
            const { file } = await controlBytes(dir);
            await expect(validate(file, { original: notOffice })).rejects.toThrow(/must be a \.docx/);
        });
    });

    it("rejects a target that is neither a directory nor an Office file", async () => {
        await withTempDir(async (dir) => {
            const fifo = path.join(dir, "not-a-package.docx");
            // A symlink to a character device passes the existence check and
            // carries a .docx suffix, but is neither a file nor a directory.
            symlinkSync("/dev/null", fifo);
            await expect(validate(fifo)).rejects.toThrow(/is not a directory or Office file/);
        });
    });

    it("rejects an --original that is a directory", async () => {
        await withTempDir(async (dir) => {
            const originalDir = path.join(dir, "original.docx");
            mkdirSync(originalDir);
            const { file } = await controlBytes(dir);
            await expect(validate(file, { original: originalDir })).rejects.toThrow(/is not a file/);
        });
    });

    it("requires an author when --original is given", async () => {
        await withTempDir(async (dir) => {
            const { file } = await controlBytes(dir);
            await expect(validate(file, { original: file })).rejects.toThrow(/`author` is required/);
        });
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
