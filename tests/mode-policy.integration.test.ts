/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Mode policy (docs/plans/2026-10-23-fixing-validator.md, tasks 10 and 11).
 *
 * Contract from plan § 1: `strict` succeeds only when the complete Word check
 * set passed AND schema validation completed. Every result therefore reports
 * execution metadata — profile, executed/omitted checks, schema completion —
 * so callers can tell a real success from a partial one.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { XsdValidator } from "../src/lib/xsd-validator";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken");

const errorCodes = (issues: readonly { severity: string; code?: string }[]): string[] => [
    ...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "")),
];

describe("strict policy: success implies complete Word checks and completed schema validation", { timeout: 120_000 }, () => {
    it("reports a full strict run as complete, with the schema walk finished", async () => {
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { profile: "strict" });
        });

        expect(result.valid).toBe(true);
        expect(result.metadata?.profile).toBe("strict");
        expect(result.metadata?.omittedChecks).toEqual([]);
        expect(result.metadata?.executedChecks?.length).toBeGreaterThan(0);
        expect(result.metadata?.schemaCompleted).toBe(true);
    });

    it("never reports a strict success when a Word check fails", async () => {
        const result = await validate(path.join(BROKEN, "package.word-row-without-cell.docx"), { profile: "strict" });

        expect(errorCodes(result.issues)).toContain("word-row-without-cell");
        expect(result.valid).toBe(false);
        expect(result.metadata?.profile).toBe("strict");
        expect(result.metadata?.executedChecks).toContain("word-row-without-cell");
        expect(result.metadata?.omittedChecks).toEqual([]);
    });

    it("never reports a strict success when the schema engine cannot load a schema", async () => {
        const failing: XsdValidator = {
            validate: async (_xml: string, schemaPath: string) => [
                { severity: "error", code: "xsd-schema-load-failed", message: `Schema load failed (${schemaPath}): injected` },
            ],
        };
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { profile: "strict", xsdValidator: failing });
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.schemaCompleted).toBe(false);
    });
});
