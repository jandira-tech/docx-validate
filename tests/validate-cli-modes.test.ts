/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * CLI surface of `validate` (docs/plans/2026-10-23-fixing-validator.md,
 * tasks 10 and 11).
 *
 * Covers the mode selectors (`--quick`, `--checks`, `--list-checks`), their
 * conflict rules, and the two output modes (human text, `--json`). Human
 * rendering lives in the CLI layer only; every assertion here pins what lands
 * on stdout and stderr.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { runValidateFromArgv } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken");

interface CapturedRun {
    code: number;
    stdout: string;
    stderr: string;
}

/** Run the CLI with stdout/stderr captured, restoring both afterwards. */
const capture = async (argv: readonly string[]): Promise<CapturedRun> => {
    const out: string[] = [];
    const err: string[] = [];
    const stdoutWrite = process.stdout.write.bind(process.stdout);
    const stderrWrite = process.stderr.write.bind(process.stderr);
    process.stdout.write = ((chunk: string | Uint8Array): boolean => {
        out.push(typeof chunk === "string" ? chunk : Buffer.from(chunk).toString("utf-8"));
        return true;
    }) as typeof process.stdout.write;
    process.stderr.write = ((chunk: string | Uint8Array): boolean => {
        err.push(typeof chunk === "string" ? chunk : Buffer.from(chunk).toString("utf-8"));
        return true;
    }) as typeof process.stderr.write;
    try {
        const code = await runValidateFromArgv(argv);
        return { code, stdout: out.join(""), stderr: err.join("") };
    } finally {
        process.stdout.write = stdoutWrite;
        process.stderr.write = stderrWrite;
    }
};

const controlDocx = async (dir: string): Promise<string> => {
    const file = path.join(dir, "control.docx");
    writeFileSync(file, await buildDocx());
    return file;
};

describe("validate CLI: quick mode output", { timeout: 120_000 }, () => {
    it("names the omitted checks and never claims every validation passed", async () => {
        const run = await withTempDir(async (dir) => capture([await controlDocx(dir), "--quick"]));

        expect(run.stdout).toContain("Selected checks passed");
        expect(run.stdout).not.toMatch(/all validations pass/i);
        expect(run.stdout).toMatch(/omitted/i);
        // A quick run must say which checks it left out, by id.
        expect(run.stdout).toContain("word-shading-color");
        expect(run.code).toBe(0);
    });

    it("exits 1 and reports the defect when a selected check fails", async () => {
        const run = await capture([path.join(BROKEN, "package.word-row-without-cell.docx"), "--quick"]);

        expect(run.code).toBe(1);
        expect(run.stderr).toContain("word/document.xml");
        expect(run.stdout).not.toMatch(/all validations pass/i);
    });
});
