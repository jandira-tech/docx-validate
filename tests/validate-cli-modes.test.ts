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

describe("validate CLI: named checks", { timeout: 120_000 }, () => {
    it("runs only the named checks and reports the rest as omitted", async () => {
        const run = await withTempDir(async (dir) => capture([await controlDocx(dir), "--checks", "math,tracked-changes"]));

        expect(run.code).toBe(0);
        expect(run.stdout).toContain("Selected checks passed");
        expect(run.stdout).not.toMatch(/all validations pass/i);
        expect(run.stdout).toMatch(/Omitted checks \(not run\): [^\n]*word-shading-color/);
        // The named groups ran: math + tracked-changes checks, plus the
        // xml-syntax prerequisite they share.
        expect(run.stdout).not.toContain("Omitted checks (not run): word-math-run-insertion");
    });

    it("lists every check with its id, groups, formats and description, validating nothing", async () => {
        const run = await capture(["--list-checks"]);

        expect(run.code).toBe(0);
        expect(run.stdout).toContain("xml-syntax");
        expect(run.stdout).toContain("word-row-without-cell");
        expect(run.stdout).toMatch(/groups/i);
        expect(run.stdout).toMatch(/formats/i);
        expect(run.stdout).toMatch(/Every XML part of the package is well-formed/);
        expect(run.stdout).toMatch(/docx/);
    });
});
