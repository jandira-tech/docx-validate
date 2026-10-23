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

import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import * as xsdEngine from "../src/lib/xsd-validator";
import { ModeConflictError } from "../src/lib/mode-policy";
import { runValidateFromArgv, validate, type ValidateOptions } from "../src/scripts/office/validate";
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

describe("validate CLI: --help", { timeout: 120_000 }, () => {
    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("succeeds without initializing the XSD engine", async () => {
        const spy = vi.spyOn(xsdEngine, "createXsdValidator");
        const run = await capture(["--help"]);

        expect(run.stdout).toMatch(/Usage:/);
        expect(run.code).toBe(0);
        expect(spy).not.toHaveBeenCalled();
    });

    it("advertises every mode selector", async () => {
        const run = await capture(["--help"]);

        for (const flag of ["--strict", "--quick", "--checks", "--list-checks", "--verify-word", "--json"]) {
            expect(run.stdout).toContain(flag);
        }
    });
});

describe("validate CLI: strict mode", { timeout: 120_000 }, () => {
    it("--strict runs the complete Word check set and the schema walk", async () => {
        const run = await withTempDir(async (dir) => capture([await controlDocx(dir), "--strict"]));

        expect(run.code).toBe(0);
        expect(run.stdout).toContain("All validations PASSED!");
    });
});

describe("validate CLI: --verify-word on an unpacked directory", { timeout: 120_000 }, () => {
    it("reports an unsatisfied Word verification without opening Word", async () => {
        // A directory has no exact bytes to hand Word, so verification is
        // unavailable — and unavailable is not satisfied. No lane opens Word.
        const run = await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            mkdirSync(path.join(unpacked, "word"), { recursive: true });
            writeFileSync(
                path.join(unpacked, "word", "document.xml"),
                '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
            );
            return capture([unpacked, "--original", await controlDocx(dir), "--author", "Test Author", "--verify-word"]);
        });

        expect(run.code).toBe(1);
        expect(run.stderr).toMatch(/WORD VERIFICATION unavailable/);
    });
});

describe("validate CLI: conflicting selectors are rejected", { timeout: 120_000 }, () => {
    const CONFLICTS: readonly (readonly string[])[] = [
        ["--quick", "--checks", "tables"],
        ["--strict", "--quick"],
        ["--strict", "--checks", "tables"],
        ["--strict", "--profile", "lenient"],
        ["--strict", "--profile", "word-valid"],
    ];

    for (const selectors of CONFLICTS) {
        it(`rejects ${selectors.join(" ")} with one stderr line and no exit-zero`, async () => {
            const run = await withTempDir(async (dir) => capture([await controlDocx(dir), ...selectors]));

            expect(run.code).not.toBe(0);
            const lines = run.stderr
                .trimEnd()
                .split("\n")
                .filter((line) => line !== "");
            expect(lines).toHaveLength(1);
            expect(lines[0]).toMatch(/cannot|not.*combin/i);
            expect(run.stdout).toBe("");
        });
    }
});

describe("validate API: conflicting selectors throw a typed error", () => {
    const CONFLICTS: readonly { readonly options: ValidateOptions; readonly message: RegExp }[] = [
        { options: { quick: true, checks: ["tables"] }, message: /quick/i },
        { options: { strict: true, quick: true }, message: /strict/i },
        { options: { strict: true, checks: ["tables"] }, message: /strict/i },
        { options: { strict: true, profile: "lenient" }, message: /strict/i },
        { options: { strict: true, profile: "word-valid" }, message: /strict/i },
    ];

    for (const { options, message } of CONFLICTS) {
        it(`throws ModeConflictError for ${JSON.stringify(options)}`, async () => {
            await withTempDir(async (dir) => {
                const file = await controlDocx(dir);
                const error = await validate(file, options).then(
                    () => undefined,
                    (err: unknown) => err,
                );
                expect(error).toBeInstanceOf(ModeConflictError);
                expect((error as Error).message).toMatch(message);
            });
        });
    }
});
