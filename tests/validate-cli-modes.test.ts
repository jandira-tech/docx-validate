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
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { runValidateFromArgv, validate, type ValidateOptions } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

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
        expect(run.stdout).toContain("word-document-root");
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

describe("validate CLI: rejected invocations", { timeout: 120_000 }, () => {
    it("exits 1 with one line when no path is given", async () => {
        const run = await capture([]);

        expect(run.code).toBe(1);
        expect(run.stderr.trim()).toMatch(/missing required argument 'path'/);
    });

    it("exits 1 with one line for an unknown --profile value", async () => {
        const run = await withTempDir(async (dir) => capture([await controlDocx(dir), "--profile", "picky"]));

        expect(run.code).toBe(1);
        expect(run.stderr.trim()).toMatch(/Invalid --profile: picky/);
    });

    it("exits 1 with one line when the target does not exist", async () => {
        const run = await withTempDir(async (dir) => capture([path.join(dir, "missing.docx")]));

        expect(run.code).toBe(1);
        expect(run.stderr.trim()).toMatch(/does not exist/);
    });

    it("exits 1 with one line when the engine probe fails", async () => {
        const probe = vi
            .spyOn(BaseSchemaValidator, "assertXsdValidationAvailable")
            .mockRejectedValue(new Error("XSD engine unavailable (libxml2-wasm): boom"));
        try {
            // `--profile strict` walks the schemas, so it needs the engine and
            // must fail loudly when the engine is broken (#318).
            const run = await withTempDir(async (dir) => capture([await controlDocx(dir), "--profile", "strict"]));
            expect(run.code).toBe(1);
            expect(run.stderr.trim()).toBe("XSD engine unavailable (libxml2-wasm): boom");
        } finally {
            probe.mockRestore();
        }
    });

    it("never probes the XSD engine for the default word-valid DOCX run (#318)", async () => {
        const probe = vi
            .spyOn(BaseSchemaValidator, "assertXsdValidationAvailable")
            .mockRejectedValue(new Error("XSD engine unavailable (libxml2-wasm): boom"));
        try {
            // The default DOCX profile is word-valid, which compiles no
            // schema: a broken engine must not fail a run that never uses it.
            const run = await withTempDir(async (dir) => capture([await controlDocx(dir)]));
            expect(probe).not.toHaveBeenCalled();
            expect(run.code).toBe(0);
            expect(run.stdout).toContain("All validations PASSED!");
            expect(run.stderr).toBe("");
        } finally {
            probe.mockRestore();
        }
    });

    it("reports the repair count on stdout when --auto-repair changes something", async () => {
        const run = await withTempDir(async (dir) => {
            const file = path.join(dir, "leading-space.docx");
            // A <w:t> with leading whitespace and no xml:space="preserve" is
            // the smallest thing the repairer rewrites.
            writeFileSync(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:t>Hello</w:t>", "<w:t> Hello</w:t>"),
                }),
            );
            return capture([file, "--auto-repair"]);
        });

        expect(run.stdout).toMatch(/Auto-repaired [1-9]\d* issue\(s\)/);
    });

    it("exits 1 with one line when --original is given without --author", async () => {
        const run = await withTempDir(async (dir) => {
            const file = await controlDocx(dir);
            return capture([file, "--original", file]);
        });

        expect(run.code).toBe(1);
        expect(run.stderr.trim()).toMatch(/`author` is required/);
    });

    it("keeps the library silent: validate() writes nothing to stdout or stderr", async () => {
        const out: string[] = [];
        const err: string[] = [];
        const stdoutWrite = process.stdout.write.bind(process.stdout);
        const stderrWrite = process.stderr.write.bind(process.stderr);
        process.stdout.write = ((chunk: string | Uint8Array): boolean => {
            out.push(typeof chunk === "string" ? chunk : "");
            return true;
        }) as typeof process.stdout.write;
        process.stderr.write = ((chunk: string | Uint8Array): boolean => {
            err.push(typeof chunk === "string" ? chunk : "");
            return true;
        }) as typeof process.stderr.write;
        try {
            await withTempDir(async (dir) => {
                const file = await controlDocx(dir);
                await validate(file);
                await validate(file, { quick: true });
            });
        } finally {
            process.stdout.write = stdoutWrite;
            process.stderr.write = stderrWrite;
        }
        expect(out.join("")).toBe("");
        expect(err.join("")).toBe("");
    });
});

describe("validate CLI: --json", { timeout: 120_000 }, () => {
    const SPECIMEN = path.join(path.dirname(new URL(import.meta.url).pathname), "lane", "jubarte-0.13.0", "k06_1.redline.docx");

    it("prints exactly one JSON document with every issue and the execution metadata", async () => {
        const run = await capture([SPECIMEN, "--json"]);

        expect(run.code).toBe(0);
        // Parses only if nothing else shares stdout.
        const doc = JSON.parse(run.stdout) as {
            valid: boolean;
            issues: { code?: string; severity: string; path?: string }[];
            metadata: { profile?: string; executedChecks?: string[]; omittedChecks?: string[]; schemaCompleted?: boolean };
        };
        expect(doc.valid).toBe(true);
        // Warnings, and the part each one came from, are present.
        const warnings = doc.issues.filter((issue) => issue.code === "word-shading-color");
        expect(warnings.length).toBeGreaterThan(0);
        expect(warnings.every((issue) => issue.severity === "warning" && typeof issue.path === "string")).toBe(true);
        expect(doc.metadata.profile).toBe("word-valid");
        expect(doc.metadata.executedChecks?.length).toBeGreaterThan(0);
        expect(doc.metadata.omittedChecks).toEqual([]);
        // The default DOCX profile, word-valid, compiles no schema (#310).
        expect(doc.metadata.schemaCompleted).toBe(false);
        // No human text.
        expect(run.stdout).not.toMatch(/All validations|Selected checks|WORD VERIFICATION/);
    });

    it("reports failures as JSON too, with no human text on stdout", async () => {
        const run = await capture([path.join(BROKEN, "package.word-row-without-cell.docx"), "--json"]);

        expect(run.code).toBe(1);
        const doc = JSON.parse(run.stdout) as { valid: boolean; issues: { code?: string }[] };
        expect(doc.valid).toBe(false);
        expect(doc.issues.some((issue) => issue.code === "word-row-without-cell")).toBe(true);
        expect(run.stdout).not.toMatch(/All validations|Selected checks/);
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
