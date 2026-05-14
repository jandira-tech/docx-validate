/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Port target: src/docx-validate/scripts/office/validate.py — task #15.
 *
 * Top-level validation CLI. Dispatches to the DOCX or PPTX schema validator
 * (and the redlining validator for DOCX with --original) based on the file
 * extension of either the unpacked target or the supplied --original.
 *
 * CLI:
 *   tsx validate.ts <path> [--original <file>] [-v|--verbose] [--auto-repair]
 *                          [--author <name>]
 *
 * `<path>` may be either an unpacked directory or a packed `.docx`/`.pptx`/
 * `.xlsx` file (the latter is unzipped to a temp dir for the duration of the
 * run). Exit code: 0 when every check passes, 1 otherwise.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";
import JSZip from "jszip";

import { commanderExitCode, runCli, withTempDir } from "../../lib/run-cli";
import { DEFAULT_PROFILE, mergeResults, type Profile, type ValidationResult } from "../../lib/types";
import { BaseSchemaValidator } from "./validators/base";
import { DOCXSchemaValidator } from "./validators/docx";
import { PPTXSchemaValidator } from "./validators/pptx";
import { validateRedlining } from "./validators/redlining";

const SUPPORTED_SUFFIXES = new Set([".docm", ".docx", ".pptx", ".xlsx"]);

export interface ValidateOptions {
    /** Path to the original packed Office file. Required for DOCX redlining. */
    original?: string;
    /** Run auto-repair before validating (mirrors `--auto-repair`). */
    autoRepair?: boolean;
    /** Author name to use for redlining validation. */
    author?: string;
    /** Verbose mode — passes through to subclasses' verbose flags. */
    verbose?: boolean;
    /** Override XSD root (defaults to the path used by `BaseSchemaValidator`). */
    schemasDir?: string;
    /**
     * Validation profile. Defaults to `"lenient"` to match real-world
     * Microsoft Office output. Pass `"strict"` for spec-purist behaviour
     * (flags BOM-prefixed parts and other tolerated-but-non-canonical
     * constructs). See {@link Profile}.
     */
    profile?: Profile;
}

export interface ValidateRunResult extends ValidationResult {
    /** Number of issues auto-repaired (0 unless `autoRepair` is true). */
    repairs: number;
    /** Detected suffix used for dispatch (e.g. ".docx"). */
    suffix: string;
}

/**
 * Public entry point — mirrors `main()` in
 * `src/docx-validate/scripts/office/validate.py` minus the argv parsing. Use
 * this from other TS code; the CLI shim at the bottom of the file just turns
 * argv into options and exits with the right code.
 */
export async function validate(target: string, opts: ValidateOptions = {}): Promise<ValidateRunResult> {
    const author = opts.author;
    const verbose = opts.verbose ?? false;
    const autoRepair = opts.autoRepair ?? false;
    const original = opts.original;
    const profile: Profile = opts.profile ?? DEFAULT_PROFILE;

    await assertExists(target, `Error: ${target} does not exist`);

    let originalFile: string | null = null;
    if (original) {
        await assertIsFile(original, `Error: ${original} is not a file`);
        const ext = path.extname(original).toLowerCase();
        if (!SUPPORTED_SUFFIXES.has(ext)) {
            throw new Error(`Error: ${original} must be a .docx, .pptx, or .xlsx file`);
        }
        originalFile = path.resolve(original);
    }

    const dispatchSuffix = path.extname(originalFile ?? target).toLowerCase();

    if (!SUPPORTED_SUFFIXES.has(dispatchSuffix)) {
        throw new Error(`Error: Cannot determine file type from ${target}. Use --original or provide a .docx/.pptx/.xlsx file.`);
    }

    const targetStat = await fs.stat(target);
    const targetSuffix = path.extname(target).toLowerCase();
    const targetIsPackedFile = targetStat.isFile() && SUPPORTED_SUFFIXES.has(targetSuffix);

    if (!targetIsPackedFile && !targetStat.isDirectory()) {
        throw new Error(`Error: ${target} is not a directory or Office file`);
    }

    const runWithUnpacked = async (unpackedDir: string): Promise<ValidateRunResult> => {
        const subclassResult = await runValidators(unpackedDir, {
            originalFile,
            suffix: dispatchSuffix,
            author,
            verbose,
            autoRepair,
            schemasDir: opts.schemasDir,
            profile,
        });
        return { ...subclassResult, suffix: dispatchSuffix };
    };

    if (targetIsPackedFile) {
        return withTempDir(async (tempDir) => {
            const buf = await fs.readFile(target);
            const zip = await JSZip.loadAsync(buf);
            await extractAll(zip, tempDir);
            return runWithUnpacked(tempDir);
        });
    }

    return runWithUnpacked(path.resolve(target));
}

interface ValidatorRunner {
    repair(): Promise<number>;
    validate(): Promise<ValidationResult>;
}

interface RunValidatorsOptions {
    originalFile: string | null;
    suffix: string;
    /** Required iff `originalFile` is set (DOCX redlining cross-check). */
    author: string | undefined;
    verbose: boolean;
    autoRepair: boolean;
    schemasDir?: string;
    profile: Profile;
}

async function runValidators(unpackedDir: string, opts: RunValidatorsOptions): Promise<ValidationResult & { repairs: number }> {
    const validators: ValidatorRunner[] = [];

    if (opts.suffix === ".docx") {
        const docx = new DOCXSchemaValidator({
            unpackedDir,
            originalFile: opts.originalFile ?? undefined,
            verbose: opts.verbose,
            schemasDir: opts.schemasDir,
            profile: opts.profile,
        });
        validators.push(docx);
        if (opts.originalFile) {
            if (!opts.author) {
                throw new Error(
                    "validate(): `author` is required when `original` is provided (used to identify whose tracked changes to verify).",
                );
            }
            const originalDocx = opts.originalFile;
            const author = opts.author;
            const verbose = opts.verbose;
            validators.push({
                repair: async () => 0,
                validate: () =>
                    validateRedlining({
                        unpackedDir,
                        originalDocx,
                        author,
                        verbose,
                    }),
            });
        }
    } else if (opts.suffix === ".pptx") {
        const pptx = new PPTXSchemaValidator({
            unpackedDir,
            originalFile: opts.originalFile ?? undefined,
            verbose: opts.verbose,
            schemasDir: opts.schemasDir,
            profile: opts.profile,
        });
        validators.push(pptx);
    } else {
        // Library code is silent (CLAUDE.md: "Validator results are
        // structured, not printed"). The CLI shim renders this issue via
        // its general issue-printing path at the bottom of
        // runValidateFromArgv, so callers driving validate() directly get
        // the structured error without stderr noise.
        return {
            valid: false,
            issues: [
                {
                    severity: "error" as const,
                    message: `Unsupported file type: ${opts.suffix}`,
                    code: "unsupported-file-type",
                },
            ],
            repairs: 0,
        };
    }

    let repairs = 0;
    if (opts.autoRepair) {
        for (const v of validators) {
            repairs += await v.repair();
        }
    }

    const results = await Promise.all(validators.map((v) => v.validate()));
    const merged = mergeResults(...results);
    return { ...merged, repairs };
}

async function assertExists(p: string, message: string): Promise<void> {
    try {
        await fs.access(p);
    } catch {
        throw new Error(message);
    }
}

async function assertIsFile(p: string, message: string): Promise<void> {
    let stat: Awaited<ReturnType<typeof fs.stat>>;
    try {
        stat = await fs.stat(p);
    } catch {
        throw new Error(message);
    }
    if (!stat.isFile()) {
        throw new Error(message);
    }
}

async function extractAll(zip: JSZip, outputPath: string): Promise<void> {
    const entries: Array<{ name: string; file: JSZip.JSZipObject }> = [];
    zip.forEach((relativePath, file) => {
        entries.push({ name: relativePath, file });
    });

    for (const { name, file } of entries) {
        const target = path.join(outputPath, name);
        const resolved = path.resolve(target);
        if (!resolved.startsWith(`${outputPath}${path.sep}`) && resolved !== outputPath) {
            throw new Error(`Refusing to extract entry outside output dir: ${name}`);
        }
        if (file.dir) {
            await fs.mkdir(resolved, { recursive: true });
            continue;
        }
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        const data = await file.async("nodebuffer");
        await fs.writeFile(resolved, data);
    }
}

interface CliOptions {
    original?: string;
    verbose: boolean;
    autoRepair: boolean;
    author?: string;
    profile: Profile;
}

export function buildValidateCommand(): Command {
    const cmd = new Command();
    cmd.name("validate")
        .description("Validate Office document XML files against XSD schemas and tracked changes")
        .argument("<path>", "Path to unpacked directory or packed Office file (.docx/.pptx/.xlsx)")
        .option(
            "--original <file>",
            "Path to original file (.docx/.pptx/.xlsx). If omitted, all XSD errors are reported and redlining validation is skipped.",
        )
        .option("-v, --verbose", "Enable verbose output", false)
        .option("--auto-repair", "Automatically repair common issues (hex IDs, whitespace preservation)", false)
        .option("--author <name>", "Author name for redlining validation (required when --original is provided)")
        .option(
            "--profile <profile>",
            "Validation profile: 'lenient' (default; tolerates real-world Office output) or 'strict' (spec-purist; flags BOMs and similar)",
            DEFAULT_PROFILE,
        );
    return cmd;
}

export async function runValidateFromArgv(argv: readonly string[]): Promise<number> {
    // Fail loudly at startup if libxmljs2's native binding is broken — otherwise
    // the per-file pipeline silently turns the same condition into per-file
    // "Invalid XSD schema" errors that look like document corruption.
    try {
        BaseSchemaValidator.assertLibxmljsAvailable();
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        process.stderr.write(`${message}\n`);
        return 1;
    }

    const cmd = buildValidateCommand();
    cmd.exitOverride();
    // Commander throws CommanderError under exitOverride() for missing args,
    // invalid options, and --help. Catch here so they bypass the validator
    // try/catch below (which would otherwise format them as validate() errors).
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [target] = cmd.args;

    if (opts.profile !== "lenient" && opts.profile !== "strict") {
        const bad = String(opts.profile);
        process.stderr.write(`Invalid --profile: ${bad}. Must be 'lenient' or 'strict'.\n`);
        return 1;
    }

    let result: ValidateRunResult;
    try {
        result = await validate(target, {
            original: opts.original,
            autoRepair: opts.autoRepair,
            author: opts.author,
            verbose: opts.verbose,
            profile: opts.profile,
        });
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        process.stderr.write(`${message}\n`);
        return 1;
    }

    if (opts.autoRepair && result.repairs > 0) {
        process.stdout.write(`Auto-repaired ${result.repairs} issue(s)\n`);
    }

    if (result.valid) {
        process.stdout.write("All validations PASSED!\n");
    } else {
        for (const issue of result.issues) {
            if (issue.severity !== "error") continue;
            const where = issue.path ? ` [${issue.path}]` : "";
            process.stderr.write(`${issue.severity.toUpperCase()}${where}: ${issue.message}\n`);
        }
    }

    return result.valid ? 0 : 1;
}

runCli(import.meta.url, () => runValidateFromArgv(process.argv.slice(2)));
