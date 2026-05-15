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
 * Port target: src/docx-validate/scripts/office/pack.py — task #13.
 *
 * Pack a directory into a DOCX, PPTX, or XLSX file.
 *
 * Validates with auto-repair, condenses XML formatting, and creates the
 * Office file.
 *
 * CLI:
 *   tsx pack.ts <input_directory> <output_file> [--original <file>] [--validate true|false]
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";
import JSZip from "jszip";

import { commanderExitCode, runCli, withTempDir } from "../../lib/run-cli";
import type { ValidationResult } from "../../lib/types";
import { parseXml, serializeXml } from "../../lib/xml-helpers";
import { inferAuthor } from "./helpers/simplify-redlines";
import { DOCXSchemaValidator } from "./validators/docx";
import { PPTXSchemaValidator } from "./validators/pptx";
import { validateRedlining } from "./validators/redlining";

const SUPPORTED_SUFFIXES = new Set([".docx", ".pptx", ".xlsx"]);
const TEXT_NODE = 3;
const COMMENT_NODE = 8;

export interface PackOptions {
    original?: string;
    validate?: boolean;
    /**
     * Author whose tracked changes to validate against `original`. Required
     * when `validate !== false` and `original` is provided, unless
     * `inferAuthorFunc` is supplied. Used as the fallback when `inferAuthor`
     * cannot deduce a single new author from the diff.
     */
    author?: string;
    /** Custom author-inference callback. When provided, replaces the default `inferAuthor`. Receives (unpackedDir, originalDocx) and returns the author string. */
    inferAuthorFunc?: (unpackedDir: string, originalDocx: string) => Promise<string> | string;
}

export interface PackResult {
    ok: boolean;
    message: string;
    /** Free-form lines emitted during validation (Python printed these to stdout). */
    validationLog?: string;
}

/**
 * Mirror of `pack(input_directory, output_file, original_file=None,
 * validate=True, infer_author_func=None)` in
 * `src/docx-validate/scripts/office/pack.py`. The Python signature returns
 * `(None, message)`; here we return `{ ok, message, validationLog? }`.
 */
export async function pack(inputDirectory: string, outputFile: string, opts: PackOptions = {}): Promise<PackResult> {
    const validate = opts.validate ?? true;
    const original = opts.original;
    const inferAuthorFunc = opts.inferAuthorFunc;

    const inputDir = path.resolve(inputDirectory);
    const outputPath = path.resolve(outputFile);
    const suffix = path.extname(outputPath).toLowerCase();

    let stat;
    try {
        stat = await fs.stat(inputDir);
    } catch {
        return {
            ok: false,
            message: `Error: ${inputDirectory} is not a directory`,
        };
    }
    if (!stat.isDirectory()) {
        return {
            ok: false,
            message: `Error: ${inputDirectory} is not a directory`,
        };
    }

    if (!SUPPORTED_SUFFIXES.has(suffix)) {
        return {
            ok: false,
            message: `Error: ${outputFile} must be a .docx, .pptx, or .xlsx file`,
        };
    }

    let validationLog: string | undefined;

    if (validate) {
        let originalPath: string | undefined;
        if (original) {
            originalPath = path.resolve(original);
            // Fail fast when the caller asked for validation against `original` but
            // the path is missing/unreadable. Silently skipping validation here
            // would mask operator error.
            try {
                await fs.access(originalPath);
            } catch {
                return {
                    ok: false,
                    message: `Error: original file not found: ${original}`,
                };
            }
        }
        const { success, log } = await runValidation(inputDir, originalPath, suffix, inferAuthorFunc, opts.author);
        if (log) {
            validationLog = log;
        }
        if (!success) {
            return {
                ok: false,
                message: `Error: Validation failed for ${inputDirectory}`,
                validationLog,
            };
        }
    }

    await withTempDir(async (tempDir) => {
        const tempContentDir = path.join(tempDir, "content");
        await copyDir(inputDir, tempContentDir);

        const xmlFiles = await collectXmlFiles(tempContentDir);
        for (const xmlFile of xmlFiles) {
            await condenseXml(xmlFile);
        }

        await fs.mkdir(path.dirname(outputPath), { recursive: true });

        const zip = new JSZip();
        const allFiles = await collectAllFiles(tempContentDir);
        for (const filePath of allFiles) {
            const rel = path.relative(tempContentDir, filePath).split(path.sep).join("/");
            const data = await fs.readFile(filePath);
            zip.file(rel, data);
        }
        const buf = await zip.generateAsync({
            type: "nodebuffer",
            compression: "DEFLATE",
        });
        await fs.writeFile(outputPath, buf);
    });

    return {
        ok: true,
        message: `Successfully packed ${inputDirectory} to ${outputFile}`,
        validationLog,
    };
}

interface ValidatorRunner {
    repair(): Promise<number>;
    runChecks(): Promise<ValidationResult>;
}

async function runValidation(
    unpackedDir: string,
    originalFile: string | undefined,
    suffix: string,
    inferAuthorFunc?: PackOptions["inferAuthorFunc"],
    explicitAuthor?: string,
): Promise<{ success: boolean; log?: string }> {
    const lines: string[] = [];
    const validators: ValidatorRunner[] = [];

    if (suffix === ".docx") {
        const docx = new DOCXSchemaValidator({ unpackedDir, originalFile });
        validators.push({
            repair: () => docx.repair(),
            runChecks: () => docx.validate(),
        });

        // Only run redlining validation and author inference when originalFile is provided
        if (originalFile) {
            let author: string;
            try {
                const fn = inferAuthorFunc ?? ((dir, orig) => inferAuthor(dir, orig, explicitAuthor ?? ""));
                author = await fn(unpackedDir, originalFile);
                if (!author) {
                    throw new Error("Could not infer author and `author` was not provided.");
                }
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                process.stderr.write(`Error: ${message}\n`);
                return { success: false, log: lines.join("\n") };
            }

            validators.push({
                repair: async () => 0,
                runChecks: () => validateRedlining({ unpackedDir, originalDocx: originalFile, author }),
            });
        }
    } else if (suffix === ".pptx") {
        const pptx = new PPTXSchemaValidator({ unpackedDir, originalFile });
        validators.push({
            repair: () => pptx.repair(),
            runChecks: () => pptx.validate(),
        });
    }

    if (validators.length === 0) {
        return { success: true };
    }

    let totalRepairs = 0;
    for (const v of validators) {
        totalRepairs += await v.repair();
    }
    if (totalRepairs > 0) {
        lines.push(`Auto-repaired ${totalRepairs} issue(s)`);
    }

    let success = true;
    for (const v of validators) {
        const result = await v.runChecks();
        if (!result.valid) {
            success = false;
        }
    }

    if (success) {
        lines.push("All validations PASSED!");
    }

    return { success, log: lines.length > 0 ? lines.join("\n") : undefined };
}

async function copyDir(src: string, dest: string): Promise<void> {
    await fs.mkdir(dest, { recursive: true });
    const entries = await fs.readdir(src, { withFileTypes: true });
    for (const entry of entries) {
        const from = path.join(src, entry.name);
        const to = path.join(dest, entry.name);
        if (entry.isDirectory()) {
            await copyDir(from, to);
        } else if (entry.isFile()) {
            await fs.copyFile(from, to);
        } else if (entry.isSymbolicLink()) {
            const target = await fs.readlink(from);
            await fs.symlink(target, to);
        }
    }
}

async function collectXmlFiles(root: string): Promise<string[]> {
    const out: string[] = [];
    const walk = async (dir: string): Promise<void> => {
        const entries = await fs.readdir(dir, { withFileTypes: true });
        for (const entry of entries) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                await walk(full);
            } else if (entry.isFile()) {
                const ext = path.extname(entry.name).toLowerCase();
                if (ext === ".xml" || ext === ".rels") {
                    out.push(full);
                }
            }
        }
    };
    await walk(root);
    return out;
}

async function collectAllFiles(root: string): Promise<string[]> {
    const out: string[] = [];
    const walk = async (dir: string): Promise<void> => {
        const entries = await fs.readdir(dir, { withFileTypes: true });
        for (const entry of entries) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                await walk(full);
            } else if (entry.isFile()) {
                out.push(full);
            }
        }
    };
    await walk(root);
    return out;
}

async function condenseXml(xmlFile: string): Promise<void> {
    let dom: Document;
    try {
        const content = await fs.readFile(xmlFile, "utf8");
        dom = parseXml(content);
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        process.stderr.write(`ERROR: Failed to parse ${path.basename(xmlFile)}: ${message}\n`);
        throw err;
    }

    const all = dom.getElementsByTagName("*");
    for (let i = 0; i < all.length; i += 1) {
        const element = all.item(i);
        if (!element) continue;
        // Preserve text-bearing <t> elements regardless of prefix; unprefixed
        // <t> is common in SpreadsheetML, and stripping whitespace-only text
        // children from those nodes silently changes content.
        const tagName = element.tagName;
        const localName = tagName.includes(":") ? tagName.slice(tagName.indexOf(":") + 1) : tagName;
        if (localName === "t") continue;

        const toRemove: Node[] = [];
        for (let child = element.firstChild; child; child = child.nextSibling) {
            if (child.nodeType === COMMENT_NODE) {
                toRemove.push(child);
                continue;
            }
            if (child.nodeType === TEXT_NODE && (child.nodeValue ?? "").trim() === "") {
                toRemove.push(child);
            }
        }
        for (const child of toRemove) {
            element.removeChild(child);
        }
    }

    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf8");
}

function parseBoolFlag(value: string): boolean {
    return value.toLowerCase() === "true";
}

export function buildPackCommand(): Command {
    const cmd = new Command();
    cmd.name("pack")
        .description("Pack a directory into a DOCX, PPTX, or XLSX file")
        .argument("<input_directory>", "Unpacked Office document directory")
        .argument("<output_file>", "Output Office file (.docx/.pptx/.xlsx)")
        .option("--original <file>", "Original file for validation comparison")
        .option("--validate <true|false>", "Run validation with auto-repair (default: true)", parseBoolFlag, true)
        .option(
            "--author <name>",
            "Author whose tracked changes to validate (used as inferAuthor fallback). Required when --original is provided.",
        );
    return cmd;
}

interface CliOptions {
    original?: string;
    validate: boolean;
    author?: string;
}

export async function runPackFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildPackCommand();
    cmd.exitOverride();
    // Catch CommanderError (missing args, invalid options, --help) so the CLI
    // returns a clean exit code instead of bubbling to runCli().
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [inputDir, outputFile] = cmd.args;

    const result = await pack(inputDir, outputFile, {
        original: opts.original,
        validate: opts.validate,
        author: opts.author,
    });
    if (result.validationLog) {
        process.stdout.write(`${result.validationLog}\n`);
    }
    process.stdout.write(`${result.message}\n`);
    return result.ok ? 0 : 1;
}

runCli(import.meta.url, () => runPackFromArgv(process.argv.slice(2)));
