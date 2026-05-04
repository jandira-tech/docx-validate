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
 * Port target: src/docx-validate/scripts/office/unpack.py — task #12.
 *
 * Unpack Office files (DOCX, PPTX, XLSX) for editing.
 *
 * Extracts the ZIP archive, pretty-prints XML files, and optionally:
 * - Merges adjacent runs with identical formatting (DOCX only)
 * - Simplifies adjacent tracked changes from same author (DOCX only)
 *
 * CLI:
 *   tsx unpack.ts <office_file> <output_dir> [--merge-runs true|false] [--simplify-redlines true|false]
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";
import JSZip from "jszip";

import { commanderExitCode, runCli } from "../../lib/run-cli.ts";
import { parseXml, prettyXml } from "../../lib/xml-helpers.ts";
import { mergeRuns } from "./helpers/merge-runs.ts";
import { simplifyRedlines } from "./helpers/simplify-redlines.ts";

const SMART_QUOTE_REPLACEMENTS: ReadonlyArray<readonly [string, string]> = [
    ["“", "&#x201C;"],
    ["”", "&#x201D;"],
    ["‘", "&#x2018;"],
    ["’", "&#x2019;"],
];

const SUPPORTED_SUFFIXES = new Set([".docx", ".pptx", ".xlsx"]);

export interface UnpackOptions {
    mergeRuns?: boolean;
    simplifyRedlines?: boolean;
}

export interface UnpackResult {
    message: string;
    ok: boolean;
}

/**
 * Mirror of `unpack(input_file, output_directory, merge_runs=True,
 * simplify_redlines=True)` in `src/docx-validate/scripts/office/unpack.py`.
 *
 * Returns a status object instead of the Python `(None, message)` tuple — the
 * `ok` flag is false whenever the message starts with "Error".
 */
export async function unpack(inputFile: string, outputDir: string, opts: UnpackOptions = {}): Promise<UnpackResult> {
    const doMergeRuns = opts.mergeRuns ?? true;
    const doSimplifyRedlines = opts.simplifyRedlines ?? true;

    const inputPath = path.resolve(inputFile);
    const outputPath = path.resolve(outputDir);
    const suffix = path.extname(inputPath).toLowerCase();

    try {
        await fs.access(inputPath);
    } catch {
        return { ok: false, message: `Error: ${inputFile} does not exist` };
    }

    if (!SUPPORTED_SUFFIXES.has(suffix)) {
        return {
            ok: false,
            message: `Error: ${inputFile} must be a .docx, .pptx, or .xlsx file`,
        };
    }

    try {
        await fs.mkdir(outputPath, { recursive: true });

        const buf = await fs.readFile(inputPath);
        let zip: JSZip;
        try {
            zip = await JSZip.loadAsync(buf);
        } catch {
            return {
                ok: false,
                message: `Error: ${inputFile} is not a valid Office file`,
            };
        }

        await extractAll(zip, outputPath);

        const xmlFiles = await collectXmlFiles(outputPath);
        for (const xmlFile of xmlFiles) {
            await prettyPrintXml(xmlFile);
        }

        let message = `Unpacked ${inputFile} (${xmlFiles.length} XML files)`;

        if (suffix === ".docx") {
            if (doSimplifyRedlines) {
                const simplified = await simplifyRedlines(outputPath);
                if (simplified.message.startsWith("Error:")) {
                    return { ok: false, message: simplified.message };
                }
                message += `, simplified ${simplified.count} tracked changes`;
            }
            if (doMergeRuns) {
                const merged = await mergeRuns(outputPath);
                if (merged.message.startsWith("Error:")) {
                    return { ok: false, message: merged.message };
                }
                message += `, merged ${merged.count} runs`;
            }
        }

        for (const xmlFile of xmlFiles) {
            await escapeSmartQuotes(xmlFile);
        }

        return { ok: true, message };
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        return { ok: false, message: `Error unpacking: ${msg}` };
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

async function prettyPrintXml(xmlFile: string): Promise<void> {
    try {
        const content = await fs.readFile(xmlFile, "utf8");
        const dom = parseXml(content);
        const pretty = prettyXml(dom, "  ");
        await fs.writeFile(xmlFile, pretty, "utf8");
    } catch {
        // Match Python: silently skip files we cannot pretty-print.
    }
}

async function escapeSmartQuotes(xmlFile: string): Promise<void> {
    try {
        let content = await fs.readFile(xmlFile, "utf8");
        for (const [char, entity] of SMART_QUOTE_REPLACEMENTS) {
            content = content.split(char).join(entity);
        }
        await fs.writeFile(xmlFile, content, "utf8");
    } catch {
        // Match Python.
    }
}

function parseBoolFlag(value: string): boolean {
    return value.toLowerCase() === "true";
}

export function buildCommand(): Command {
    const cmd = new Command();
    cmd.name("unpack")
        .description("Unpack an Office file (DOCX, PPTX, XLSX) for editing")
        .argument("<input_file>", "Office file to unpack")
        .argument("<output_directory>", "Output directory")
        .option(
            "--merge-runs <true|false>",
            "Merge adjacent runs with identical formatting (DOCX only, default: true)",
            parseBoolFlag,
            true,
        )
        .option(
            "--simplify-redlines <true|false>",
            "Merge adjacent tracked changes from same author (DOCX only, default: true)",
            parseBoolFlag,
            true,
        );
    return cmd;
}

interface CliOptions {
    mergeRuns: boolean;
    simplifyRedlines: boolean;
}

export async function runFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildCommand();
    cmd.exitOverride();
    // Commander throws CommanderError under exitOverride() for missing args,
    // invalid options, and --help. Catch here so the CLI returns a clean exit
    // code instead of bubbling the error up to runCli().
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [inputFile, outputDir] = cmd.args;

    const result = await unpack(inputFile, outputDir, {
        mergeRuns: opts.mergeRuns,
        simplifyRedlines: opts.simplifyRedlines,
    });
    process.stdout.write(`${result.message}\n`);
    return result.ok ? 0 : 1;
}

runCli(import.meta.url, () => runFromArgv(process.argv.slice(2)));
