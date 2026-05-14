/*
 * Regenerate the Word-invalid fixture corpus through the current DOCX repair
 * pipeline, preserving enough metadata to compare against Microsoft Word.
 *
 * This script deliberately keeps the transformation narrow:
 *   1. unpack without merge-runs / redline simplification;
 *   2. run DOCXSchemaValidator.repair();
 *   3. pack without extra validation;
 *   4. compare text-bearing Word parts before/after.
 */

import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync, appendFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import JSZip from "jszip";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { parseXml } from "../src/lib/xml-helpers";
import { pack } from "../src/scripts/office/pack";
import { unpack } from "../src/scripts/office/unpack";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const DEFAULT_SOURCE_ROOT = path.join(REPO, "tests/fixtures/word-regenerate-invalid/original");
const DEFAULT_OUTPUT_ROOT = path.join(REPO, "tests/fixtures/word-regenerate-invalid/regenerated");
const DEFAULT_LOG = path.join(REPO, "tests/fixtures/word-regenerate-invalid/regeneration-results.jsonl");

interface CliOptions {
    sourceRoot: string;
    outputRoot: string;
    log: string;
    match?: string;
    limit?: string;
    resume: boolean;
}

interface TextSignature {
    ok: boolean;
    text: string;
    error?: string;
}

interface RegenerationRecord {
    relativePath: string;
    source: string;
    output: string;
    sourceSize: number;
    outputSize: number | null;
    unpackOk: boolean;
    unpackMessage: string;
    repairs: number;
    packOk: boolean;
    packMessage: string;
    contentSame: boolean | null;
    contentError?: string;
    validator: {
        valid: boolean;
        errorCodes: string[];
        threw?: string;
    } | null;
}

function walkDocx(dir: string, out: string[]): void {
    for (const entry of readdirSync(dir)) {
        if (entry.startsWith("~$")) continue;
        const file = path.join(dir, entry);
        const stat = statSync(file);
        if (stat.isDirectory()) {
            walkDocx(file, out);
        } else if (entry.toLowerCase().endsWith(".docx")) {
            out.push(file);
        }
    }
}

function buildCommand(): Command {
    return new Command()
        .name("regenerate-word-invalid-fixtures")
        .description("Run the current DOCX repair/regenerate pipeline over the Word-invalid fixture corpus")
        .option("--source-root <dir>", "Root containing copied invalid DOCX fixtures", DEFAULT_SOURCE_ROOT)
        .option("--output-root <dir>", "Root where regenerated DOCX files are written", DEFAULT_OUTPUT_ROOT)
        .option("--log <file>", "JSONL log path", DEFAULT_LOG)
        .option("--match <text>", "Only regenerate files whose relative path contains this text")
        .option("--limit <n>", "Maximum number of files to regenerate")
        .option("--resume", "Skip records already present in the JSONL log", false);
}

function parseOptions(argv: readonly string[]): CliOptions {
    const command = buildCommand();
    command.parse(argv as string[], { from: "user" });
    return command.opts<CliOptions>();
}

function completedRecords(logFile: string): Set<string> {
    if (!existsSync(logFile)) return new Set();
    const done = new Set<string>();
    for (const line of readFileSync(logFile, "utf-8").split(/\r?\n/)) {
        if (!line.trim()) continue;
        try {
            const parsed = JSON.parse(line) as Partial<RegenerationRecord>;
            if (parsed.relativePath) done.add(parsed.relativePath);
        } catch {
            // Allow resuming after an interrupted / partial write.
        }
    }
    return done;
}

async function textSignature(docxPath: string): Promise<TextSignature> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(readFileSync(docxPath));
    } catch (err) {
        return { ok: false, text: "", error: err instanceof Error ? err.message : String(err) };
    }

    const parts: string[] = [];
    zip.forEach((relativePath, entry) => {
        if (entry.dir) return;
        if (!relativePath.startsWith("word/")) return;
        if (relativePath.includes("/_rels/")) return;
        if (!relativePath.endsWith(".xml")) return;
        parts.push(relativePath);
    });
    parts.sort();

    const chunks: string[] = [];
    for (const part of parts) {
        const entry = zip.file(part);
        if (!entry) continue;
        let xml: string;
        try {
            xml = await entry.async("text");
        } catch (err) {
            return { ok: false, text: "", error: `${part}: ${err instanceof Error ? err.message : String(err)}` };
        }
        try {
            const dom = parseXml(xml);
            collectTextBearingContent(dom.documentElement, chunks);
        } catch (err) {
            return { ok: false, text: "", error: `${part}: ${err instanceof Error ? err.message : String(err)}` };
        }
    }

    return { ok: true, text: chunks.join("") };
}

function collectTextBearingContent(node: Node | null, chunks: string[]): void {
    if (!node) return;
    if (node.nodeType !== 1) {
        return;
    }
    const elem = node as Element;
    const local = elem.localName || elem.tagName.split(":").pop() || elem.tagName;
    if (local === "t" || local === "delText" || local === "instrText") {
        chunks.push(textContent(elem));
    } else if (local === "tab") {
        chunks.push("\t");
    } else if (local === "br" || local === "cr") {
        chunks.push("\n");
    }

    for (let child = elem.firstChild; child; child = child.nextSibling) {
        collectTextBearingContent(child, chunks);
    }
}

function textContent(elem: Element): string {
    let out = "";
    for (let child = elem.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 3) {
            out += (child as Text).data ?? child.nodeValue ?? "";
        }
    }
    return out;
}

async function runValidator(file: string): Promise<RegenerationRecord["validator"]> {
    try {
        const result = await validate(file, { profile: "word-valid" });
        const errorCodes = Array.from(
            new Set(
                result.issues
                    .filter((issue: ValidationIssue) => issue.severity === "error")
                    .map((issue: ValidationIssue) => issue.code)
                    .filter((code): code is string => Boolean(code)),
            ),
        ).sort();
        return { valid: result.valid, errorCodes };
    } catch (err) {
        return {
            valid: false,
            errorCodes: [],
            threw: err instanceof Error ? err.message.slice(0, 240) : String(err).slice(0, 240),
        };
    }
}

async function regenerateOne(source: string, output: string, relativePath: string): Promise<RegenerationRecord> {
    const sourceSize = statSync(source).size;
    let outputSize: number | null = null;
    let unpackOk = false;
    let unpackMessage = "";
    let repairs = 0;
    let packOk = false;
    let packMessage = "";

    await withTempDir(async (dir) => {
        const unpacked = path.join(dir, "unpacked");
        const unpackedResult = await unpack(source, unpacked, {
            mergeRuns: false,
            simplifyRedlines: false,
        });
        unpackOk = unpackedResult.ok;
        unpackMessage = unpackedResult.message;
        if (!unpackedResult.ok) return;

        try {
            const validator = new DOCXSchemaValidator({ unpackedDir: unpacked, profile: "word-valid" });
            repairs = await validator.repair();
        } catch (err) {
            unpackMessage = `repair failed after unpack: ${err instanceof Error ? err.message : String(err)}`;
            return;
        }

        mkdirSync(path.dirname(output), { recursive: true });
        try {
            const packResult = await pack(unpacked, output, { validate: false });
            packOk = packResult.ok;
            packMessage = packResult.message;
            if (packOk) outputSize = statSync(output).size;
        } catch (err) {
            packOk = false;
            packMessage = `pack threw: ${err instanceof Error ? err.message : String(err)}`;
        }
    });

    let contentSame: boolean | null = null;
    let contentError: string | undefined;
    if (packOk) {
        const before = await textSignature(source);
        const after = await textSignature(output);
        if (before.ok && after.ok) {
            contentSame = before.text === after.text;
        } else {
            contentSame = null;
            contentError = before.error ?? after.error;
        }
    }

    return {
        relativePath,
        source,
        output,
        sourceSize,
        outputSize,
        unpackOk,
        unpackMessage,
        repairs,
        packOk,
        packMessage,
        contentSame,
        contentError,
        validator: packOk ? await runValidator(output) : null,
    };
}

async function run(argv: readonly string[]): Promise<number> {
    const opts = parseOptions(argv);
    const sourceRoot = path.resolve(opts.sourceRoot);
    const outputRoot = path.resolve(opts.outputRoot);
    const logFile = path.resolve(opts.log);
    const limit = opts.limit ? Number.parseInt(opts.limit, 10) : undefined;
    if (limit !== undefined && (!Number.isFinite(limit) || limit <= 0)) throw new Error("--limit must be a positive integer");

    const sources: string[] = [];
    walkDocx(sourceRoot, sources);
    sources.sort();
    const filtered = opts.match ? sources.filter((f) => path.relative(sourceRoot, f).includes(opts.match ?? "")) : sources;
    const completed = opts.resume ? completedRecords(logFile) : new Set<string>();

    mkdirSync(path.dirname(logFile), { recursive: true });
    if (!opts.resume) writeFileSync(logFile, "");

    let processed = 0;
    let repaired = 0;
    let validatorValid = 0;
    let contentSame = 0;
    for (const source of filtered) {
        const relativePath = path.relative(sourceRoot, source);
        if (completed.has(relativePath)) continue;
        if (limit !== undefined && processed >= limit) break;
        const output = path.join(outputRoot, relativePath);
        const record = await regenerateOne(source, output, relativePath);
        appendFileSync(logFile, `${JSON.stringify(record)}\n`);
        processed += 1;
        if (record.repairs > 0) repaired += 1;
        if (record.validator?.valid) validatorValid += 1;
        if (record.contentSame) contentSame += 1;
        process.stderr.write(
            `[${processed}] ${relativePath}: unpack=${record.unpackOk ? "ok" : "fail"} repairs=${record.repairs} ` +
                `pack=${record.packOk ? "ok" : "fail"} content=${record.contentSame} validator=${record.validator?.valid ?? "n/a"}\n`,
        );
    }

    process.stdout.write(
        JSON.stringify(
            {
                processed,
                repaired,
                validatorValid,
                contentSame,
                log: logFile,
                outputRoot,
            },
            null,
            2,
        ) + "\n",
    );
    return 0;
}

run(process.argv.slice(2)).then(
    (code) => {
        process.exitCode = code;
    },
    (err: unknown) => {
        process.stderr.write(`${err instanceof Error ? err.message : String(err)}\n`);
        process.exitCode = 1;
    },
);
