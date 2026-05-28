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

/*
 * Fixture repair-drift survey.
 *
 *   1. Copy every .docx fixture into <out>/copies/<relpath> (untouched original).
 *   2. Repair the copy (DOCXSchemaValidator.repair, strict profile) and write the
 *      repaired .docx into <out>/repaired/<relpath>.
 *   3. Determine drift by comparing the pre-repair vs post-repair semantic
 *      inventory (diffDocxInventories), classifying severity (inventoryDiffToIssues).
 *
 * Writes <out>/DRIFT_REPORT.md and prints a summary.
 *
 * Usage:
 *   bunx tsx scripts/fixtures-drift-report.ts [--out <dir>] [--limit <n>] [--profile strict|lenient]
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import {
    collectDocxSemanticInventory,
    type DocxSemanticInventory,
} from "../src/scripts/office/validators/docx-diagnostics";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { diffDocxInventories, inventoryDiffToIssues } from "../src/scripts/office/validators/docx-inventory-diff";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_ROOT = path.resolve(HERE, "..", "tests", "fixtures");

interface FixtureResult {
    relativePath: string;
    repairs: number;
    added: number;
    removed: number;
    changed: number;
    errorCount: number; // content-loss (error-tier) deltas
    warnCount: number;
    lossLabels: string[]; // up to a few content-loss descriptions
    error?: string; // populated if the fixture could not be processed
}

function parseArgs(argv: readonly string[]): { out: string; limit: number; profile: Profile } {
    let out = path.resolve(HERE, "..", ".drift-run");
    let limit = Number.POSITIVE_INFINITY;
    let profile: Profile = "strict";
    for (let i = 0; i < argv.length; i += 1) {
        if (argv[i] === "--out" && argv[i + 1]) out = path.resolve(argv[++i]!);
        else if (argv[i] === "--limit" && argv[i + 1]) limit = Number.parseInt(argv[++i]!, 10);
        else if (argv[i] === "--profile" && argv[i + 1]) {
            const p = argv[++i]!;
            if (p === "strict" || p === "lenient" || p === "word-valid") profile = p;
        }
    }
    return { out, limit, profile };
}

async function walkDocx(dir: string, out: string[]): Promise<void> {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) await walkDocx(full, out);
        else if (entry.isFile() && full.toLowerCase().endsWith(".docx")) out.push(full);
    }
}

async function extractZip(buf: Buffer, outputPath: string): Promise<void> {
    const zip = await JSZip.loadAsync(buf);
    const entries: Array<{ name: string; file: JSZip.JSZipObject }> = [];
    zip.forEach((relativePath, file) => entries.push({ name: relativePath, file }));
    for (const { name, file } of entries) {
        const resolved = path.resolve(outputPath, name);
        if (!resolved.startsWith(`${outputPath}${path.sep}`) && resolved !== outputPath) {
            throw new Error(`zip-slip: ${name}`);
        }
        if (file.dir) {
            await fs.mkdir(resolved, { recursive: true });
            continue;
        }
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        await fs.writeFile(resolved, await file.async("nodebuffer"));
    }
}

async function repackZip(dir: string, outFile: string): Promise<void> {
    const zip = new JSZip();
    async function add(current: string): Promise<void> {
        for (const entry of await fs.readdir(current, { withFileTypes: true })) {
            const full = path.join(current, entry.name);
            const rel = path.relative(dir, full).split(path.sep).join("/");
            if (entry.isDirectory()) await add(full);
            else zip.file(rel, await fs.readFile(full));
        }
    }
    await add(dir);
    await fs.mkdir(path.dirname(outFile), { recursive: true });
    await fs.writeFile(outFile, await zip.generateAsync({ type: "nodebuffer" }));
}

async function processFixture(file: string, out: string, profile: Profile): Promise<FixtureResult> {
    const relativePath = path.relative(FIXTURES_ROOT, file).split(path.sep).join("/");
    const base: FixtureResult = { relativePath, repairs: 0, added: 0, removed: 0, changed: 0, errorCount: 0, warnCount: 0, lossLabels: [] };
    try {
        const buf = await fs.readFile(file);

        // 1. Copy the original untouched into <out>/copies/<relpath>.
        const copyPath = path.join(out, "copies", relativePath);
        await fs.mkdir(path.dirname(copyPath), { recursive: true });
        await fs.writeFile(copyPath, buf);

        return await withTempDir(async (tmp) => {
            const beforeDir = path.join(tmp, "before");
            const afterDir = path.join(tmp, "after");
            await extractZip(buf, beforeDir);
            await extractZip(buf, afterDir);

            const beforeInv: DocxSemanticInventory = await collectDocxSemanticInventory(beforeDir, profile);

            // 2. Repair the copy's unpacked tree.
            const repairs = await new DOCXSchemaValidator({ unpackedDir: afterDir, profile }).repair();
            await repackZip(afterDir, path.join(out, "repaired", relativePath));

            const afterInv = await collectDocxSemanticInventory(afterDir, profile);

            // 3. Determine drift.
            const diff = diffDocxInventories(beforeInv, afterInv);
            const issues = inventoryDiffToIssues(diff);
            const errors = issues.filter((i) => i.severity === "error");
            return {
                relativePath,
                repairs,
                added: diff.added.length,
                removed: diff.removed.length,
                changed: diff.changed.length,
                errorCount: errors.length,
                warnCount: issues.filter((i) => i.severity === "warning").length,
                lossLabels: errors.slice(0, 5).map((i) => i.message),
            };
        });
    } catch (err) {
        return { ...base, error: err instanceof Error ? err.message : String(err) };
    }
}

function report(results: FixtureResult[], profile: Profile): string {
    const errored = results.filter((r) => r.error);
    const ok = results.filter((r) => !r.error);
    const withDrift = ok.filter((r) => r.added + r.removed + r.changed > 0);
    const withLoss = ok.filter((r) => r.errorCount > 0);
    const repaired = ok.filter((r) => r.repairs > 0);

    const lines: string[] = [
        `# Fixture repair-drift report (profile: ${profile})`,
        "",
        `- Fixtures processed: **${ok.length}** (${errored.length} could not be processed)`,
        `- Repaired (repairs > 0): **${repaired.length}**`,
        `- Showed drift (any add/remove/change): **${withDrift.length}**`,
        `- **Content loss (error-tier): ${withLoss.length}**`,
        "",
    ];

    if (withLoss.length > 0) {
        lines.push("## Content loss (error-tier — a content-bearing element decreased)", "");
        lines.push("| Fixture | repairs | -content | drift (a/r/c) |", "|---|---|---|---|");
        for (const r of [...withLoss].sort((a, b) => b.errorCount - a.errorCount)) {
            lines.push(`| \`${r.relativePath}\` | ${r.repairs} | ${r.errorCount} | ${r.added}/${r.removed}/${r.changed} |`);
        }
        lines.push("");
        lines.push("### Loss detail (first few per fixture)", "");
        for (const r of withLoss) {
            lines.push(`**\`${r.relativePath}\`**`);
            for (const l of r.lossLabels) lines.push(`- ${l}`);
            lines.push("");
        }
    }

    const driftOnly = withDrift.filter((r) => r.errorCount === 0);
    if (driftOnly.length > 0) {
        lines.push("## Non-loss drift (warn/info only — fidelity/shape/metadata changed, no content lost)", "");
        lines.push("| Fixture | repairs | drift (a/r/c) | warns |", "|---|---|---|---|");
        for (const r of [...driftOnly].sort((a, b) => b.warnCount - a.warnCount)) {
            lines.push(`| \`${r.relativePath}\` | ${r.repairs} | ${r.added}/${r.removed}/${r.changed} | ${r.warnCount} |`);
        }
        lines.push("");
    }

    if (errored.length > 0) {
        lines.push("## Could not process", "");
        for (const r of errored) lines.push(`- \`${r.relativePath}\`: ${r.error}`);
        lines.push("");
    }
    return lines.join("\n");
}

async function main(): Promise<void> {
    const { out, limit, profile } = parseArgs(process.argv.slice(2));
    const files: string[] = [];
    await walkDocx(FIXTURES_ROOT, files);
    files.sort((a, b) => a.localeCompare(b));
    const selected = Number.isFinite(limit) ? files.slice(0, limit) : files;

    process.stdout.write(`Processing ${selected.length} .docx fixture(s) → ${out} (profile: ${profile})\n`);
    const results: FixtureResult[] = [];
    for (let i = 0; i < selected.length; i += 1) {
        const r = await processFixture(selected[i]!, out, profile);
        results.push(r);
        if ((i + 1) % 25 === 0 || i + 1 === selected.length) {
            process.stdout.write(`  ${i + 1}/${selected.length}\n`);
        }
    }

    const md = report(results, profile);
    await fs.mkdir(out, { recursive: true });
    await fs.writeFile(path.join(out, "DRIFT_REPORT.md"), md, "utf-8");

    const ok = results.filter((r) => !r.error);
    const withLoss = ok.filter((r) => r.errorCount > 0).length;
    const withDrift = ok.filter((r) => r.added + r.removed + r.changed > 0).length;
    process.stdout.write(
        `\nDone. processed=${ok.length} errored=${results.length - ok.length} drift=${withDrift} content-loss=${withLoss}\n` +
            `Report: ${path.join(out, "DRIFT_REPORT.md")}\nCopies: ${path.join(out, "copies")}\nRepaired: ${path.join(out, "repaired")}\n`,
    );
}

main().catch((err: unknown) => {
    process.stderr.write(`${err instanceof Error ? (err.stack ?? err.message) : String(err)}\n`);
    process.exit(1);
});
