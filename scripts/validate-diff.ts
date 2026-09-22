// validate-diff — branch-over-branch validator-result regression harness.
//
// docx-validate is a *validator*, not a converter, so its regression analog of
// a "DOCX diff" is the stability of its OWN verdicts: run `validate()` over a
// representative fixture set (scripts/validate-diff-fixtures.json) and compare
// each fixture's result (valid flag + per-issue-code counts) between two builds
// of the library — the PR branch vs its merge base. A drift means a validator
// changed what it flags.
//
// Two subcommands, run under **Node + tsx** (libxmljs2's native addon is built
// for Node's ABI and crashes under Bun's runtime):
//
//   # On each checkout (base + PR): snapshot every fixture's verdict to JSON.
//   node --import tsx scripts/validate-diff.ts build --out <file.json>
//
//   # Once, from the PR checkout: diff base vs PR snapshots into Markdown.
//   node --import tsx scripts/validate-diff.ts report \
//     --baseline <file.json> --current <file.json> --output diff-report.md
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { runCli } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(HERE, "..");
const MANIFEST = path.join(HERE, "validate-diff-fixtures.json");

interface FixtureEntry {
    name: string;
    path: string;
}
interface Snapshot {
    valid: boolean;
    codes: Record<string, number>;
}

function loadFixtures(): FixtureEntry[] {
    const raw = JSON.parse(readFileSync(MANIFEST, "utf8")) as {
        fixtures: FixtureEntry[];
    };
    return raw.fixtures;
}

export function parseFlag(argv: string[], name: string): string | undefined {
    const i = argv.indexOf(name);
    if (i === -1 || i + 1 >= argv.length) return undefined;
    const val = argv[i + 1];
    if (!val || val.startsWith("-")) return undefined;
    return val;
}

export function errorMessage(err: unknown): string {
    return err instanceof Error ? err.message : String(err);
}

function flag(name: string): string | undefined {
    // Skip the node binary, script path, and subcommand (`build`/`report`).
    return parseFlag(process.argv.slice(3), name);
}

async function snapshot(absPath: string): Promise<Snapshot> {
    const result = await validate(absPath);
    const codes: Record<string, number> = {};
    for (const issue of result.issues) {
        const key = issue.code ?? `severity:${issue.severity}`;
        codes[key] = (codes[key] ?? 0) + 1;
    }
    return { valid: result.valid, codes };
}

async function cmdBuild(): Promise<number> {
    const out = flag("--out");
    if (!out) {
        console.error("Usage: validate-diff.ts build --out <file.json>");
        return 2;
    }
    const fixtures = loadFixtures();
    const snaps: Record<string, Snapshot | { error: string }> = {};
    for (const fx of fixtures) {
        try {
            snaps[fx.name] = await snapshot(path.resolve(REPO_ROOT, fx.path));
        } catch (e) {
            snaps[fx.name] = { error: errorMessage(e) };
        }
    }
    writeFileSync(path.resolve(out), `${JSON.stringify(snaps, null, 2)}\n`);
    console.log(`validate-diff build: ${fixtures.length} fixtures → ${out}`);
    return 0;
}

function diffCodes(base: Record<string, number>, cur: Record<string, number>): string[] {
    const keys = [...new Set([...Object.keys(base), ...Object.keys(cur)])].sort();
    const rows: string[] = [];
    for (const k of keys) {
        const b = base[k] ?? 0;
        const c = cur[k] ?? 0;
        if (b !== c) rows.push(`| \`${k}\` | ${b} | ${c} | ${c - b > 0 ? `+${c - b}` : c - b} |`);
    }
    return rows;
}

async function cmdReport(): Promise<number> {
    const baseFile = flag("--baseline");
    const curFile = flag("--current");
    const output = flag("--output") ?? "diff-report.md";
    if (!baseFile || !curFile) {
        console.error("Usage: validate-diff.ts report --baseline <file> --current <file> [--output <file>]");
        return 2;
    }
    const base = JSON.parse(readFileSync(path.resolve(baseFile), "utf8")) as Record<string, Snapshot | { error: string }>;
    const cur = JSON.parse(readFileSync(path.resolve(curFile), "utf8")) as Record<string, Snapshot | { error: string }>;
    const fixtures = loadFixtures();

    let anyDrift = false;
    const sections: string[] = [];
    for (const fx of fixtures) {
        const b = base[fx.name];
        const c = cur[fx.name];
        if (!b || !c || "error" in b || "error" in c) {
            anyDrift = true;
            const msg = (b && "error" in b ? b.error : undefined) ?? (c && "error" in c ? c.error : "missing snapshot");
            sections.push(`### \`${fx.name}\` ⚠️\n\nCould not compare: ${msg}\n`);
            continue;
        }
        const codeRows = diffCodes(b.codes, c.codes);
        const validChanged = b.valid !== c.valid;
        if (!validChanged && codeRows.length === 0) {
            sections.push(`### \`${fx.name}\` ✅\n\nVerdict unchanged (valid=${c.valid}).\n`);
            continue;
        }
        anyDrift = true;
        let body = `### \`${fx.name}\` ⚠️\n\n`;
        if (validChanged) body += `\`valid\`: **${b.valid} → ${c.valid}**\n\n`;
        if (codeRows.length > 0) {
            body += `| Issue code | Base | PR | Δ |\n|---|---:|---:|---:|\n${codeRows.join("\n")}\n`;
        }
        sections.push(body);
    }

    const header =
        "# 🔎 docx-validate Validator Diff Report\n\n" +
        "Branch-over-branch stability of the validator's own verdicts: each " +
        "fixture's `valid` flag and per-issue-code counts, compared between the " +
        "merge base and this PR. A drift means a validator changed what it flags.\n\n";
    const verdict = anyDrift
        ? "**Result: ⚠️ validator drift detected** — review the per-fixture changes below.\n"
        : "**Result: ✅ no validator drift** — every fixture's verdict is identical between base and PR.\n";
    const footer = "\n---\n*Generated by `scripts/validate-diff.ts` • validator-verdict stability • run under Node + tsx.*\n";
    writeFileSync(path.resolve(output), `${header}${verdict}\n${sections.join("\n")}${footer}`);
    console.log(`validate-diff report → ${output} (drift=${anyDrift})`);
    return anyDrift ? 1 : 0;
}

export async function main(argv: string[] = process.argv): Promise<number> {
    try {
        const cmd = argv[2];
        if (cmd === "build") return await cmdBuild();
        if (cmd === "report") return await cmdReport();
        console.error("Usage: validate-diff.ts <build|report> [flags]");
        return 2;
    } catch (err) {
        console.error(`Error: ${errorMessage(err)}`);
        return 1;
    }
}

runCli(import.meta.url, () => main());
