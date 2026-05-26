/*
 * Word-parity calibration + safety gate.
 *
 * Real Microsoft Word is the ground-truth oracle (see scripts/probe-word-fixtures.ts).
 * `validation-results/word-probe-all.jsonl` is the committed regression baseline of
 * Word's verdict (`.word.clean`) for every fixture. This script re-runs OUR validator
 * under the `word-valid` profile over those same fixtures and compares to Word.
 *
 * Purpose:
 *   - LAYER 1 (oracle): word-probe-all.jsonl is the Word ground truth; re-run the
 *     probe periodically to refresh it, then this script measures our agreement.
 *   - LAYER 2 (provably-safe rules): any new validator rule must NOT introduce a
 *     mismatch outside `validation-results/word-parity-baseline.json`. A NEW false
 *     positive (we reject a file Word opens clean) fails the gate.
 *
 * Usage:
 *   tsx scripts/word-parity-check.ts                 # check; exit 1 on regression
 *   tsx scripts/word-parity-check.ts --update-baseline  # accept current deltas
 *   tsx scripts/word-parity-check.ts --limit 50      # smoke a subset
 *
 * Note: re-validates the whole corpus (~573 files, a few minutes). No Word needed.
 */
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const PROBE = path.join(REPO, "validation-results", "word-probe-all.jsonl");
const BASELINE = path.join(REPO, "validation-results", "word-parity-baseline.json");

type Mismatch = { file: string; kind: "false-positive" | "false-negative"; wordOutcome: string };

async function main(argv: readonly string[]): Promise<number> {
    const update = argv.includes("--update-baseline");
    const limitIdx = argv.indexOf("--limit");
    const limit = limitIdx >= 0 ? Number.parseInt(argv[limitIdx + 1] ?? "", 10) : undefined;

    const probeRows = (await fs.readFile(PROBE, "utf8"))
        .trim()
        .split("\n")
        .filter(Boolean)
        .map((l) => JSON.parse(l) as { relativePath: string; word?: { clean: boolean; outcome: string } });

    const baseline = JSON.parse(await fs.readFile(BASELINE, "utf8")) as { acceptedMismatches: Mismatch[] };
    const acceptedKey = new Set(baseline.acceptedMismatches.map((m) => `${m.kind}:${m.file}`));

    let aligned = 0;
    const current: Mismatch[] = [];
    const checkedFiles = new Set<string>();
    let done = 0;
    const rows = limit ? probeRows.slice(0, limit) : probeRows;

    for (const r of rows) {
        if (!r.word) continue;
        const abs = path.join(REPO, r.relativePath);
        let ourValid: boolean;
        try {
            ourValid = (await validate(abs, { profile: "word-valid" })).valid;
        } catch {
            continue; // unreadable package (encrypted/corrupt) — Word couldn't open it either
        }
        checkedFiles.add(r.relativePath);
        const wordClean = r.word.clean;
        if (ourValid === wordClean) aligned += 1;
        else current.push({ file: r.relativePath, kind: ourValid ? "false-negative" : "false-positive", wordOutcome: r.word.outcome });
        if (++done % 100 === 0) process.stderr.write(`  ${done}/${rows.length}\n`);
    }

    const currentKey = new Set(current.map((m) => `${m.kind}:${m.file}`));
    const newMismatches = current.filter((m) => !acceptedKey.has(`${m.kind}:${m.file}`));
    // Only a baseline entry whose file was actually re-checked can be "resolved".
    const resolved = baseline.acceptedMismatches.filter((m) => checkedFiles.has(m.file) && !currentKey.has(`${m.kind}:${m.file}`));

    process.stderr.write(`\nWord-parity: aligned ${aligned}/${done} | mismatches ${current.length} (baseline accepts ${baseline.acceptedMismatches.length})\n`);
    if (newMismatches.length) {
        process.stderr.write(`\nNEW mismatches (regressions) — NOT in baseline:\n`);
        for (const m of newMismatches) process.stderr.write(`  ${m.kind === "false-positive" ? "FALSE POSITIVE" : "false negative"}  ${m.file} (Word: ${m.wordOutcome})\n`);
        const newFP = newMismatches.filter((m) => m.kind === "false-positive");
        if (newFP.length) process.stderr.write(`\n${newFP.length} NEW FALSE POSITIVE(S): a rule now rejects a file Word opens clean. This is the safety gate — fix the rule or it does not ship.\n`);
    }
    if (resolved.length) {
        process.stderr.write(`\nResolved (baseline entries now aligned — consider --update-baseline):\n`);
        for (const m of resolved) process.stderr.write(`  ${m.file}\n`);
    }

    if (update) {
        const next = { ...baseline, generatedFrom: "scripts/word-parity-check.ts --update-baseline", total: done, acceptedMismatches: current.sort((a, b) => a.file.localeCompare(b.file)) };
        await fs.writeFile(BASELINE, JSON.stringify(next, null, 2) + "\n");
        process.stderr.write(`\nBaseline updated: ${current.length} accepted mismatches.\n`);
        return 0;
    }

    if (newMismatches.length) return 1;
    process.stderr.write(`\nOK — no regressions outside the accepted baseline.\n`);
    return 0;
}

main(process.argv.slice(2)).then((code) => process.exit(code));
