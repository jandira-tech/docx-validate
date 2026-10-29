/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Enforces the scoped coverage gates of the Word-first validation plan
 * (docs/plans/2026-10-23-fixing-validator.md, section 1): every module the
 * plan adds needs at least 80% lines and 70% branches; validity policy, check
 * scheduling and result aggregation need 90% and 85%. A full-suite
 * percentage cannot show either, so this reads the per-file json-summary that
 * `vitest run --coverage` writes. Run it after a full coverage run:
 * `pnpm run test:coverage-gate`.
 */

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export type CoverageGate = {
    /** A repository-relative path, or a glob over them (`path.matchesGlob`). */
    readonly pattern: string;
    readonly lines: number;
    readonly branches: number;
};

type Percent = { readonly pct: number };
export type CoverageSummary = Readonly<Record<string, { readonly lines: Percent; readonly branches: Percent }>>;

const NEW_MODULE = { lines: 80, branches: 70 } as const;
const POLICY = { lines: 90, branches: 85 } as const;

export const COVERAGE_GATES: readonly CoverageGate[] = [
    // Validity policy: which profile, defaults and selectors a run executes.
    { pattern: "src/lib/mode-policy.ts", ...POLICY },
    // Check scheduling, and the registry it schedules from.
    { pattern: "src/lib/check-scheduler.ts", ...POLICY },
    { pattern: "src/lib/check-registry.ts", ...POLICY },
    // Result aggregation (`mergeResults`, `finalize`).
    { pattern: "src/lib/types.ts", ...POLICY },
    // Which errors the word-valid profile keeps as errors.
    { pattern: "src/scripts/office/validators/word-blocking.ts", ...POLICY },
    // Modules the plan adds.
    { pattern: "src/lib/content-types.ts", ...NEW_MODULE },
    { pattern: "src/lib/properties-parts.ts", ...NEW_MODULE },
    { pattern: "src/lib/markup-compatibility.ts", ...NEW_MODULE },
    { pattern: "src/lib/validation-context.ts", ...NEW_MODULE },
    { pattern: "src/lib/word-provenance.ts", ...NEW_MODULE },
    { pattern: "src/lib/word-verifier.ts", ...NEW_MODULE },
    { pattern: "src/lib/xml-part-io.ts", ...NEW_MODULE },
    { pattern: "src/lib/xml-syntax.ts", ...NEW_MODULE },
    { pattern: "src/lib/zip-central-directory.ts", ...NEW_MODULE },
    { pattern: "src/scripts/office/package-context.ts", ...NEW_MODULE },
    { pattern: "src/scripts/office/word-namespaces.ts", ...NEW_MODULE },
    { pattern: "src/scripts/office/checks/*.ts", ...NEW_MODULE },
];

/** One message per gate a file misses, or per gate no file in the summary matches. */
export function checkCoverageGates(summary: CoverageSummary, gates: readonly CoverageGate[], root: string): string[] {
    const files = Object.entries(summary)
        .filter(([file]) => file !== "total")
        .map(([file, measures]) => [path.relative(root, file).split(path.sep).join("/"), measures] as const)
        .sort(([a], [b]) => a.localeCompare(b));
    const failures: string[] = [];
    for (const gate of gates) {
        const matched = files.filter(([file]) => path.posix.matchesGlob(file, gate.pattern));
        if (matched.length === 0) {
            failures.push(`${gate.pattern}: no coverage recorded (run the full suite with --coverage)`);
            continue;
        }
        for (const [file, { lines, branches }] of matched) {
            if (lines.pct >= gate.lines && branches.pct >= gate.branches) continue;
            failures.push(
                `${file}: ${lines.pct}% lines, ${branches.pct}% branches (gate ${gate.lines}% lines, ${gate.branches}% branches)`,
            );
        }
    }
    return failures;
}

const main = (): void => {
    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
    const summaryFile = path.join(root, "coverage", "coverage-summary.json");
    if (!existsSync(summaryFile)) {
        process.stderr.write(`coverage-gate: ${path.relative(root, summaryFile)} is missing; run \`vitest run --coverage\` first\n`);
        process.exit(1);
    }
    const summary = JSON.parse(readFileSync(summaryFile, "utf-8")) as CoverageSummary;
    const failures = checkCoverageGates(summary, COVERAGE_GATES, root);
    for (const failure of failures) process.stderr.write(`coverage-gate: ${failure}\n`);
    if (failures.length > 0) process.exit(1);
    process.stdout.write(`coverage-gate: ${COVERAGE_GATES.length} gates met\n`);
};

if (import.meta.url === `file://${process.argv[1]}`) main();
