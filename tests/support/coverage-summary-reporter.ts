/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Prints one line per coverage run: `Coverage: 92.1% lines, 87.3% branches`.
 * Vitest hands reporters the istanbul coverage map after the coverage
 * reporters have written; with `reportOnFailure` that happens on failed runs
 * too, so red runs print their coverage as well as green ones.
 */

import type { Reporter } from "vitest/node";

interface Pct {
    pct: number;
}
interface CoverageMapLike {
    getCoverageSummary?: () => { lines: Pct; branches: Pct };
}

export default class CoverageSummaryReporter implements Reporter {
    onCoverage(coverage: unknown): void {
        const summary = (coverage as CoverageMapLike | null)?.getCoverageSummary?.();
        if (!summary) return;
        process.stdout.write(`Coverage: ${summary.lines.pct}% lines, ${summary.branches.pct}% branches\n`);
    }
}
