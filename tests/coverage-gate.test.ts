/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { checkCoverageGates, COVERAGE_GATES, type CoverageSummary } from "../scripts/coverage-gate";

const ROOT = "/repo";
const at = (lines: number, branches: number) => ({ lines: { pct: lines }, branches: { pct: branches } });

describe("checkCoverageGates", () => {
    const gates = [
        { pattern: "src/lib/check-scheduler.ts", lines: 90, branches: 85 },
        { pattern: "src/scripts/office/checks/*.ts", lines: 80, branches: 70 },
    ];

    it("passes when every gated file meets its thresholds", () => {
        const summary: CoverageSummary = {
            total: at(50, 40),
            "/repo/src/lib/check-scheduler.ts": at(90, 85),
            "/repo/src/scripts/office/checks/a.ts": at(80, 70),
            "/repo/src/scripts/office/checks/b.ts": at(100, 100),
            "/repo/src/lib/run-cli.ts": at(10, 10),
        };
        expect(checkCoverageGates(summary, gates, ROOT)).toEqual([]);
    });

    it("names each file below a threshold, with both measures", () => {
        const summary: CoverageSummary = {
            "/repo/src/lib/check-scheduler.ts": at(97.8, 84.9),
            "/repo/src/scripts/office/checks/a.ts": at(79, 90),
        };
        expect(checkCoverageGates(summary, gates, ROOT)).toEqual([
            "src/lib/check-scheduler.ts: 97.8% lines, 84.9% branches (gate 90% lines, 85% branches)",
            "src/scripts/office/checks/a.ts: 79% lines, 90% branches (gate 80% lines, 70% branches)",
        ]);
    });

    it("fails a gate that matches no file, so a partial run cannot pass", () => {
        expect(checkCoverageGates({ "/repo/src/scripts/office/checks/a.ts": at(100, 100) }, gates, ROOT)).toEqual([
            "src/lib/check-scheduler.ts: no coverage recorded (run the full suite with --coverage)",
        ]);
    });

    it("gates the modules the plan names", () => {
        const patterns = COVERAGE_GATES.map((gate) => `${gate.pattern} ${gate.lines}/${gate.branches}`);
        expect(patterns).toEqual(
            expect.arrayContaining([
                "src/lib/mode-policy.ts 90/85",
                "src/lib/check-scheduler.ts 90/85",
                "src/lib/check-registry.ts 90/85",
                "src/lib/validation-context.ts 80/70",
                "src/lib/properties-parts.ts 80/70",
                "src/scripts/office/package-context.ts 80/70",
                "src/scripts/office/checks/*.ts 80/70",
            ]),
        );
    });
});
