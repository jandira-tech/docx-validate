/*
 * Accuracy-accounting unit tests for the Word probe summary (#274, Task 13
 * part A): `unknown`, `timeout` and stale observations stay out of both the
 * clean and the rejected denominators.
 */

import { describe, expect, it } from "vitest";

import { classify, emptySummary, type WordProbe, updateSummary } from "../scripts/probe-word-fixtures";

const record = (overrides: { validatorValid: boolean; outcome?: WordProbe["outcome"]; clean?: boolean }): Parameters<typeof updateSummary>[1] => ({
    relativePath: "broken/a.docx",
    file: "/fixtures/broken/a.docx",
    validator: { valid: overrides.validatorValid, errorCodes: [], threw: undefined, profile: "word-valid" },
    word:
        overrides.outcome === undefined
            ? null
            : { outcome: overrides.outcome, clean: overrides.clean ?? false, details: "", durationMs: 1 },
    aligned: null,
    mismatch: null,
});

describe("probe summary accuracy denominators", () => {
    it("excludes timeout observations from clean, rejected and alignment counts", () => {
        const summary = emptySummary();
        updateSummary(summary, record({ validatorValid: true, outcome: "timeout", clean: false }));
        updateSummary(summary, record({ validatorValid: true, outcome: "clean-open", clean: true }));
        updateSummary(summary, record({ validatorValid: false, outcome: "open-error", clean: false }));

        expect(summary.probed).toBe(3);
        expect(summary.outcomes).toEqual({ timeout: 1, "clean-open": 1, "open-error": 1 });
        // Only the two attributable observations enter the denominators.
        expect(summary.wordClean).toBe(1);
        expect(summary.wordNotClean).toBe(1);
        expect(summary.aligned).toBe(0);
        expect(summary.falsePositive).toBe(0);
        expect(summary.falseNegative).toBe(0);

        // A timeout is unattributable, not a validator false negative.
        expect(classify(record({ validatorValid: true, outcome: "timeout", clean: false })).aligned).toBe(null);
    });
});
