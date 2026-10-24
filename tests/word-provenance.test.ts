/*
 * Unit tests for the hash-bound Word provenance module (#274, Task 13 part A).
 *
 * Built slice by slice under red/green TDD: every RED run below failed on a
 * behavioural assertion, never on an import or type error.
 */

import { describe, expect, it } from "vitest";

import {
    classifyWordOutcome,
    effectiveOutcome,
    mergeWordObservation,
    parseWordField,
    REJECTED_OUTCOMES,
    toManifestOutcome,
    UNATTRIBUTABLE_OUTCOMES,
    VERIFIER_OUTCOME_TO_MANIFEST,
    wordAccuracy,
    type WordOutcomeClass,
} from "../src/lib/word-provenance";
import { WORD_OUTCOMES, type WordOutcome } from "../src/lib/word-verifier";

describe("effectiveOutcome", () => {
    it("binds a structured observation to its sha256: matching hash keeps the outcome, a differing hash is stale (unknown), never clean or rejected", () => {
        const clean = parseWordField({ outcome: "clean-open", sha256: "aaaa" });
        expect(effectiveOutcome(clean, "aaaa")).toBe("clean-open");
        expect(effectiveOutcome(clean, "bbbb")).toBe("unknown");

        const rejected = parseWordField({ outcome: "open-error", sha256: "aaaa" });
        expect(effectiveOutcome(rejected, "aaaa")).toBe("open-error");
        expect(effectiveOutcome(rejected, "bbbb")).toBe("unknown");
    });
});

describe("accuracy accounting", () => {
    it("keeps unknown, timeout and stale observations out of both the clean and the rejected denominators", () => {
        expect(classifyWordOutcome("clean-open")).toBe("clean");
        expect(classifyWordOutcome("unreadable-content-warning")).toBe("rejected");
        expect(classifyWordOutcome("unknown")).toBe("unattributable");
        expect(classifyWordOutcome("timeout")).toBe("unattributable");

        const counts = wordAccuracy([
            { field: "clean-open", sha256: "aaaa" },
            { field: "timeout", sha256: "aaaa" },
            { field: { outcome: "open-error", sha256: "aaaa" }, sha256: "bbbb" }, // stale
            { field: { outcome: "unknown", sha256: "aaaa" }, sha256: "aaaa" },
            { field: null, sha256: "aaaa" },
        ]);
        expect(counts).toEqual({ clean: 1, rejected: 0, unattributable: 4 });
    });

    it("counts only Word's verdict as a rejection: crash, unavailable and unmapped spellings are unattributable (#316)", () => {
        expect(classifyWordOutcome("crash")).toBe("unattributable");
        expect(classifyWordOutcome("unavailable")).toBe("unattributable");
        // Verifier spellings that were never translated to the manifest
        // vocabulary must not silently count as rejections.
        expect(classifyWordOutcome("repair-warning")).toBe("unattributable");
        expect(classifyWordOutcome("password-prompt")).toBe("unattributable");
        // Nor may an unknown future outcome.
        expect(classifyWordOutcome("a-future-outcome")).toBe("unattributable");
    });
});

describe("verifier → manifest outcome mapping (#316)", () => {
    it("maps every verifier outcome to a manifest outcome the accounting knows — a new WordOutcome without a mapping fails here", () => {
        // A duplicate spelling in WORD_OUTCOMES could shadow a missing mapping.
        expect(WORD_OUTCOMES).toHaveLength(new Set(WORD_OUTCOMES).size);
        // The mapping table is keyed by exactly the verifier outcomes.
        expect(Object.keys(VERIFIER_OUTCOME_TO_MANIFEST).sort()).toEqual([...WORD_OUTCOMES].sort());
        const known = ["clean-open", ...REJECTED_OUTCOMES, ...UNATTRIBUTABLE_OUTCOMES];
        for (const outcome of WORD_OUTCOMES) {
            expect(known, `verifier outcome "${outcome}" maps to no known manifest outcome`).toContain(toManifestOutcome(outcome));
        }
    });

    it("spells the dialog outcomes the manifest's way and never turns a no-verdict outcome into a rejection", () => {
        expect(toManifestOutcome("repair-warning")).toBe("unreadable-content-warning");
        expect(toManifestOutcome("password-prompt")).toBe("password-required");
        expect(toManifestOutcome("unavailable")).toBe("unknown");
        for (const outcome of ["clean-open", "open-error", "crash", "unknown-dialog", "timeout"] as const) {
            expect(toManifestOutcome(outcome)).toBe(outcome);
        }
        const expected: Record<WordOutcome, WordOutcomeClass> = {
            "clean-open": "clean",
            "repair-warning": "rejected",
            "open-error": "rejected",
            "password-prompt": "rejected",
            "unknown-dialog": "rejected",
            crash: "unattributable",
            timeout: "unattributable",
            unavailable: "unattributable",
        };
        for (const outcome of WORD_OUTCOMES) {
            expect(classifyWordOutcome(toManifestOutcome(outcome)), outcome).toBe(expected[outcome]);
        }
    });
});

describe("mergeWordObservation", () => {
    it("never discards an existing observation: unchanged hash keeps it verbatim, changed hash retires it to history under its old hash", () => {
        // Hash unchanged, no fresh probe: the prior observation (and its
        // metadata) is kept and stays bound to the same bytes.
        const kept = mergeWordObservation({
            prior: { outcome: "clean-open", sha256: "aaaa", observedAt: "2026-10-01T00:00:00Z", wordVersion: "16.89" },
            currentSha256: "aaaa",
        });
        expect(kept).toEqual({
            word: { outcome: "clean-open", sha256: "aaaa", observedAt: "2026-10-01T00:00:00Z", wordVersion: "16.89" },
            history: [],
        });

        // Hash changed, no fresh probe: the prior moves to `history` carrying
        // its old hash; the new bytes carry no observation yet.
        const retired = mergeWordObservation({
            prior: { outcome: "open-error", sha256: "aaaa" },
            priorHistory: [{ outcome: "password-required", sha256: "9999" }],
            currentSha256: "bbbb",
        });
        expect(retired).toEqual({
            word: { outcome: "unknown", sha256: "bbbb" },
            history: [
                { outcome: "password-required", sha256: "9999" },
                { outcome: "open-error", sha256: "aaaa" },
            ],
        });

        // Fresh probe outcome wins and is bound to the current bytes; a prior
        // observation of different bytes is retired, not discarded.
        const reProbed = mergeWordObservation({
            prior: { outcome: "clean-open", sha256: "aaaa" },
            freshOutcome: "open-error",
            currentSha256: "bbbb",
            observedAt: "2026-10-02T00:00:00Z",
        });
        expect(reProbed).toEqual({
            word: { outcome: "open-error", sha256: "bbbb", observedAt: "2026-10-02T00:00:00Z" },
            history: [{ outcome: "clean-open", sha256: "aaaa" }],
        });
    });

    it("retires a same-bytes observation that a fresh probe supersedes, and never lets an unattributable re-probe replace attributable evidence", () => {
        const superseded = mergeWordObservation({
            prior: { outcome: "open-error", sha256: "aaaa", observedAt: "2026-10-01T00:00:00Z" },
            freshOutcome: "clean-open",
            currentSha256: "aaaa",
            observedAt: "2026-10-02T00:00:00Z",
        });
        expect(superseded).toEqual({
            word: { outcome: "clean-open", sha256: "aaaa", observedAt: "2026-10-02T00:00:00Z" },
            history: [{ outcome: "open-error", sha256: "aaaa", observedAt: "2026-10-01T00:00:00Z" }],
        });

        // A timeout says nothing about the bytes: the clean-open stays current
        // and the timeout is recorded, not the other way round.
        const timedOut = mergeWordObservation({
            prior: { outcome: "clean-open", sha256: "aaaa" },
            freshOutcome: "timeout",
            currentSha256: "aaaa",
            observedAt: "2026-10-03T00:00:00Z",
        });
        expect(timedOut).toEqual({
            word: { outcome: "clean-open", sha256: "aaaa" },
            history: [{ outcome: "timeout", sha256: "aaaa", observedAt: "2026-10-03T00:00:00Z" }],
        });
    });
});
