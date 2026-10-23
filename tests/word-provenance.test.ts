/*
 * Unit tests for the hash-bound Word provenance module (#274, Task 13 part A).
 *
 * Built slice by slice under red/green TDD: every RED run below failed on a
 * behavioural assertion, never on an import or type error.
 */

import { describe, expect, it } from "vitest";

import { classifyWordOutcome, effectiveOutcome, mergeWordObservation, parseWordField, wordAccuracy } from "../src/lib/word-provenance";

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
