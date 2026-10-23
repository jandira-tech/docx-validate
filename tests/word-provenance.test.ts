/*
 * Unit tests for the hash-bound Word provenance module (#274, Task 13 part A).
 *
 * Built slice by slice under red/green TDD: every RED run below failed on a
 * behavioural assertion, never on an import or type error.
 */

import { describe, expect, it } from "vitest";

import { classifyWordOutcome, effectiveOutcome, parseWordField, wordAccuracy } from "../src/lib/word-provenance";

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
