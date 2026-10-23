/*
 * Unit tests for the hash-bound Word provenance module (#274, Task 13 part A).
 *
 * Built slice by slice under red/green TDD: every RED run below failed on a
 * behavioural assertion, never on an import or type error.
 */

import { describe, expect, it } from "vitest";

import { effectiveOutcome, parseWordField } from "../src/lib/word-provenance";

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
