/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word-valid regression pin for every corpus fixture with a Word observation
 * (#313, plan Task 13).
 *
 * Word-valid predicts whether Microsoft Word opens a file cleanly, so its
 * expectation is not taken from the validator's own output. It is the
 * fixture's hash-bound Word observation in `tests/fixtures-all.manifest.json`:
 * `clean-open` means valid, and a rejection (open error, unreadable-content
 * warning, password prompt, unknown dialog) means invalid. Observations that
 * carry no evidence (`unknown`, `timeout`) or no longer match the fixture
 * bytes are not asserted, as in `wordAccuracy`. An observation with a verdict
 * but no recorded sha256 is bound to no bytes, so it fails the binding guard
 * instead of counting toward the floors (#330).
 *
 * The strict and lenient counterparts are `fixtures-all-strict.test.ts` and
 * `fixtures-all-lenient.test.ts`.
 */

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { classifyWordOutcome, effectiveOutcome, parseWordField } from "../src/lib/word-provenance";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");
const MANIFEST_PATH = path.join(HERE, "fixtures-all.manifest.json");

// The attributable observations: 351 clean opens and 62 rejections when this
// suite was introduced, plus the 12 fixtures probed after the #315 fix.
// Evidence may grow; it must not silently shrink.
const MIN_CLEAN = 358;
const MIN_REJECTED = 67;

interface ManifestEntry {
    relativePath: string;
    word?: unknown;
}

const manifest = JSON.parse(readFileSync(MANIFEST_PATH, "utf-8")) as { entries: ManifestEntry[] };

const unbound = manifest.entries.flatMap((entry) => {
    const observation = parseWordField(entry.word);
    return observation && classifyWordOutcome(observation.outcome) !== "unattributable" && observation.sha256 === undefined
        ? [entry.relativePath]
        : [];
});

const observed = manifest.entries.flatMap((entry) => {
    if (!/\.docx$/i.test(entry.relativePath)) return [];
    const bytes = readFileSync(path.join(FIXTURES, entry.relativePath));
    const outcome = effectiveOutcome(parseWordField(entry.word), createHash("sha256").update(bytes).digest("hex"));
    const verdict = classifyWordOutcome(outcome);
    return verdict === "clean" || verdict === "rejected" ? [{ relativePath: entry.relativePath, outcome, clean: verdict === "clean" }] : [];
});

describe("all fixtures — word-valid profile against Word's own verdict", () => {
    it("binds every Word verdict in the manifest to the bytes it was observed on", () => {
        expect(unbound, "word observations with a verdict but no sha256").toEqual([]);
    });

    it(`asserts at least ${MIN_CLEAN} clean opens and ${MIN_REJECTED} rejections`, () => {
        expect(observed.filter((entry) => entry.clean).length).toBeGreaterThanOrEqual(MIN_CLEAN);
        expect(observed.filter((entry) => !entry.clean).length).toBeGreaterThanOrEqual(MIN_REJECTED);
    });

    for (const entry of observed) {
        it(`[${entry.outcome}] ${entry.relativePath}`, { timeout: 60_000 }, async () => {
            const result = await validate(path.join(FIXTURES, entry.relativePath), { profile: "word-valid" });
            const blocking = result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code);
            expect(result.valid, `Word: ${entry.outcome}; word-valid errors: ${JSON.stringify(blocking)}`).toBe(entry.clean);
        });
    }
});
