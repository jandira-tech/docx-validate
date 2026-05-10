/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Strict-mode regression pin for every `.docx` / `.pptx` under
 * `tests/fixtures/` (broken/, external/{apache-poi,docx-templates,docx4j,
 * mammoth-js,open-xml-sdk,superdoc}/, plus the lone vfdsdfcACawesd.docx).
 *
 * Each fixture is asserted with a specific expected outcome — pass or fail
 * — so a *change* in validator behaviour shows up as a CI red, not a
 * silent drift in coverage. The manifest of expected outcomes lives in
 * `tests/fixtures-all.manifest.json`, regenerated via
 * `bunx tsx scripts/probe-all-fixtures.ts`.
 *
 * The lenient counterpart of this suite is `fixtures-all-lenient.test.ts`.
 */

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");
const MANIFEST_PATH = path.join(HERE, "fixtures-all.manifest.json");

interface ManifestOutcome {
    valid: boolean;
    errorCodes: string[];
    threw?: string;
}

interface ManifestEntry {
    relativePath: string;
    strict: ManifestOutcome;
    lenient: ManifestOutcome;
}

interface Manifest {
    generatedAt: string;
    totalFixtures: number;
    entries: ManifestEntry[];
}

const manifest = JSON.parse(readFileSync(MANIFEST_PATH, "utf-8")) as Manifest;

describe("all fixtures — strict profile", () => {
    it(`manifest covers all ${manifest.totalFixtures} fixtures`, () => {
        expect(manifest.entries.length).toBe(manifest.totalFixtures);
    });

    for (const entry of manifest.entries) {
        const expected = entry.strict;
        const label = expected.valid ? "PASS" : expected.threw ? "THROW" : "FAIL";
        const fixturePath = path.join(FIXTURES, entry.relativePath);

        // 60s per fixture: a few large/complex fixtures (heavy XSD-validation
        // passes through libxmljs2) take longer than vitest's 5s default.
        it(`[${label}] ${entry.relativePath}`, { timeout: 60_000 }, async () => {
            if (expected.threw) {
                let captured: unknown;
                try {
                    await validate(fixturePath, { profile: "strict" });
                } catch (e) {
                    captured = e;
                }
                expect(captured, `expected validate() to throw for ${entry.relativePath}`).toBeInstanceOf(Error);
                return;
            }

            const result = await validate(fixturePath, { profile: "strict" });
            const errorCodes = Array.from(
                new Set(
                    result.issues
                        .filter((i) => i.severity === "error")
                        .map((i) => i.code)
                        .filter((c): c is string => Boolean(c)),
                ),
            ).sort();

            // Pin pass/fail.
            expect(result.valid, `expected ${expected.valid ? "valid" : "invalid"} in strict mode`).toBe(expected.valid);

            // Pin the *set* of error codes — drift indicates an unintended
            // behaviour change. We compare sorted arrays for stable diffs.
            expect(errorCodes).toEqual(expected.errorCodes);
        });
    }
});
