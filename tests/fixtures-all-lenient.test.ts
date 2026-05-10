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
 * Lenient-profile regression pin for every `.docx` / `.pptx` under
 * `tests/fixtures/`.
 *
 * Same shape as `fixtures-all-strict.test.ts`, but uses
 * `profile: "lenient"`. Strict-only checks (BOMs, threaded-comment
 * paraId-missing) downgrade to warnings here, so fixtures that fail
 * strict but pass lenient are NOT flagged as test failures — instead an
 * `afterAll` block prints a big banner listing every fixture that would
 * have failed under strict. That keeps the integrity gap visible without
 * coupling the lenient-profile regression suite to strict-profile
 * findings.
 */

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterAll, describe, expect, it } from "vitest";

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

interface StrictGap {
    relativePath: string;
    warningCodes: string[];
}

const strictGaps: StrictGap[] = [];

describe("all fixtures — lenient profile", () => {
    it(`manifest covers all ${manifest.totalFixtures} fixtures`, () => {
        expect(manifest.entries.length).toBe(manifest.totalFixtures);
    });

    for (const entry of manifest.entries) {
        const expected = entry.lenient;
        const label = expected.valid ? "PASS" : expected.threw ? "THROW" : "FAIL";
        const fixturePath = path.join(FIXTURES, entry.relativePath);
        const willTrigger = expected.valid && !entry.strict.valid;

        const suffix = willTrigger ? " (passes lenient; would fail strict)" : "";

        // 60s per fixture: matches fixtures-all-strict.test.ts. A handful of
        // heavy-XSD fixtures take longer than vitest's 5s default.
        it(`[${label}] ${entry.relativePath}${suffix}`, { timeout: 60_000 }, async () => {
            if (expected.threw) {
                let captured: unknown;
                try {
                    await validate(fixturePath, { profile: "lenient" });
                } catch (e) {
                    captured = e;
                }
                expect(captured, `expected validate() to throw for ${entry.relativePath}`).toBeInstanceOf(Error);
                return;
            }

            const result = await validate(fixturePath, { profile: "lenient" });
            const errorCodes = Array.from(
                new Set(
                    result.issues
                        .filter((i) => i.severity === "error")
                        .map((i) => i.code)
                        .filter((c): c is string => Boolean(c)),
                ),
            ).sort();

            expect(result.valid, `expected ${expected.valid ? "valid" : "invalid"} in lenient mode`).toBe(expected.valid);
            expect(errorCodes).toEqual(expected.errorCodes);

            // If this fixture would fail strict but passes lenient, capture
            // the warnings so the afterAll banner can render them — every
            // strict-only check should still surface as a warning here.
            if (willTrigger && result.valid) {
                const warningCodes = Array.from(
                    new Set(
                        result.issues
                            .filter((i) => i.severity === "warning")
                            .map((i) => i.code)
                            .filter((c): c is string => Boolean(c)),
                    ),
                ).sort();
                strictGaps.push({ relativePath: entry.relativePath, warningCodes });
            }
        });
    }

    afterAll(() => {
        if (strictGaps.length === 0) return;
        const lines: string[] = [];
        const banner = "═".repeat(78);
        lines.push("");
        lines.push(banner);
        lines.push(`⚠  ${strictGaps.length} fixture(s) pass LENIENT but would fail STRICT validation.`);
        lines.push("   Run `bun run test:strict` to see the failure detail.");
        lines.push(banner);
        for (const gap of strictGaps) {
            const codes = gap.warningCodes.length === 0 ? "(no warnings emitted — investigate)" : gap.warningCodes.join(", ");
            lines.push(`  - ${gap.relativePath}`);
            lines.push(`      warnings: ${codes}`);
        }
        lines.push(banner);
        lines.push("");
        process.stderr.write(`${lines.join("\n")}\n`);
    });
});
