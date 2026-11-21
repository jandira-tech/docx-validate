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
 * CI gate for `word-element-invalid-fixtures/`: every broken fixture must
 * FAIL `validate()` under the lenient profile with an XSD error that names
 * the element it represents, and its bytes must match the sha256 pinned in
 * `manifest.jsonl`. Each fixture carries the absolute minimal mutation that
 * breaks the element's own schema rule (see scripts/break-element-fixtures.ts).
 */

import { createHash } from "node:crypto";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const CORPUS = path.join(REPO, "word-element-invalid-fixtures");
const MANIFEST_PATH = path.join(CORPUS, "manifest.jsonl");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

interface ManifestRow {
    element: string;
    file: string;
    sha256: string;
    profile: string;
    valid: boolean;
    mutation: string;
    issueCount: number;
    errorCodes: string[];
    errorMessages: string[];
}

const rows = readFileSync(MANIFEST_PATH, "utf8")
    .trim()
    .split("\n")
    .map((line) => JSON.parse(line) as ManifestRow);

describe("word-element-invalid-fixtures — every fixture fails its own element's rule", () => {
    it("manifest covers every .docx in the corpus exactly once", () => {
        const files = readdirSync(CORPUS)
            .filter((f) => f.endsWith(".docx"))
            .sort();
        expect(files).toEqual(rows.map((r) => r.file).sort());
        expect(new Set(rows.map((r) => r.element)).size).toBe(rows.length);
    });

    for (const row of rows) {
        it(`[FAIL] w:${row.element} (${row.mutation})`, { timeout: 60_000 }, async () => {
            const fixturePath = path.join(CORPUS, row.file);
            const bytes = readFileSync(fixturePath);
            expect(createHash("sha256").update(bytes).digest("hex"), "fixture bytes must match manifest sha256").toBe(row.sha256);

            const result = await validate(fixturePath, { profile: "lenient" });
            const errors = result.issues.filter((i) => i.severity === "error");
            const errorCodes = [...new Set(errors.map((i) => i.code).filter((c): c is string => Boolean(c)))].sort();

            expect(result.valid, `w:${row.element} fixture must be invalid under lenient`).toBe(false);
            expect(errorCodes).toContain("xsd-error");
            expect(errorCodes, "error-code drift vs manifest").toEqual(row.errorCodes);
            expect(
                errors.some((e) => e.message.includes(`{${W}}${row.element}`)),
                `an error must name the element {…}w:${row.element}`,
            ).toBe(true);
        });
    }
});
