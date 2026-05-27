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

import { describe, expect, it } from "vitest";

import { type FingerprintedFile, planMoves } from "../scripts/apply-fixture-names";

const fp = (overrides: Partial<FingerprintedFile["fingerprint"]> = {}) => ({
    strictErrorCodes: [] as string[],
    lenientErrorCodes: [] as string[],
    insCount: 0,
    delCount: 0,
    commentCount: 0,
    firstCommentText: null,
    tableCount: 0,
    hasTextBox: false,
    hasHeaderFooter: false,
    titleText: null,
    contentHash: "h0",
    ...overrides,
});

describe("planMoves", () => {
    it("renames in place (keeps source dir) when intoCategories is false", () => {
        const moves = planMoves(
            [{ sourcePath: "tests/fixtures/word-strict/Ouch.docx", fingerprint: fp({ insCount: 1, contentHash: "a" }) }],
            { intoCategories: false, fixturesRoot: "tests/fixtures" },
        );
        expect(moves).toEqual([
            { from: "tests/fixtures/word-strict/Ouch.docx", to: "tests/fixtures/word-strict/document.suggesting-insertions.docx" },
        ]);
    });

    it("routes into category dirs when intoCategories is true", () => {
        const moves = planMoves(
            [{ sourcePath: "fixtures/eigen-extended/Untitled (1).docx", fingerprint: fp({ strictErrorCodes: ["x-code"], contentHash: "b" }) }],
            { intoCategories: true, fixturesRoot: "tests/fixtures" },
        );
        expect(moves[0].to).toBe("tests/fixtures/broken/document.x-code.docx");
    });

    it("disambiguates colliding target names with a numeric suffix", () => {
        const moves = planMoves(
            [
                { sourcePath: "a.docx", fingerprint: fp({ insCount: 1, contentHash: "c1" }) },
                { sourcePath: "b.docx", fingerprint: fp({ insCount: 1, contentHash: "c2" }) },
            ],
            { intoCategories: true, fixturesRoot: "tests/fixtures" },
        );
        expect(moves.map((m) => m.to)).toEqual([
            "tests/fixtures/working/document.suggesting-insertions.docx",
            "tests/fixtures/working/document.suggesting-insertions-2.docx",
        ]);
    });

    it("drops content duplicates, keeping the first source path", () => {
        const moves = planMoves(
            [
                { sourcePath: "b.docx", fingerprint: fp({ insCount: 1, contentHash: "dup" }) },
                { sourcePath: "a.docx", fingerprint: fp({ insCount: 1, contentHash: "dup" }) },
            ],
            { intoCategories: true, fixturesRoot: "tests/fixtures", dedup: true },
        );
        expect(moves).toHaveLength(1);
        expect(moves[0].from).toBe("a.docx"); // lexicographically first kept
    });
});
