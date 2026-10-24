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

import { resolveWordColumn, resolveWordOutcome } from "../scripts/update-manifest";
import type { WordHistoryEntry } from "../src/lib/word-provenance";

describe("resolveWordOutcome", () => {
    const prior = new Map<string, string>([["broken/a.docx", "clean-open"]]);

    it("prefers a fresh probe outcome", () => {
        expect(resolveWordOutcome("broken/a.docx", "recovered", prior)).toBe("recovered");
    });

    it("falls back to the prior manifest value when no probe", () => {
        expect(resolveWordOutcome("broken/a.docx", undefined, prior)).toBe("clean-open");
    });

    it("is 'unknown' when neither probe nor prior value exists", () => {
        expect(resolveWordOutcome("broken/new.docx", undefined, prior)).toBe("unknown");
    });
});

describe("resolveWordColumn", () => {
    it("writes hash-bound structured observations and never discards prior evidence (#274)", () => {
        // Unchanged hash: prior observation kept, now recorded bound to its bytes.
        expect(resolveWordColumn({ outcome: "clean-open", sha256: "aaaa" }, [], undefined, "aaaa")).toEqual({
            word: { outcome: "clean-open", sha256: "aaaa" },
            history: [],
        });

        // Changed hash, no fresh probe: prior retired to history under its old
        // hash; the current bytes carry no observation.
        expect(resolveWordColumn({ outcome: "open-error", sha256: "aaaa" }, [], undefined, "bbbb")).toEqual({
            word: { outcome: "unknown", sha256: "bbbb" },
            history: [{ outcome: "open-error", sha256: "aaaa" }],
        });

        // Fresh probe outcome wins and is bound to the current bytes.
        expect(resolveWordColumn({ outcome: "clean-open", sha256: "aaaa" }, [], "open-error", "bbbb")).toEqual({
            word: { outcome: "open-error", sha256: "bbbb" },
            history: [{ outcome: "clean-open", sha256: "aaaa" }],
        });

        // A legacy bare-string prior with no hash is preserved and adopts the
        // current hash (the one-time binding migration).
        expect(resolveWordColumn("clean-open", undefined, undefined, "aaaa")).toEqual({
            word: { outcome: "clean-open", sha256: "aaaa" },
            history: [],
        });

        // Pre-existing history survives every path.
        const history: WordHistoryEntry[] = [{ outcome: "timeout", sha256: "0000" }];
        expect(resolveWordColumn(undefined, history, undefined, "aaaa")).toEqual({
            word: { outcome: "unknown", sha256: "aaaa" },
            history,
        });
    });
});
