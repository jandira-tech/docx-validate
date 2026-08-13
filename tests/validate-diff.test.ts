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

import { errorMessage, parseFlag } from "../scripts/validate-diff";

describe("parseFlag", () => {
    it("does not treat the next dashed token as a value", () => {
        expect(parseFlag(["--baseline", "--current", "cur.json"], "--baseline")).toBeUndefined();
        expect(parseFlag(["--baseline", "--current", "cur.json"], "--current")).toBe("cur.json");
    });

    it("returns the following non-flag token", () => {
        expect(parseFlag(["--out", "snap.json"], "--out")).toBe("snap.json");
    });
});

describe("errorMessage", () => {
    it("stringifies non-Error throws", () => {
        expect(errorMessage("nope")).toBe("nope");
        expect(errorMessage(new Error("boom"))).toBe("boom");
    });
});
