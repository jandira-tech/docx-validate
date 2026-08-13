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

import * as crypto from "node:crypto";

import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("node:crypto", async (importOriginal) => {
    const actual = await importOriginal<typeof import("node:crypto")>();
    return {
        ...actual,
        randomInt: vi.fn(actual.randomInt),
    };
});

import { nextSecureLongHexNumber, SECURE_LONG_HEX_EXCLUSIVE_MAX } from "../src/lib/secure-id";

describe("nextSecureLongHexNumber", () => {
    afterEach(() => {
        vi.mocked(crypto.randomInt).mockReset();
    });

    it("delegates to crypto.randomInt(1, SECURE_LONG_HEX_EXCLUSIVE_MAX)", () => {
        vi.mocked(crypto.randomInt).mockImplementation(() => 0x10);
        expect(nextSecureLongHexNumber()).toBe(0x10);
        expect(crypto.randomInt).toHaveBeenCalledWith(1, SECURE_LONG_HEX_EXCLUSIVE_MAX);
    });
});
