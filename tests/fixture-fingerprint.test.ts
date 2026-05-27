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

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { fingerprint } from "../scripts/fixture-fingerprint";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");

describe("fingerprint", () => {
    it("captures strict errors and a content hash for a known-broken fixture", async () => {
        const fp = await fingerprint(path.join(FIXTURES, "broken/tables.missing-namespace.docx"));
        expect(fp.strictErrorCodes.length).toBeGreaterThan(0);
        expect(fp.contentHash).toMatch(/^[0-9a-f]{64}$/);
        expect([...fp.strictErrorCodes].sort()).toEqual(fp.strictErrorCodes);
    }, 20000);

    it("reports zero strict errors for a structurally-valid fixture", async () => {
        const fp = await fingerprint(path.join(FIXTURES, "external/docx-templates/exec.docx"));
        expect(fp.strictErrorCodes).toEqual([]);
        expect(fp.contentHash).toMatch(/^[0-9a-f]{64}$/);
    }, 20000);
});
