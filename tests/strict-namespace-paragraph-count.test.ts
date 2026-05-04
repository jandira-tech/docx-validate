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
 * Regression test for issue #138 — Strict OOXML paragraph counting.
 *
 * Port of `src/docx-validate/tests/test_strict_namespace_paragraph_count.py`
 * (task #16).
 *
 * Note: the Python `test_strict_format_detection_in_summary` used `capsys` to
 * grep stdout for "Strict OOXML conformance class". The TS port returns a
 * structured `ParagraphCounts` object (see `compareParagraphCounts` in
 * `validators/docx.ts`) instead of printing, so the assertion is reframed
 * against the `originalUsesStrictNamespace` field that drives that string.
 */

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli.ts";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx.ts";
import { STRICT_FIXTURE, makeEmptyPlaceholder, unpackDocxFixture } from "./_fixtures.ts";

describe("DOCXSchemaValidator strict-namespace paragraph counts", () => {
    it("counts one body paragraph in the unpacked strict-format fixture", async () => {
        await withTempDir(async (tmp) => {
            const unpacked = await unpackDocxFixture(STRICT_FIXTURE, tmp);
            const validator = new DOCXSchemaValidator({
                unpackedDir: unpacked,
                originalFile: STRICT_FIXTURE,
            });
            expect(validator.countParagraphsInUnpacked()).toBe(1);
        });
    });

    it("counts one body paragraph in the original strict-format docx", async () => {
        await withTempDir(async (tmp) => {
            const placeholder = await makeEmptyPlaceholder(tmp);
            const validator = new DOCXSchemaValidator({
                unpackedDir: placeholder,
                originalFile: STRICT_FIXTURE,
            });
            expect(await validator.countParagraphsInOriginal()).toBe(1);
        });
    });

    it("flags strict OOXML conformance in compareParagraphCounts result", async () => {
        await withTempDir(async (tmp) => {
            const unpacked = await unpackDocxFixture(STRICT_FIXTURE, tmp);
            const validator = new DOCXSchemaValidator({
                unpackedDir: unpacked,
                originalFile: STRICT_FIXTURE,
            });
            const summary = await validator.compareParagraphCounts();
            expect(summary.originalUsesStrictNamespace).toBe(true);
            expect(summary.original).toBe(1);
            expect(summary.modified).toBe(1);
            expect(summary.delta).toBe(0);
        });
    });
});
