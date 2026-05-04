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
 * Regression test for issue #132 — text-box paragraph counting.
 *
 * Port of `src/docx-validate/tests/test_textbox_paragraph_count.py` (task #16).
 *
 * The original `text-box.docx` has one body paragraph hosting a floating text
 * box; the box itself contains another paragraph with the visible text "Datum
 * plane". Older `count_paragraphs_*` recursed into both `<w:txbxContent>`
 * (DML) and `<v:textbox>` (VML fallback), so the fixture reported 3
 * paragraphs even though the body has 1. This test pins the contract: count
 * body paragraphs only, ignoring those nested inside text boxes.
 */

import { existsSync } from "node:fs";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { makeEmptyPlaceholder, TEXTBOX_FIXTURE, unpackDocxFixture } from "./_fixtures";

const FIXTURE_PRESENT = existsSync(TEXTBOX_FIXTURE);

describe.skipIf(!FIXTURE_PRESENT)("DOCXSchemaValidator text-box paragraph counts", () => {
    it("does not count text-box inner paragraphs as body paragraphs (unpacked)", async () => {
        await withTempDir(async (tmp) => {
            const unpacked = await unpackDocxFixture(TEXTBOX_FIXTURE, tmp);
            const validator = new DOCXSchemaValidator({
                unpackedDir: unpacked,
                originalFile: TEXTBOX_FIXTURE,
            });
            expect(validator.countParagraphsInUnpacked()).toBe(1);
        });
    });

    it("does not count text-box inner paragraphs as body paragraphs (original)", async () => {
        await withTempDir(async (tmp) => {
            const placeholder = await makeEmptyPlaceholder(tmp);
            const validator = new DOCXSchemaValidator({
                unpackedDir: placeholder,
                originalFile: TEXTBOX_FIXTURE,
            });
            expect(await validator.countParagraphsInOriginal()).toBe(1);
        });
    });
});
