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
 * Shared test helpers for the ported pytest suite (task #16).
 *
 * The Python conftest just put `scripts/office` on sys.path; the per-file
 * fixtures unzipped a `.docx` into `tmp_path/unpacked` before constructing
 * the validator. Both pieces are bundled here so individual test files stay
 * lean and the fixture-path resolution lives in one place.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import { extractZipEntries } from "../src/lib/zip-path";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const MAMMOTH_FIXTURES = path.join(HERE, "fixtures", "external", "mammoth-js");

export const STRICT_FIXTURE = path.join(MAMMOTH_FIXTURES, "strict-format.docx");

export const TEXTBOX_FIXTURE = path.join(MAMMOTH_FIXTURES, "text-box.docx");

/**
 * Extract a `.docx` (zip) into `<tmpDir>/unpacked` and return that path.
 * Mirrors the per-test pytest fixture body: `tmp_path / "unpacked"` then
 * `zipfile.ZipFile(...).extractall(target)`.
 */
export async function unpackDocxFixture(docxPath: string, tmpDir: string): Promise<string> {
    const target = path.join(tmpDir, "unpacked");
    await fs.mkdir(target, { recursive: true });

    const buf = await fs.readFile(docxPath);
    await extractZipEntries(await JSZip.loadAsync(buf), target);

    return target;
}

/**
 * Mirror pytest's `tmp_path / "empty"` placeholder used when a test needs an
 * unpacked-dir argument it never reads from.
 */
export async function makeEmptyPlaceholder(tmpDir: string): Promise<string> {
    const placeholder = path.join(tmpDir, "empty");
    await fs.mkdir(placeholder, { recursive: true });
    return placeholder;
}
