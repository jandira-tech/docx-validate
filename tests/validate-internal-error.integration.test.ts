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
 * An internal error after the archive loads is not a defect in the user's
 * file: `ValidationContext.fromArchive` reports only the two failures a real
 * extraction reports — a refused item name and an unreadable entry — as
 * `package-extract-failed` (#427). Anything else propagates as an internal
 * failure instead of wearing a file-defect code.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { poisonEntry } from "./support/poison-entry";
import { buildRawZip } from "./support/raw-zip";

describe("an internal error after the archive loads (#427)", { timeout: 120_000 }, () => {
    it("a thrown TypeError in the package view propagates instead of reporting package-extract-failed", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            const bytes = await buildDocx();
            writeFileSync(file, bytes);
            // The loaded view's own accessor throws, as a JSZip regression or
            // our own post-load code would. A load failure reads differently:
            // `loadAsync` rejects, which is the archive's open failure.
            const stub = Object.create(JSZip.prototype) as JSZip;
            Object.defineProperty(stub, "files", {
                get: () => {
                    throw new TypeError("Cannot read properties of undefined (reading 'files')");
                },
            });
            const spy = vi.spyOn(JSZip, "loadAsync").mockResolvedValue(stub);
            try {
                // The packed path reads `zip.files` outside any wrapper, so it
                // already propagates this error as an internal failure.
                await expect(validate(file)).rejects.toBeInstanceOf(TypeError);
                await expect(validate(file, { quick: true })).rejects.toBeInstanceOf(TypeError);
            } finally {
                spy.mockRestore();
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("keeps a refused item name as package-extract-failed", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "refused.docx");
            const bytes = buildRawZip([
                ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                { name: "word/a/../b.xml", content: "<x/>" },
            ]);
            writeFileSync(file, bytes);
            for (const selection of [{}, { quick: true }] as const) {
                const result = await validate(file, selection);
                expect(result.issues.map((issue) => issue.code)).toContain("package-extract-failed");
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("keeps an unreadable entry as package-extract-failed", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "poisoned.docx");
            const bytes = poisonEntry(await buildDocx(), "word/document.xml");
            writeFileSync(file, bytes);
            for (const selection of [{}, { quick: true }] as const) {
                const result = await validate(file, selection);
                expect(result.issues.map((issue) => issue.code)).toContain("package-extract-failed");
            }
            expect(readFileSync(file)).toEqual(Buffer.from(bytes));
        });
    });
});
