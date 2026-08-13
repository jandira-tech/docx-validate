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

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { formatSweepError, walkDocxFiles } from "../scripts/sweep-word-valid";
import { withTempDir } from "../src/lib/run-cli";

describe("walkDocxFiles", () => {
    it("finds .docx files and skips symbolic-link directories", async () => {
        await withTempDir(async (dir) => {
            const real = path.join(dir, "real");
            const other = path.join(dir, "other");
            await fs.mkdir(real);
            await fs.mkdir(other);
            await fs.writeFile(path.join(real, "keep.docx"), "x");
            await fs.writeFile(path.join(other, "skip.docx"), "x");
            await fs.symlink(other, path.join(real, "loop"));

            const found = [...walkDocxFiles(real)].sort();
            expect(found).toEqual([path.join(real, "keep.docx")]);
        });
    });
});

describe("formatSweepError", () => {
    it("uses a path relative to the sweep root, not basename", () => {
        expect(formatSweepError("/tmp/root", "/tmp/root/a/foo.docx", "code msg")).toBe("a/foo.docx: code msg");
    });
});
