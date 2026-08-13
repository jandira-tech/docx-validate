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

import { describe, expect, it } from "vitest";

import { resolveSafeZipEntry } from "../src/lib/zip-path";

describe("resolveSafeZipEntry", () => {
    it("returns the resolved path for a nested package entry", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(resolveSafeZipEntry(out, "word/document.xml")).toBe(path.resolve(out, "word/document.xml"));
    });

    it("allows the output directory itself", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(resolveSafeZipEntry(out, ".")).toBe(out);
    });

    it("rejects parent-directory traversal", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(() => resolveSafeZipEntry(out, "../../etc/passwd")).toThrow(/Refusing to extract entry outside output dir: \.\.\/\.\.\/etc\/passwd/);
    });

    it("rejects a sibling directory that shares a prefix (startsWith trap)", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(() => resolveSafeZipEntry(out, "../docx-out-evil/x")).toThrow(/Refusing to extract entry outside output dir/);
    });

    it("rejects an absolute entry name", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(() => resolveSafeZipEntry(out, "/etc/passwd")).toThrow(/Refusing to extract entry outside output dir/);
    });

    it("resolves a relative output dir before the containment check", () => {
        const out = "relative-out";
        const resolved = resolveSafeZipEntry(out, "word/document.xml");
        expect(resolved).toBe(path.resolve(out, "word/document.xml"));
        expect(path.isAbsolute(resolved)).toBe(true);
    });

    it("allows a contained entry whose name begins with dots", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(resolveSafeZipEntry(out, "..metadata/item.xml")).toBe(path.resolve(out, "..metadata/item.xml"));
    });
});
