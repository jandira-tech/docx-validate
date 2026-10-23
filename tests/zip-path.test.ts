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

import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { extractPackageParts, resolveSafeZipEntry } from "../src/lib/zip-path";
import { poisonEntry } from "./support/poison-entry";

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
        expect(() => resolveSafeZipEntry(out, "../../etc/passwd")).toThrow(
            /Refusing to extract entry outside output dir: \.\.\/\.\.\/etc\/passwd/,
        );
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

describe("extractPackageParts", () => {
    const zipOf = async (entries: readonly (readonly [string, string])[]): Promise<JSZip> => {
        const zip = new JSZip();
        for (const [name, content] of entries) {
            zip.file(name, content);
        }
        return JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
    };
    const filesUnder = (dir: string): string[] =>
        readdirSync(dir, { recursive: true, withFileTypes: true })
            .filter((entry) => entry.isFile())
            .map((entry) => path.relative(dir, path.join(entry.parentPath, entry.name)).split(path.sep).join("/"))
            .sort();
    // Whether this filesystem holds the two file names as one file.
    const holdsAsOne = (dir: string, a: string, b: string): boolean => {
        const probe = path.join(dir, "probe");
        mkdirSync(probe);
        writeFileSync(path.join(probe, a), "");
        return existsSync(path.join(probe, b));
    };

    it("writes each part once, under its part name, with the bytes of the entry the package view keeps (#371)", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const zip = await zipOf([
                ["word/document.xml", "<kept/>"],
                ["word\\header1.xml", "<hdr/>"],
                ["word\\document.xml", "<stray/>"],
            ]);
            await extractPackageParts(zip, out);
            expect(filesUnder(out)).toEqual(["word/document.xml", "word/header1.xml"]);
            expect(readFileSync(path.join(out, "word", "document.xml"), "utf8")).toBe("<kept/>");
        });
    });

    it("reads every entry, so one it drops for a duplicate name still rejects the package when unreadable (#369)", async () => {
        await withTempDir(async (dir) => {
            const zip = new JSZip();
            zip.file("word/document.xml", "<kept/>");
            zip.file("word\\document.xml", "<stray/>");
            const archive = poisonEntry(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), "word\\document.xml");
            await expect(extractPackageParts(await JSZip.loadAsync(archive), path.join(dir, "out"))).rejects.toThrow();
        });
    });

    it.each(["/word/bonus.xml", "\\word\\bonus.xml", "C:/bonus.xml", "c:\\bonus.xml"])(
        "refuses the item name %j, which places no part (#372)",
        async (name) => {
            await withTempDir(async (dir) => {
                const zip = await zipOf([
                    ["word/document.xml", "<d/>"],
                    [name, "<b/>"],
                ]);
                await expect(extractPackageParts(zip, path.join(dir, "out"))).rejects.toThrow(/names no package part/);
            });
        },
    );

    it.each([
        ["Unicode normalization", "caf\u00e9.xml", "cafe\u0301.xml"],
        ["case", "a.xml", "A.xml"],
    ])("never writes two parts this filesystem holds as one file (%s, #373)", async (_relation, first, second) => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const merges = holdsAsOne(dir, first, second);
            const extraction = extractPackageParts(
                await zipOf([
                    [`word/${first}`, "<first/>"],
                    [`word/${second}`, "<second/>"],
                ]),
                out,
            );
            if (merges) {
                await expect(extraction).rejects.toThrow(/already holds/);
            } else {
                await extraction;
                expect(readFileSync(path.join(out, "word", first), "utf8")).toBe("<first/>");
                expect(readFileSync(path.join(out, "word", second), "utf8")).toBe("<second/>");
            }
        });
    });

    it("passes on any other write failure unchanged", async () => {
        await withTempDir(async (dir) => {
            // No common filesystem takes a 300-character file name.
            const extraction = extractPackageParts(await zipOf([[`word/${"a".repeat(300)}.xml`, "<a/>"]]), path.join(dir, "out"));
            await expect(extraction).rejects.toThrow();
            await expect(extraction).rejects.not.toThrow(/already holds/);
        });
    });

    it.each([
        ["case", "word", "WORD"],
        ["Unicode normalization", "caf\u00e9", "cafe\u0301"],
    ])("never renames a part through a folder this filesystem folds (%s, #376)", async (_relation, first, second) => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const merges = holdsAsOne(dir, first, second);
            const extraction = extractPackageParts(
                await zipOf([
                    [`${first}/a.xml`, "<a/>"],
                    [`${second}/b.xml`, "<b/>"],
                ]),
                out,
            );
            if (merges) {
                await expect(extraction).rejects.toThrow(/under another name/);
            } else {
                await extraction;
                expect(filesUnder(out)).toEqual([`${first}/a.xml`, `${second}/b.xml`].sort());
            }
        });
    });

    it.each(["/../word/document.xml", "./word/x.xml", "word/./x.xml", "word//x.xml", "docProps\\nested\\..\\custom.xml"])(
        "refuses the stored item name %j, which JSZip or extraction would rename (#378)",
        async (name) => {
            await withTempDir(async (dir) => {
                const zip = await zipOf([
                    ["word/styles.xml", "<s/>"],
                    [name, "<x/>"],
                ]);
                await expect(extractPackageParts(zip, path.join(dir, "out"))).rejects.toThrow(/names no package part/);
            });
        },
    );
});
