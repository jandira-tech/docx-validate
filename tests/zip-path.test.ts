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

import { existsSync, promises as fsPromises, readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { escapeControlCharacters, extractPackageParts, extractZipEntries, findLinkBelow, resolveSafeZipEntry } from "../src/lib/zip-path";
import { poisonEntry } from "./support/poison-entry";
import { buildRawZip } from "./support/raw-zip";

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
        expect(() => resolveSafeZipEntry(out, "C:\\Windows\\System32\\cmd.exe")).toThrow(/Refusing to extract entry outside output dir/);
        expect(() => resolveSafeZipEntry(out, "C:/Windows/System32/cmd.exe")).toThrow(/Refusing to extract entry outside output dir/);
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

    it("escapes control characters in the refused name, so it cannot drive the terminal that prints it (#675)", () => {
        const out = path.resolve("/tmp/docx-out");
        expect(() => resolveSafeZipEntry(out, "../\u001b[2J\nevil\u202e.txt")).toThrow(
            "Refusing to extract entry outside output dir: ../\\u{1b}[2J\\u{a}evil\\u{202e}.txt",
        );
    });
});

describe("escapeControlCharacters", () => {
    it("leaves printable names unchanged, including non-ASCII letters", () => {
        expect(escapeControlCharacters("word/média/图像 1.png")).toBe("word/média/图像 1.png");
    });

    it("escapes C0 and C1 controls, DEL and bidirectional formatting characters", () => {
        expect(escapeControlCharacters("a\u0000b\tc\u007fd\u0085e\u200ff\u2066g")).toBe(
            "a\\u{0}b\\u{9}c\\u{7f}d\\u{85}e\\u{200f}f\\u{2066}g",
        );
    });
});

describe("extractZipEntries (#675)", () => {
    // Raw archives store only the named entries: JSZip's own writer adds a
    // folder entry for every parent, and one for `../` loads as `/`.
    const zipOf = (entries: readonly (readonly [string, string])[]): Promise<JSZip> =>
        JSZip.loadAsync(buildRawZip(entries.map(([name, content]) => ({ name, content }))));
    const entriesUnder = (dir: string): string[] =>
        readdirSync(dir, { recursive: true, withFileTypes: true })
            .map((entry) => path.relative(dir, path.join(entry.parentPath, entry.name)).split(path.sep).join("/"))
            .sort();

    it.each(["../evil.txt", "word/../../evil.txt"])(
        "refuses the stored name %j, which JSZip loads as evil.txt inside the output directory (#676)",
        async (name) => {
            await withTempDir(async (dir) => {
                const out = path.join(dir, "out");
                await fsPromises.mkdir(out);
                const extraction = extractZipEntries(await zipOf([[name, "x"]]), out);
                await expect(extraction).rejects.toThrow(`Refusing to extract entry outside output dir: ${name}`);
                expect(entriesUnder(out)).toEqual([]);
            });
        },
    );

    it("keeps extracting a stored name that resolves inside the output directory", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            await extractZipEntries(await zipOf([["a/../b.txt", "x"]]), out);
            expect(readFileSync(path.join(out, "b.txt"), "utf8")).toBe("x");
        });
    });

    it("writes nothing when a later entry is refused", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const zip = await zipOf([
                ["word/document.xml", "<d/>"],
                ["docProps/core.xml", "<c/>"],
                ["../evil.txt", "x"],
            ]);
            await expect(extractZipEntries(zip, out)).rejects.toThrow(/outside output dir/);
            expect(existsSync(out)).toBe(false);
        });
    });

    it("refuses to write through a symbolic link already below the output directory, writing nothing", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const elsewhere = path.join(dir, "elsewhere");
            await fsPromises.mkdir(out);
            await fsPromises.mkdir(elsewhere);
            await fsPromises.symlink(elsewhere, path.join(out, "word"), "dir");
            const zip = await zipOf([
                ["[Content_Types].xml", "<t/>"],
                ["word/document.xml", "<d/>"],
            ]);
            await expect(extractZipEntries(zip, out)).rejects.toThrow("Refusing to extract word/document.xml: word is a symbolic link");
            expect(entriesUnder(elsewhere)).toEqual([]);
            expect(entriesUnder(out)).toEqual(["word"]);
        });
    });

    it("refuses to write over a symbolic link that names the entry itself", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const target = path.join(dir, "target.txt");
            await fsPromises.mkdir(out);
            await fsPromises.writeFile(target, "keep");
            await fsPromises.symlink(target, path.join(out, "evil.xml"), "file");
            await expect(extractZipEntries(await zipOf([["evil.xml", "x"]]), out)).rejects.toThrow(
                "Refusing to extract evil.xml: evil.xml is a symbolic link",
            );
            expect(readFileSync(target, "utf8")).toBe("keep");
        });
    });

    it("extracts into an output directory that is itself a symbolic link: the caller chose it", async () => {
        await withTempDir(async (dir) => {
            const real = path.join(dir, "real");
            const out = path.join(dir, "out");
            await fsPromises.mkdir(real);
            await fsPromises.symlink(real, out, "dir");
            await extractZipEntries(await zipOf([["word/document.xml", "<d/>"]]), out);
            expect(readFileSync(path.join(real, "word", "document.xml"), "utf8")).toBe("<d/>");
        });
    });

    it("removes what it created, and only that, when a later entry cannot be read", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            await fsPromises.mkdir(path.join(out, "word"), { recursive: true });
            await fsPromises.writeFile(path.join(out, "notes.txt"), "mine");
            await fsPromises.writeFile(path.join(out, "word", "keep.xml"), "mine");
            const zip = new JSZip();
            zip.file("word/document.xml", "<d/>");
            zip.file("word/media/image1.png", "png");
            zip.file("docProps/core.xml", "<c/>");
            zip.file("z/broken.xml", "<b/>".repeat(64));
            const archive = poisonEntry(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), "z/broken.xml");
            await expect(extractZipEntries(await JSZip.loadAsync(archive), out)).rejects.toThrow();
            expect(entriesUnder(out)).toEqual(["notes.txt", "word", "word/keep.xml"]);
        });
    });
});

describe("findLinkBelow (#675, #939)", () => {
    it("returns the first link on the path to a destination, and extractZipEntries still refuses with its old message", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const elsewhere = path.join(dir, "elsewhere");
            await fsPromises.mkdir(out);
            await fsPromises.mkdir(elsewhere);
            await fsPromises.symlink(elsewhere, path.join(out, "word"), "dir");
            expect(await findLinkBelow(out, path.join(out, "word", "document.xml"))).toBe(path.join(out, "word"));
            const zip = await JSZip.loadAsync(await buildRawZip([{ name: "word/document.xml", content: "<d/>" }]));
            await expect(extractZipEntries(zip, out)).rejects.toThrow("Refusing to extract word/document.xml: word is a symbolic link");
        });
    });

    it("returns the link when the destination itself is one", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const target = path.join(dir, "target.xml");
            await fsPromises.mkdir(out);
            await fsPromises.writeFile(target, "keep");
            await fsPromises.symlink(target, path.join(out, "evil.xml"), "file");
            expect(await findLinkBelow(out, path.join(out, "evil.xml"))).toBe(path.join(out, "evil.xml"));
        });
    });

    it("returns undefined when no segment of the path is a link", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            await fsPromises.mkdir(path.join(out, "word"), { recursive: true });
            await fsPromises.writeFile(path.join(out, "word", "document.xml"), "<d/>");
            expect(await findLinkBelow(out, path.join(out, "word", "document.xml"))).toBeUndefined();
        });
    });

    it("returns undefined when a segment does not exist yet: nothing below it is on disk", async () => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            await fsPromises.mkdir(out);
            expect(await findLinkBelow(out, path.join(out, "word", "document.xml"))).toBeUndefined();
        });
    });

    it("leaves the output directory itself alone: the caller chose it", async () => {
        await withTempDir(async (dir) => {
            const real = path.join(dir, "real");
            const out = path.join(dir, "out");
            await fsPromises.mkdir(real);
            await fsPromises.symlink(real, out, "dir");
            expect(await findLinkBelow(out, path.join(out, "word", "document.xml"))).toBeUndefined();
        });
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
    ])("refuses colliding part spellings before writing on every host (%s, #373, #441)", async (_relation, first, second) => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const extraction = extractPackageParts(
                await zipOf([
                    [`word/${first}`, "<first/>"],
                    [`word/${second}`, "<second/>"],
                ]),
                out,
            );
            await expect(extraction).rejects.toThrow(/collide under portable ASCII-case\/NFC folding/);
            expect(existsSync(out)).toBe(false);
        });
    });

    it("walks its output with no readdir feature newer than Node 20.0", async () => {
        // `recursive` arrived in Node 20.1 and `Dirent.parentPath` in 20.12;
        // the package supports Node 20 (CodeRabbit on #379).
        const readdir = fsPromises.readdir.bind(fsPromises);
        const spy = vi.spyOn(fsPromises, "readdir").mockImplementation((async (dir: string, options?: { recursive?: boolean }) => {
            if (options?.recursive === true) {
                throw new Error("readdir: recursive needs Node 20.1");
            }
            const entries = await readdir(dir, { withFileTypes: true });
            return entries.map((entry) => Object.defineProperty(entry, "parentPath", { value: undefined }));
        }) as unknown as typeof fsPromises.readdir);
        try {
            await withTempDir(async (dir) => {
                const out = path.join(dir, "out");
                await extractPackageParts(
                    await zipOf([
                        ["word/document.xml", "<d/>"],
                        ["word/_rels/document.xml.rels", "<r/>"],
                    ]),
                    out,
                );
                expect(readFileSync(path.join(out, "word", "_rels", "document.xml.rels"), "utf8")).toBe("<r/>");
            });
            expect(spy).toHaveBeenCalled();
        } finally {
            spy.mockRestore();
        }
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
    ])("refuses colliding folder spellings before writing on every host (%s, #376, #441)", async (_relation, first, second) => {
        await withTempDir(async (dir) => {
            const out = path.join(dir, "out");
            const extraction = extractPackageParts(
                await zipOf([
                    [`${first}/a.xml`, "<a/>"],
                    [`${second}/b.xml`, "<b/>"],
                ]),
                out,
            );
            await expect(extraction).rejects.toThrow(/collide under portable ASCII-case\/NFC folding/);
            expect(existsSync(out)).toBe(false);
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
