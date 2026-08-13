import { promises as fs } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { describe, expect, it } from "vitest";

import { isCliEntry, withTempDir } from "../src/lib/run-cli";

describe("isCliEntry — process entry-point detection", () => {
    it("matches when metaUrl and argv1 are the identical absolute path", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "script.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            expect(isCliEntry(pathToFileURL(file).href, file)).toBe(true);
        });
    });

    it("matches when argv1 is a relative path resolving to the same file", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "script.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const rel = path.relative(process.cwd(), file);
            expect(isCliEntry(pathToFileURL(file).href, rel)).toBe(true);
        });
    });

    // The real bug: Node resolves `import.meta.url` through symlinks (realpath),
    // but `process.argv[1]` keeps the symlinked spelling. A raw string compare
    // then mismatches and the CLI silently no-ops. (Reproduces /Users/arthrod/T
    // -> /Users/arthrod/temp/T breaking python-jubarte's docx-validate calls.)
    it("matches when argv1 reaches the same file through a symlinked directory", async () => {
        await withTempDir(async (dir) => {
            const realDir = path.join(dir, "real");
            await fs.mkdir(realDir, { recursive: true });
            const realFile = path.join(realDir, "script.ts");
            await fs.writeFile(realFile, "// entry\n", "utf-8");

            const linkDir = path.join(dir, "link");
            await fs.symlink(realDir, linkDir, "dir");
            const argvSpelling = path.join(linkDir, "script.ts"); // un-dereferenced

            // metaUrl is the realpath (as Node provides it); argv1 is via the symlink.
            expect(isCliEntry(pathToFileURL(realFile).href, argvSpelling)).toBe(true);
        });
    });

    it("does not match two different files", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a.ts");
            const b = path.join(dir, "b.ts");
            await fs.writeFile(a, "// a\n", "utf-8");
            await fs.writeFile(b, "// b\n", "utf-8");
            expect(isCliEntry(pathToFileURL(a).href, b)).toBe(false);
        });
    });

    it("returns false when argv1 is missing", () => {
        expect(isCliEntry(pathToFileURL(path.join(process.cwd(), "x.ts")).href, undefined)).toBe(false);
    });
});
