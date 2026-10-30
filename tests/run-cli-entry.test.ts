import { promises as fs } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { describe, expect, it, vi } from "vitest";

import { commanderExitCode, isCliEntry, runCli, withTempDir } from "../src/lib/run-cli";

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

    it("matches a real file's metaUrl against an argv1 that does not exist on disk", async () => {
        // argv1 that is not a file (or no longer exists): realpathSync throws,
        // so the comparison falls back to path.resolve — a non-file argv value
        // must not crash the guard.
        await withTempDir(async (dir) => {
            const file = path.join(dir, "script.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const missing = path.join(dir, "does-not-exist.ts");
            expect(isCliEntry(pathToFileURL(file).href, missing)).toBe(false);
            // Resolving the missing path to itself still matches its own spelling.
            expect(isCliEntry(pathToFileURL(missing).href, missing)).toBe(true);
        });
    });
});

describe("commanderExitCode — Commander error exit-code extraction", () => {
    it("returns a Commander error's numeric exitCode", () => {
        expect(commanderExitCode({ exitCode: 0 })).toBe(0);
        expect(commanderExitCode({ exitCode: 2, code: "commander.missingArgument" })).toBe(2);
    });

    it("returns 1 for non-Commander errors", () => {
        expect(commanderExitCode(new Error("boom"))).toBe(1);
        expect(commanderExitCode("boom")).toBe(1);
        expect(commanderExitCode(null)).toBe(1);
        expect(commanderExitCode({ exitCode: "not-a-number" })).toBe(1);
    });
});

describe("runCli — entry-point dispatch", () => {
    it("does not invoke fn when metaUrl is not the process entry", async () => {
        const fn = vi.fn(() => 0);
        runCli(pathToFileURL(path.join(process.cwd(), "not-the-entry.ts")).href, fn);
        await new Promise((resolve) => setTimeout(resolve, 10));
        expect(fn).not.toHaveBeenCalled();
    });

    it("exits with the code fn returns when metaUrl is the process entry", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "entry.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const argv1 = process.argv[1];
            const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
            process.argv[1] = file;
            try {
                runCli(pathToFileURL(file).href, async () => 7);
                await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(7));
            } finally {
                process.argv[1] = argv1;
            }
        });
    });

    it("exits 0 when fn resolves to a non-number", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "entry.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const argv1 = process.argv[1];
            const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
            process.argv[1] = file;
            try {
                runCli(pathToFileURL(file).href, () => undefined as unknown as number);
                await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(0));
            } finally {
                process.argv[1] = argv1;
            }
        });
    });

    it("logs an Error's stack and exits 1 when fn throws", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "entry.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const argv1 = process.argv[1];
            const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
            const written: string[] = [];
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation((chunk: unknown) => {
                written.push(String(chunk));
                return true;
            });
            process.argv[1] = file;
            try {
                runCli(pathToFileURL(file).href, () => {
                    throw new Error("kapow");
                });
                await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(1));
                expect(written.join("")).toContain("kapow");
                expect(written.join("")).toContain("Error");
            } finally {
                process.argv[1] = argv1;
                stderr.mockRestore();
            }
        });
    });

    it("logs an Error's message when it has no stack and exits 1", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "entry.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const argv1 = process.argv[1];
            const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
            const written: string[] = [];
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation((chunk: unknown) => {
                written.push(String(chunk));
                return true;
            });
            process.argv[1] = file;
            try {
                runCli(pathToFileURL(file).href, () => {
                    const error = new Error("no stack here");
                    error.stack = undefined;
                    throw error;
                });
                // An Error without a stack falls back to its message.
                await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(1));
                expect(written.join("")).toBe("no stack here\n");
            } finally {
                process.argv[1] = argv1;
                stderr.mockRestore();
            }
        });
    });

    it("stringifies a non-Error rejection and exits 1", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "entry.ts");
            await fs.writeFile(file, "// entry\n", "utf-8");
            const argv1 = process.argv[1];
            const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
            const written: string[] = [];
            const stderr = vi.spyOn(process.stderr, "write").mockImplementation((chunk: unknown) => {
                written.push(String(chunk));
                return true;
            });
            process.argv[1] = file;
            try {
                runCli(pathToFileURL(file).href, async () => Promise.reject("rejected string"));
                await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(1));
                expect(written.join("")).toContain("rejected string");
            } finally {
                process.argv[1] = argv1;
                stderr.mockRestore();
            }
        });
    });
});

describe("withTempDir — temporary-directory lifecycle", () => {
    it("removes the directory after fn succeeds", async () => {
        let dir = "";
        const result = await withTempDir(async (d) => {
            dir = d;
            await fs.writeFile(path.join(d, "file.txt"), "x", "utf-8");
            return 42;
        });
        expect(result).toBe(42);
        await expect(fs.access(dir)).rejects.toMatchObject({ code: "ENOENT" });
    });

    it("removes the directory when fn throws", async () => {
        let dir = "";
        await expect(
            withTempDir(async (d) => {
                dir = d;
                throw new Error("boom");
            }),
        ).rejects.toThrow("boom");
        await expect(fs.access(dir)).rejects.toMatchObject({ code: "ENOENT" });
    });
});
