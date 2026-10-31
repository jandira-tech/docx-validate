/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/*
 * `getGitWordDiff` runs `git diff` through `spawnSync`, which blocks the
 * event loop, on text extracted from the documents under test. Both
 * attempts share one time budget and capture at most 10 MiB of output
 * (bot PRs #207, #204 and #264; the shared budget answers a review
 * comment on #207: two fresh 10-second timeouts could block for 20).
 */

import * as childProcess from "node:child_process";

import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("node:child_process", async (importOriginal) => {
    const actual = await importOriginal<typeof import("node:child_process")>();
    return { ...actual, spawnSync: vi.fn(actual.spawnSync) };
});

import { getGitWordDiff } from "../src/scripts/office/validators/redlining";

const spawnSyncMock = vi.mocked(childProcess.spawnSync);

const timedOut = (): ReturnType<typeof childProcess.spawnSync> =>
    ({
        pid: 0,
        output: [],
        stdout: "",
        stderr: "",
        status: null,
        signal: "SIGTERM",
        error: Object.assign(new Error("spawnSync git ETIMEDOUT"), { code: "ETIMEDOUT" }),
    }) as unknown as ReturnType<typeof childProcess.spawnSync>;

const optionsOfCall = (n: number): childProcess.SpawnSyncOptions => spawnSyncMock.mock.calls[n]?.[2] as childProcess.SpawnSyncOptions;

afterEach(() => {
    spawnSyncMock.mockClear();
    vi.useRealTimers();
});

describe("getGitWordDiff execution limits", () => {
    it("runs git with a timeout and a 10 MiB output cap", async () => {
        const result = await getGitWordDiff("the quick brown fox\n", "the quick red fox\n");

        expect(result).toBe("the quick [-b-]r[-own-]{+ed+} fox");
        expect(spawnSyncMock).toHaveBeenCalled();
        const options = optionsOfCall(0);
        expect(options.timeout).toBeGreaterThan(0);
        expect(options.timeout).toBeLessThanOrEqual(10_000);
        expect(options.maxBuffer).toBe(10 * 1024 * 1024);
    });

    it("gives the fallback attempt only the time the first attempt left", async () => {
        vi.useFakeTimers({ toFake: ["Date"] });
        spawnSyncMock.mockImplementationOnce(() => {
            vi.setSystemTime(Date.now() + 9_000);
            return timedOut();
        });
        spawnSyncMock.mockImplementationOnce(() => timedOut());

        expect(await getGitWordDiff("a\n", "b\n")).toBeNull();

        expect(spawnSyncMock).toHaveBeenCalledTimes(2);
        expect(optionsOfCall(1).timeout).toBeGreaterThan(0);
        expect(optionsOfCall(1).timeout).toBeLessThanOrEqual(1_000);
    });

    it("skips the fallback attempt once the budget is spent", async () => {
        vi.useFakeTimers({ toFake: ["Date"] });
        spawnSyncMock.mockImplementationOnce(() => {
            vi.setSystemTime(Date.now() + 10_000);
            return timedOut();
        });

        expect(await getGitWordDiff("a\n", "b\n")).toBeNull();

        expect(spawnSyncMock).toHaveBeenCalledTimes(1);
    });

    it("falls back to the word-level diff when the granular diff filters to nothing", async () => {
        const plainOutput = (stdout: string): ReturnType<typeof childProcess.spawnSync> =>
            ({ pid: 1, output: [], stdout, stderr: "", status: 1, signal: null }) as unknown as ReturnType<typeof childProcess.spawnSync>;
        spawnSyncMock.mockImplementationOnce(() => plainOutput("diff --git a/original.txt b/modified.txt\nindex 1..2 100644\n"));
        spawnSyncMock.mockImplementationOnce(() => plainOutput("@@ -1 +1 @@\n[-a-]{+b+}\n"));

        expect(await getGitWordDiff("a\n", "b\n")).toBe("[-a-]{+b+}");
        expect(spawnSyncMock).toHaveBeenCalledTimes(2);
        const [, granularArgs] = spawnSyncMock.mock.calls[0]!;
        expect(spawnSyncMock).toHaveBeenNthCalledWith(
            2,
            "git",
            ["diff", "--word-diff=plain", "-U0", "--no-index", granularArgs![5], granularArgs![6]],
            expect.any(Object),
        );
    });

    it("returns null when spawning git throws", async () => {
        spawnSyncMock.mockImplementationOnce(() => {
            throw new Error("boom");
        });
        spawnSyncMock.mockImplementationOnce(() => {
            throw new Error("boom");
        });

        expect(await getGitWordDiff("a\n", "b\n")).toBeNull();
        expect(spawnSyncMock).toHaveBeenCalledTimes(2);
    });

    it("treats missing stdout as empty output", async () => {
        const noStdout = (): ReturnType<typeof childProcess.spawnSync> =>
            ({ pid: 1, output: [], stdout: null, stderr: "", status: 1, signal: null }) as unknown as ReturnType<
                typeof childProcess.spawnSync
            >;
        spawnSyncMock.mockImplementationOnce(noStdout).mockImplementationOnce(noStdout);

        expect(await getGitWordDiff("a\n", "b\n")).toBeNull();
        expect(spawnSyncMock).toHaveBeenCalledTimes(2);
    });

    it("runs the real git again after a stubbed case (Cursor on #1205)", async () => {
        expect(await getGitWordDiff("a\n", "b\n")).not.toBeNull();
    });
});
