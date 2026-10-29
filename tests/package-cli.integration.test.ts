/*
 * Copyright 2026 Jandira Technologies, LLC
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
 * Packaged-CLI surface of the built bundle (#314).
 *
 * Each source CLI module ends with `runCli(import.meta.url, …)`. Bundling
 * them into one `dist/index.mjs` gives every shim the same module URL, so a
 * single launch used to start every CLI at once (`--help` printed several
 * usages). The package ships a dedicated `dist/cli.mjs` validator entry
 * instead, and the library bundle starts no CLI as a side effect.
 *
 * This builds the package once (beforeAll) and then runs the emitted files
 * with plain `node`:
 *
 *   pnpm exec vitest run tests/package-cli.integration.test.ts
 */

import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { beforeAll, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const DIST = path.join(REPO, "dist");

interface RunOutcome {
    code: number;
    stdout: string;
    stderr: string;
}

/** Run `node <args>` and report the exit code instead of throwing on it. */
const run = async (args: readonly string[], timeout = 60_000): Promise<RunOutcome> => {
    try {
        const { stdout, stderr } = await execFileAsync(process.execPath, args, {
            cwd: REPO,
            timeout,
            maxBuffer: 16 * 1024 * 1024,
        });
        return { code: 0, stdout, stderr };
    } catch (err) {
        const failed = err as { code?: number; stdout?: string; stderr?: string };
        return {
            code: typeof failed.code === "number" ? failed.code : -1,
            stdout: failed.stdout ?? "",
            stderr: failed.stderr ?? String(err),
        };
    }
};

const usageCount = (text: string): number => (text.match(/Usage:/g) ?? []).length;

describe("packaged CLI (#314)", () => {
    beforeAll(async () => {
        await execFileAsync("pnpm", ["run", "build"], { cwd: REPO, timeout: 300_000, maxBuffer: 16 * 1024 * 1024 });
    }, 300_000);

    it("emits a dedicated validator CLI entry", () => {
        expect(existsSync(path.join(DIST, "cli.mjs"))).toBe(true);
    });

    it("cli.mjs --help prints exactly one Usage block — the validator's — and exits 0", async () => {
        const result = await run([path.join(DIST, "cli.mjs"), "--help"]);

        expect(result.stderr).toBe("");
        expect(result.code).toBe(0);
        expect(usageCount(result.stdout)).toBe(1);
        expect(result.stdout).toMatch(/Usage: validate/);
        expect(result.stdout).toContain("--strict");
        expect(result.stdout).toContain("--list-checks");
    });

    it("index.mjs --help starts no CLI at all", async () => {
        const result = await run([path.join(DIST, "index.mjs"), "--help"]);

        expect(result.code).toBe(0);
        expect(result.stdout).not.toContain("Usage:");
        expect(result.stderr).not.toContain("Usage:");
    });
});
