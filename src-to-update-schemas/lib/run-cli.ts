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
 * Shared CLI bootstrap. Replaces the Python `if __name__ == "__main__": main()`
 * idiom with a single helper.
 *
 * Each script does:
 *
 *     import { runCli } from "../../lib/run-cli";
 *     import { Command } from "commander";
 *
 *     export function buildCommand(): Command { ... }
 *     export async function main(opts: ...): Promise<number> { ... }
 *
 *     runCli(import.meta.url, async () => {
 *       const cmd = buildCommand().parse(process.argv);
 *       return main(cmd.opts(), cmd.args);
 *     });
 *
 * `runCli` only invokes `fn` when the file is the process entry point
 * (so importing it from another module — e.g. a test or another CLI —
 * does not re-trigger argv parsing).
 */

import { fileURLToPath } from "node:url";

import * as tmp from "tmp";

/**
 * Run `fn` if `metaUrl` (always pass `import.meta.url`) refers to the
 * process's entry point. The result of `fn` is treated as the desired
 * exit code; thrown errors are logged to stderr and exit 1.
 *
 * Use this at the bottom of every script — never call `process.exit`
 * directly inside `main`; just return a number.
 */
export const runCli = (metaUrl: string, fn: () => number | Promise<number>): void => {
    const [, entryArg] = process.argv;
    if (!entryArg) {
        return;
    }
    const modulePath = fileURLToPath(metaUrl);
    if (modulePath !== entryArg) {
        return;
    }
    Promise.resolve()
        .then(fn)
        .then((code) => {
            process.exit(typeof code === "number" ? code : 0);
        })
        .catch((err: unknown) => {
            const message = err instanceof Error ? (err.stack ?? err.message) : String(err);
            process.stderr.write(`${message}\n`);
            process.exit(1);
        });
};

/**
 * Translate a thrown Commander error (raised by `cmd.exitOverride()` on
 * missing args, invalid options, or `--help`) into the exit code Commander
 * intended. Falls back to `1` for non-Commander errors so callers can use
 * this from a generic `catch (err)` block.
 *
 * Commander already writes the user-facing message before the error throws,
 * so callers should NOT re-emit the message themselves.
 */
export const commanderExitCode = (err: unknown): number => {
    if (
        typeof err === "object" &&
        err !== null &&
        "exitCode" in err &&
        typeof (err as { readonly exitCode?: unknown }).exitCode === "number"
    ) {
        return (err as { readonly exitCode: number }).exitCode;
    }
    return 1;
};

/**
 * Run `fn` with a fresh temporary directory; the directory is recursively
 * removed afterwards (success or failure). Mirrors Python's
 * `with tempfile.TemporaryDirectory() as d:` block.
 *
 * Use this everywhere — do not call `tmp.dirSync` / `mkdtempSync` inline.
 */
export const withTempDir = async <T>(fn: (dir: string) => Promise<T> | T): Promise<T> => {
    const handle = tmp.dirSync({ unsafeCleanup: true });
    try {
        return await fn(handle.name);
    } finally {
        handle.removeCallback();
    }
};
