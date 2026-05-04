#!/usr/bin/env bun
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
 * Release wrapper. Reads `.release-checklist.md`, prompts the user to
 * confirm each unchecked item (`- [ ] …`), then forwards to `bumpp` with
 * any extra args.
 *
 * Usage:
 *   bun run release                  # interactive prompts + interactive bumpp
 *   bun run release 0.1.2            # interactive prompts + bumpp 0.1.2
 *   bun run release patch            # interactive prompts + bumpp patch
 *   bun run release --yes 0.1.2      # skip prompts (CI), bumpp 0.1.2
 *
 * Aborts (exit 1) on the first "no" answer. Unrecognised answer = "no".
 *
 * The checklist isn't hard-coded — it lives in `.release-checklist.md`
 * at the repo root. Add/remove items there.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import * as readline from "node:readline/promises";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const CHECKLIST = path.resolve(REPO, ".release-checklist.md");

interface Options {
    skipPrompts: boolean;
    bumppArgs: string[];
}

function parseArgs(argv: readonly string[]): Options {
    const skipPrompts = argv.includes("--yes") || argv.includes("-y");
    const bumppArgs = argv.filter((a) => a !== "--yes" && a !== "-y");
    return { skipPrompts, bumppArgs };
}

async function loadChecklist(): Promise<string[]> {
    let text: string;
    try {
        text = await fs.readFile(CHECKLIST, "utf8");
    } catch (err) {
        if ((err as NodeJS.ErrnoException).code === "ENOENT") return [];
        throw err;
    }
    // Match top-level "- [ ] …" lines only. "- [x]" lines (already-checked)
    // are skipped — those are items the user has marked as "always green".
    const items: string[] = [];
    for (const line of text.split("\n")) {
        const m = line.match(/^- \[ \] (.+)$/);
        if (m?.[1]) items.push(m[1].trim());
    }
    return items;
}

async function confirmEach(items: readonly string[]): Promise<void> {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    try {
        for (const item of items) {
            const answer = (await rl.question(`  [?] ${item}\n      (y/N) `)).trim().toLowerCase();
            if (answer !== "y" && answer !== "yes") {
                process.stderr.write(
                    `\n  ✗ Not confirmed: "${item}"\n  Aborting release. Edit .release-checklist.md or fix the item, then re-run.\n`,
                );
                process.exit(1);
            }
            process.stdout.write("      ✓\n");
        }
    } finally {
        rl.close();
    }
}

function runBumpp(args: readonly string[]): Promise<number> {
    return new Promise((resolve, reject) => {
        const child = spawn("bunx", ["bumpp", ...args], { stdio: "inherit", cwd: REPO });
        child.on("error", reject);
        child.on("close", (code) => resolve(code ?? 0));
    });
}

async function main(): Promise<void> {
    const opts = parseArgs(process.argv.slice(2));
    const items = await loadChecklist();

    if (items.length === 0) {
        process.stdout.write("(no unchecked items in .release-checklist.md — skipping prompts)\n\n");
    } else if (opts.skipPrompts) {
        process.stdout.write(`Release checklist (${items.length} items, --yes given so all auto-confirmed):\n`);
        for (const item of items) process.stdout.write(`  ✓ ${item}\n`);
        process.stdout.write("\n");
    } else {
        process.stdout.write(
            `\nRelease checklist — confirm each item before bumpp runs.\n${items.length} items from .release-checklist.md:\n\n`,
        );
        await confirmEach(items);
        process.stdout.write("\n  All items confirmed. Handing off to bumpp.\n\n");
    }

    const code = await runBumpp(opts.bumppArgs);
    process.exit(code);
}

main().catch((err: unknown) => {
    process.stderr.write(`release: ${err instanceof Error ? err.message : String(err)}\n`);
    process.exit(1);
});
