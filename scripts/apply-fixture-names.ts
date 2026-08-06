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
 * Fingerprint a set of .docx files, derive content-descriptive names, and
 * `git mv` them — either in place (rename only) or sorted into
 * tests/fixtures/<category>/. Dedups by content hash. Supports --dry-run.
 *
 *   bunx tsx scripts/apply-fixture-names.ts [--into-categories] [--dedup] \
 *     [--dry-run] [--fixtures-root tests/fixtures] \
 *     [--descriptor error-first|content-first] [--category <dir>] <path>...
 *
 * --descriptor content-first names files by their distinguishing content
 * feature, falling back to the error code (for real-world specimens that share
 * a benign validation quirk). --category forces a single target dir under
 * --fixtures-root instead of the broken/working split.
 */

import { execFileSync } from "node:child_process";
import path from "node:path";
import { mkdirSync, readdirSync, statSync } from "node:fs";

import { type DescriptorMode, type FixtureFingerprint, deriveName } from "./derive-fixture-name";
import { fingerprint } from "./fixture-fingerprint";

export interface FingerprintedFile {
    sourcePath: string;
    fingerprint: FixtureFingerprint;
}

export interface PlanOptions {
    intoCategories: boolean;
    fixturesRoot: string;
    dedup?: boolean;
    /** Passed through to deriveName; defaults to "error-first". */
    descriptorMode?: DescriptorMode;
    /** Force a single category dir (e.g. "eigen") instead of broken/working. */
    fixedCategory?: string;
}

export interface Move {
    from: string;
    to: string;
}

/** Pure: turn fingerprinted files into a deduped, collision-free move list. */
export function planMoves(files: FingerprintedFile[], opts: PlanOptions): Move[] {
    const seenHash = new Set<string>();
    const takenTargets = new Set<string>();
    const moves: Move[] = [];

    // Deterministic order; dedup keeps the lexicographically-first source.
    const sorted = [...files].sort((a, b) => a.sourcePath.localeCompare(b.sourcePath));

    for (const file of sorted) {
        if (opts.dedup) {
            if (seenHash.has(file.fingerprint.contentHash)) continue;
            seenHash.add(file.fingerprint.contentHash);
        }
        const derived = deriveName(file.fingerprint, {
            descriptorMode: opts.descriptorMode,
            fixedCategory: opts.fixedCategory,
        });
        const dir = opts.intoCategories ? path.join(opts.fixturesRoot, derived.category) : path.dirname(file.sourcePath);

        let candidate = path.join(dir, derived.fileName);
        let n = 2;
        while (takenTargets.has(candidate)) {
            candidate = path.join(dir, `${derived.subjectSlug}.${derived.descriptor}-${n}.docx`);
            n += 1;
        }
        takenTargets.add(candidate);
        moves.push({ from: file.sourcePath, to: candidate });
    }
    return moves;
}

function collectDocx(target: string, out: string[]): void {
    const stat = statSync(target);
    if (stat.isDirectory()) {
        for (const entry of readdirSync(target)) {
            if (entry.startsWith("~$")) continue;
            collectDocx(path.join(target, entry), out);
        }
    } else if (target.toLowerCase().endsWith(".docx")) {
        out.push(target);
    }
}

async function main(): Promise<void> {
    const argv = process.argv.slice(2);
    const intoCategories = argv.includes("--into-categories");
    const dedup = argv.includes("--dedup");
    const dryRun = argv.includes("--dry-run");

    // Flags that take a value; their value-index must be excluded from targets.
    const valueFlagIndices = new Set<number>();
    const flagValue = (name: string): string | undefined => {
        const i = argv.indexOf(name);
        if (i < 0) return undefined;
        valueFlagIndices.add(i + 1);
        return argv[i + 1];
    };
    const fixturesRoot = flagValue("--fixtures-root") ?? "tests/fixtures";
    const descriptorMode = flagValue("--descriptor") as DescriptorMode | undefined;
    const fixedCategory = flagValue("--category");

    const targets = argv.filter((a, i) => !a.startsWith("--") && !valueFlagIndices.has(i));

    const docxPaths: string[] = [];
    for (const t of targets) collectDocx(t, docxPaths);

    const files: FingerprintedFile[] = [];
    for (const p of docxPaths) {
        files.push({ sourcePath: p, fingerprint: await fingerprint(p) });
    }

    const moves = planMoves(files, {
        intoCategories,
        fixturesRoot,
        dedup,
        descriptorMode,
        fixedCategory,
    });
    const dropped = files.length - moves.length;

    for (const m of moves) {
        process.stdout.write(`${m.from}  ->  ${m.to}\n`);
        if (!dryRun) {
            // git mv won't create a missing target directory (e.g. a new category dir).
            mkdirSync(path.dirname(m.to), { recursive: true });
            execFileSync("git", ["mv", m.from, m.to]);
        }
    }
    process.stdout.write(`\n${moves.length} renamed, ${dropped} dropped as duplicates${dryRun ? " (dry-run)" : ""}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
    main().catch((err) => {
        process.stderr.write(`Error: ${err}\n`);
        process.exit(1);
    });
}
