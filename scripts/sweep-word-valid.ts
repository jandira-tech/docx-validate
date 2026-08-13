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

// Walk a dir of *.docx, run docx-validate word-valid profile on each.
// Usage: bun scripts/sweep-word-valid.ts <dir> [--json out.json]
import * as fs from "node:fs";
import * as path from "node:path";

import { runCli } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";

export function* walkDocxFiles(d: string): Generator<string> {
    let names: string[];
    try {
        names = fs.readdirSync(d);
    } catch {
        return;
    }
    for (const n of names) {
        const abs = path.join(d, n);
        let st: fs.Stats;
        try {
            st = fs.lstatSync(abs);
        } catch {
            continue;
        }
        if (st.isSymbolicLink()) {
            continue;
        }
        if (st.isDirectory()) {
            yield* walkDocxFiles(abs);
        } else if (n.toLowerCase().endsWith(".docx")) {
            yield abs;
        }
    }
}

export function formatSweepError(root: string, file: string, rest: string): string {
    return `${path.relative(root, file)}: ${rest}`;
}

export async function sweepWordValid(dir: string, jsonOut = ""): Promise<number> {
    const files = [...walkDocxFiles(dir)].sort();
    let ok = 0;
    let fail = 0;
    let err = 0;
    const errors: string[] = [];
    for (const f of files) {
        try {
            const r = await validate(f, { profile: "word-valid" });
            if (r.valid) {
                ok += 1;
            } else {
                fail += 1;
                const first = (r.issues || []).find((i) => (i.severity ?? "error") === "error");
                errors.push(formatSweepError(dir, f, `${first?.code ?? ""} ${first?.message?.slice(0, 90) ?? "?"}`));
            }
        } catch (e) {
            err += 1;
            const message = e instanceof Error ? e.message : String(e);
            errors.push(formatSweepError(dir, f, `THREW ${message.slice(0, 90)}`));
        }
    }
    process.stdout.write(`word-valid: ${ok}/${files.length} ok | ${fail} invalid | ${err} threw\n`);
    process.stdout.write("--- first 25 failures ---\n");
    for (const e of errors.slice(0, 25)) {
        process.stdout.write(`  ${e}\n`);
    }
    if (jsonOut) {
        fs.writeFileSync(jsonOut, JSON.stringify({ dir, total: files.length, ok, fail, err, errors }, null, 2));
    }
    return fail + err > 0 ? 1 : 0;
}

export async function main(argv: string[] = process.argv.slice(2)): Promise<number> {
    if (argv.length === 0 || !argv[0]) {
        process.stderr.write("Usage: bun scripts/sweep-word-valid.ts <dir> [--json out.json]\n");
        return 1;
    }
    const dir = path.resolve(argv[0]);
    if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) {
        process.stderr.write(`Error: not a directory: ${dir}\n`);
        return 1;
    }
    const jsonIdx = argv.indexOf("--json");
    let jsonOut = "";
    if (jsonIdx >= 0) {
        const outPath = argv[jsonIdx + 1];
        if (!outPath) {
            process.stderr.write("Error: --json requires an output file path\n");
            return 1;
        }
        jsonOut = path.resolve(outPath);
    }
    return sweepWordValid(dir, jsonOut);
}

runCli(import.meta.url, () => main());
