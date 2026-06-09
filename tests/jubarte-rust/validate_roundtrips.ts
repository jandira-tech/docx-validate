/**
 * Validate the 571 roundtrip DOCX files in `roundtrip_docx/` against the
 * docx-validate `lenient` and `word-valid` profiles. Mirrors the
 * harness style used by tests/probe-all-fixtures.ts in the upstream repo.
 *
 * Usage:
 *   bun run validate_roundtrips.ts
 *
 * Output:
 *   - validate_report.json  (per-file results + aggregate)
 *   - prints summary to stdout
 */

import * as fs from "node:fs/promises";
import * as path from "node:path";
import { validate } from "../../src/index";

const ROOT = path.dirname(new URL(import.meta.url).pathname);
const ROUNDTRIP_DIR = path.join(ROOT, "roundtrip_docx");
const REPORT_PATH = path.join(ROOT, "validate_report.json");

type StagedResult = {
    name: string;
    lenient_ok: boolean;
    lenient_error?: string;
    word_valid_ok: boolean;
    word_valid_error?: string;
};

async function listDocx(dir: string): Promise<string[]> {
    const entries = await fs.readdir(dir);
    return entries.filter((e) => e.endsWith(".docx")).sort();
}

function summarize(res: { valid?: boolean; issues?: { severity?: string; message?: string }[]; error?: unknown } | null | undefined): {
    ok: boolean;
    error?: string;
} {
    if (!res) return { ok: false, error: "no result" };
    if (res.valid === true) return { ok: true };
    const issues = res.issues ?? [];
    if (issues.length === 0 && typeof res.error === "string") {
        return { ok: false, error: res.error };
    }
    // Keep just the first 3 issue messages for the report; full detail is on disk.
    const top = issues
        .filter((i) => i.severity === "error" || i.severity === undefined)
        .slice(0, 3)
        .map((i) => i.message ?? "")
        .join(" | ");
    return { ok: false, error: top || `${issues.length} issues` };
}

async function main() {
    const files = await listDocx(ROUNDTRIP_DIR);
    console.log(`validating ${files.length} roundtrip docx files...`);

    const results: StagedResult[] = [];
    let lenient_ok = 0;
    let word_valid_ok = 0;

    let i = 0;
    for (const f of files) {
        i += 1;
        if (i % 25 === 0) {
            console.log(`  [${i}/${files.length}] ${f}`);
        }
        const target = path.join(ROUNDTRIP_DIR, f);
        let lenientResult: { ok?: boolean; error?: unknown; failingValidators?: unknown[] } | null = null;
        try {
            lenientResult = (await validate(target, { profile: "lenient" })) as never;
        } catch (e) {
            lenientResult = { ok: false, error: (e as Error).message };
        }
        const lenient = summarize(lenientResult);

        let wordResult: { ok?: boolean; error?: unknown; failingValidators?: unknown[] } | null = null;
        try {
            wordResult = (await validate(target, { profile: "word-valid" })) as never;
        } catch (e) {
            wordResult = { ok: false, error: (e as Error).message };
        }
        const word = summarize(wordResult);

        if (lenient.ok) lenient_ok += 1;
        if (word.ok) word_valid_ok += 1;

        results.push({
            name: f,
            lenient_ok: lenient.ok,
            lenient_error: lenient.error,
            word_valid_ok: word.ok,
            word_valid_error: word.error,
        });
    }

    const report = {
        total: files.length,
        lenient_ok,
        word_valid_ok,
        results,
    };
    await fs.writeFile(REPORT_PATH, JSON.stringify(report, null, 2));

    console.log("");
    console.log("== validate_roundtrips summary ==");
    console.log(`total:                     ${files.length}`);
    console.log(`lenient ok:                ${lenient_ok} (${((lenient_ok / files.length) * 100).toFixed(1)}%)`);
    console.log(`word-valid ok:             ${word_valid_ok} (${((word_valid_ok / files.length) * 100).toFixed(1)}%)`);
    console.log(`report:                    ${REPORT_PATH}`);
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
