/*
 * Probe utility: walk every .docx / .pptx under tests/fixtures/, run the
 * validator in BOTH strict and lenient mode, and print a JSON manifest of
 * per-fixture outcomes. The manifest is consumed by the two
 * `tests/fixtures-all-*.test.ts` suites, which use it as a regression pin.
 *
 * Run with: bunx tsx scripts/probe-all-fixtures.ts > tests/fixtures-all.manifest.json
 */
import { readdirSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.resolve(HERE, "..", "tests/fixtures");

function walk(dir: string, out: string[]): void {
    for (const entry of readdirSync(dir)) {
        // Skip Office lock files (`~$*`) — Word creates these next to open
        // documents; they're 162-byte sentinels, not real OOXML archives,
        // and JSZip throws on them. Same filter as jubarte's
        // `discoverDocxArtifacts`.
        if (entry.startsWith("~$")) continue;
        const f = path.join(dir, entry);
        const s = statSync(f);
        if (s.isDirectory()) walk(f, out);
        else if (entry.toLowerCase().endsWith(".docx") || entry.toLowerCase().endsWith(".pptx")) out.push(f);
    }
}

interface Outcome {
    valid: boolean;
    /** Distinct error codes only (warnings ignored — see lenient downgrades). */
    errorCodes: string[];
    /** Set when the validator threw. Mutually exclusive with `valid`. */
    threw?: string;
}

interface Entry {
    relativePath: string;
    strict: Outcome;
    lenient: Outcome;
}

async function run(file: string, profile: "strict" | "lenient"): Promise<Outcome> {
    try {
        const r = await validate(file, { profile });
        const errorCodes = Array.from(
            new Set(
                r.issues
                    .filter((i) => i.severity === "error")
                    .map((i) => i.code)
                    .filter((c): c is string => Boolean(c)),
            ),
        ).sort();
        return { valid: r.valid, errorCodes };
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        return { valid: false, errorCodes: [], threw: msg.slice(0, 120) };
    }
}

(async () => {
    const all: string[] = [];
    walk(FIXTURES, all);
    all.sort();

    const entries: Entry[] = [];
    let i = 0;
    for (const f of all) {
        i += 1;
        const rel = path.relative(FIXTURES, f);
        process.stderr.write(`\r[${i}/${all.length}] ${rel}                            `);
        const [strict, lenient] = await Promise.all([run(f, "strict"), run(f, "lenient")]);
        entries.push({ relativePath: rel, strict, lenient });
    }
    process.stderr.write("\n");

    const out = {
        generatedAt: new Date().toISOString(),
        totalFixtures: entries.length,
        entries,
    };
    writeFileSync(path.resolve(HERE, "..", "tests/fixtures-all.manifest.json"), `${JSON.stringify(out, null, 2)}\n`);

    // Summary
    const strictPass = entries.filter((e) => e.strict.valid).length;
    const lenientPass = entries.filter((e) => e.lenient.valid).length;
    const strictOnlyFail = entries.filter((e) => e.lenient.valid && !e.strict.valid).length;
    const bothFail = entries.filter((e) => !e.lenient.valid && !e.strict.valid).length;
    const threw = entries.filter((e) => e.strict.threw || e.lenient.threw).length;

    console.log(`Total: ${entries.length}`);
    console.log(`Strict pass:           ${strictPass}`);
    console.log(`Lenient pass:          ${lenientPass}`);
    console.log(`Strict-only failures:  ${strictOnlyFail}`);
    console.log(`Both-mode failures:    ${bothFail}`);
    console.log(`Threw (any mode):      ${threw}`);
})();
