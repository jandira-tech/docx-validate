/*
 * Update fixtures-all.manifest.json by running the validator with all profiles
 * and combining with Word probe results.
 */

import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { validate } from "../src/scripts/office/validate.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_ROOT = path.resolve(HERE, "..", "tests/fixtures");
const PROBE_RESULTS = "/tmp/word-probe-results.jsonl";
const MANIFEST = path.resolve(HERE, "..", "tests", "fixtures-all.manifest.json");

interface ValidatorResult {
    valid: boolean;
    errorCodes: string[];
    threw?: string;
}

interface ProbeRecord {
    relativePath: string;
    file: string;
    validator: ValidatorResult & { profile: string };
    word: { outcome: string; clean: boolean; details: string; durationMs: number } | null;
    aligned: boolean | null;
    mismatch: string | null;
}

interface ManifestEntry {
    relativePath: string;
    strict: ValidatorResult;
    lenient: ValidatorResult;
    word: string;
}

function walkDocx(dir: string, out: string[]): void {
    for (const entry of readdirSync(dir)) {
        if (entry.startsWith("~$")) continue;
        const file = path.join(dir, entry);
        const stat = statSync(file);
        if (stat.isDirectory()) {
            walkDocx(file, out);
        } else if (entry.toLowerCase().endsWith(".docx")) {
            out.push(file);
        }
    }
}

async function runValidator(file: string, profile: "strict" | "lenient" | "word-valid"): Promise<ValidatorResult> {
    try {
        const result = await validate(file, { profile });
        const errorCodes = Array.from(
            new Set(
                result.issues
                    .filter((i) => i.severity === "error")
                    .map((i) => i.code)
                    .filter((c): c is string => Boolean(c)),
            ),
        ).sort();
        return { valid: result.valid, errorCodes };
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        return { valid: false, errorCodes: [], threw: msg.slice(0, 240) };
    }
}

function readProbeResults(): Map<string, ProbeRecord> {
    const map = new Map<string, ProbeRecord>();
    const content = readFileSync(PROBE_RESULTS, "utf-8");
    for (const line of content.split(/\r?\n/)) {
        if (!line.trim()) continue;
        try {
            const record = JSON.parse(line) as ProbeRecord;
            map.set(record.relativePath, record);
        } catch {
            // Skip malformed lines
        }
    }
    return map;
}

async function main(): Promise<void> {
    const files: string[] = [];
    walkDocx(FIXTURES_ROOT, files);
    files.sort();

    const probeResults = readProbeResults();
    const entries: ManifestEntry[] = [];

    for (const file of files) {
        const relativePath = path.relative(FIXTURES_ROOT, file);
        process.stderr.write(`Processing: ${relativePath}\n`);

        const strict = await runValidator(file, "strict");
        const lenient = await runValidator(file, "lenient");
        
        const probeRecord = probeResults.get(relativePath);
        const wordOutcome = probeRecord?.word?.outcome ?? "unknown";

        entries.push({
            relativePath,
            strict,
            lenient,
            word: wordOutcome,
        });
    }

    const manifest = {
        generatedAt: new Date().toISOString(),
        totalFixtures: entries.length,
        entries,
    };

    writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2) + "\n");
    process.stdout.write(`Updated manifest with ${entries.length} fixtures\n`);
}

main().catch((err) => {
    process.stderr.write(`Error: ${err}\n`);
    process.exit(1);
});
