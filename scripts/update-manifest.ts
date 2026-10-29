/*
 * Update fixtures-all.manifest.json by running the validator with all profiles
 * and combining with Word probe results.
 */

import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { validate } from "../src/scripts/office/validate.js";
import { type ManifestWordField, mergeWordObservation, parseWordField, type WordHistoryEntry } from "../src/lib/word-provenance.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_ROOT = path.resolve(HERE, "..", "tests/fixtures");
const PROBE_RESULTS = path.resolve(HERE, "..", "tests", "word-probe-results.jsonl");
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
    word: ManifestWordField;
    history?: WordHistoryEntry[];
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
    if (!existsSync(PROBE_RESULTS)) return map;
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

/**
 * Resolve the `word` outcome for a fixture: a fresh probe result wins; else
 * the value carried in the previous manifest is preserved; else "unknown".
 * Keeps regen idempotent w.r.t. Word-probe metadata when the probe JSONL is
 * absent (no LibreOffice in CI).
 */
/**
 * The manifest key of a fixture: its path under the fixtures root, always with
 * forward slashes, the form the manifest and the probe results are keyed by.
 */
export function manifestKey(root: string, file: string, pathApi: typeof path = path): string {
    return pathApi.relative(root, file).split(pathApi.sep).join("/");
}

export function resolveWordOutcome(relativePath: string, probeOutcome: string | undefined, priorWord: Map<string, string>): string {
    return probeOutcome ?? priorWord.get(relativePath) ?? "unknown";
}

/**
 * Resolve the full `word` column (observation + history) for one manifest
 * entry: the hash-bound merge of #274. A prior observation is kept when its
 * recorded sha256 still matches the file, retired to `history` (under its old
 * hash) when the bytes changed, and superseded by a fresh probe outcome.
 */
export function resolveWordColumn(
    priorWord: unknown,
    priorHistory: readonly WordHistoryEntry[] | undefined,
    probeOutcome: string | undefined,
    currentSha256: string,
): { word: ManifestWordField; history: WordHistoryEntry[] } {
    const merged = mergeWordObservation({
        prior: parseWordField(priorWord),
        priorHistory,
        freshOutcome: probeOutcome,
        currentSha256,
    });
    return { word: merged.word, history: [...merged.history] };
}

interface PriorManifestEntry {
    relativePath?: string;
    strict?: ValidatorResult;
    lenient?: ValidatorResult;
    word?: unknown;
    history?: WordHistoryEntry[];
}

function readPriorManifestEntries(): Map<string, PriorManifestEntry> {
    const map = new Map<string, PriorManifestEntry>();
    if (!existsSync(MANIFEST)) return map;
    try {
        const prior = JSON.parse(readFileSync(MANIFEST, "utf-8")) as { entries?: PriorManifestEntry[] };
        for (const entry of prior.entries ?? []) {
            if (entry.relativePath) map.set(entry.relativePath, entry);
        }
    } catch {
        // Malformed prior manifest — treat as no prior data.
    }
    return map;
}

function sha256OfFile(file: string): string {
    return createHash("sha256").update(readFileSync(file)).digest("hex");
}

async function main(wordOnly: boolean): Promise<void> {
    const walked: string[] = [];
    walkDocx(FIXTURES_ROOT, walked);
    walked.sort();

    const probeResults = readProbeResults();
    const priorEntries = readPriorManifestEntries();

    // `--word-only` rebinds the Word column of the PRIOR manifest's entries
    // without re-running the validators, so the resulting diff is exactly the
    // sha256/structural migration. Fixtures on disk that the prior manifest
    // does not list are reported, not silently added.
    const files: string[] = [];
    if (wordOnly) {
        const walkedSet = new Set(walked.map((f) => manifestKey(FIXTURES_ROOT, f)));
        for (const relativePath of [...priorEntries.keys()].sort()) {
            if (!walkedSet.has(relativePath)) {
                throw new Error(`--word-only: manifest entry ${relativePath} has no fixture on disk`);
            }
            files.push(path.join(FIXTURES_ROOT, relativePath));
        }
        for (const relativePath of walkedSet) {
            if (!priorEntries.has(relativePath)) {
                process.stderr.write(`--word-only: skipping unlisted fixture ${relativePath}\n`);
            }
        }
    } else {
        files.push(...walked);
    }

    const entries: ManifestEntry[] = [];
    let migrated = 0;

    for (const file of files) {
        const relativePath = manifestKey(FIXTURES_ROOT, file);
        process.stderr.write(`Processing: ${relativePath}\n`);

        const prior = priorEntries.get(relativePath);
        if (wordOnly && (!prior || !prior.strict || !prior.lenient)) {
            throw new Error(`--word-only: no prior manifest entry with strict/lenient outcomes for ${relativePath}`);
        }

        const strict = wordOnly ? prior!.strict! : await runValidator(file, "strict");
        const lenient = wordOnly ? prior!.lenient! : await runValidator(file, "lenient");

        // The fixture's sha256 is computed once per run and binds the Word
        // observation to the exact bytes it was made against (#274).
        const sha256 = sha256OfFile(file);
        const probeRecord = probeResults.get(relativePath);
        const { word, history } = resolveWordColumn(prior?.word, prior?.history, probeRecord?.word?.outcome, sha256);
        if (prior && JSON.stringify(prior.word) !== JSON.stringify(word)) migrated += 1;

        entries.push(history.length > 0 ? { relativePath, strict, lenient, word, history } : { relativePath, strict, lenient, word });
    }

    const manifest = {
        generatedAt: new Date().toISOString(),
        totalFixtures: entries.length,
        entries,
    };

    writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2) + "\n");
    process.stdout.write(`Updated manifest with ${entries.length} fixtures (${migrated} word column(s) changed)\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
    const wordOnly = process.argv.includes("--word-only");
    main(wordOnly).catch((err) => {
        process.stderr.write(`Error: ${err}\n`);
        process.exit(1);
    });
}
