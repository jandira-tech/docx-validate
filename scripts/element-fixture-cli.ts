/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Shared CLI/manifest helpers for the element-fixture scripts
 * (generate-element-fixtures.ts, repair-element-fixtures.ts,
 * break-element-fixtures.ts).
 */

import { createHash } from "node:crypto";
import { existsSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import path from "node:path";

/**
 * Parse a --limit option value. `explicit` must be true when the flag was
 * actually passed on the command line (an omitted flag and an explicit
 * `--limit=` must not be conflated). Returns null when no limit was given;
 * otherwise a nonnegative safe integer. Empty or whitespace-only values are
 * rejected when the flag was passed explicitly.
 */
export const parseLimitOption = (raw: string, explicit: boolean): number | null => {
    if (!explicit) return null;
    const trimmed = raw.trim();
    if (!/^\d+$/.test(trimmed)) {
        throw new Error(`invalid --limit "${raw}": expected a nonnegative integer`);
    }
    const n = Number(trimmed);
    if (!Number.isSafeInteger(n)) {
        throw new Error(`invalid --limit "${raw}": expected a nonnegative integer`);
    }
    return n;
};

/** True when --limit was passed on the command line (either form). */
export const limitFlagPassed = (argv: string[]): boolean => argv.some((a) => a === "--limit" || a.startsWith("--limit="));

/** Deterministic code-unit ordering, independent of locale. */
export const compareCodeUnits = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);

/**
 * Resolve a user-supplied path and require it to stay inside `root`
 * (the repository). Rejects `..` escapes and absolute paths elsewhere so
 * CLI arguments cannot direct writes outside the repo.
 */
export const resolveInside = (root: string, candidate: string, label: string): string => {
    const resolvedRoot = path.resolve(root);
    const resolved = path.resolve(resolvedRoot, candidate);
    const rel = path.relative(resolvedRoot, resolved);
    if (rel === ".." || rel.startsWith(`..${path.sep}`) || path.isAbsolute(rel)) {
        throw new Error(`invalid ${label} "${candidate}": must be inside ${resolvedRoot}`);
    }
    return resolved;
};

export interface ManifestRow {
    element: string;
    file: string;
    sha256: string;
    profile: string;
    valid: boolean;
}

/**
 * Load a JSONL manifest into a map keyed by element, replacing duplicate rows.
 * Malformed lines (e.g. a torn final line from an interrupted write) are
 * skipped so a damaged manifest cannot break the next --resume.
 */
export const loadManifestRows = <T extends ManifestRow>(resultsPath: string): Map<string, T> => {
    const rows = new Map<string, T>();
    if (!existsSync(resultsPath)) return rows;
    for (const line of readFileSync(resultsPath, "utf8").split("\n")) {
        if (!line.trim()) continue;
        try {
            const row = JSON.parse(line) as T;
            rows.set(row.element, row);
        } catch {
            // skip malformed line
        }
    }
    return rows;
};

/**
 * Persist the manifest atomically: write to a temp file next to the target and
 * rename over it, so a crash mid-write can never truncate the previous
 * manifest (which would destroy rows for untouched fixtures).
 */
export const persistManifestRows = <T>(resultsPath: string, rows: Map<string, T>): void => {
    const tmp = `${resultsPath}.tmp-${process.pid}`;
    writeFileSync(tmp, [...rows.values()].map((r) => JSON.stringify(r) + "\n").join(""));
    renameSync(tmp, resultsPath);
};

/** True when the fixture on disk exists and still matches the row's recorded hash. */
export const fixtureMatchesRow = (outDir: string, row: { file: string; sha256: string }): boolean => {
    if (!row.sha256) return false;
    const fixturePath = path.join(outDir, row.file);
    return existsSync(fixturePath) && createHash("sha256").update(readFileSync(fixturePath)).digest("hex") === row.sha256;
};
