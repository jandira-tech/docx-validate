/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The Word outcome the fixtures manifest records for a fixture, through its
 * hash-bound provenance: an observation of other bytes reads as `unknown`.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { effectiveOutcome, parseWordField } from "../../src/lib/word-provenance";

const FIXTURES = path.join(import.meta.dirname, "..", "fixtures");
const MANIFEST = path.join(import.meta.dirname, "..", "fixtures-all.manifest.json");

export async function observedWordOutcome(relativePath: string): Promise<string> {
    const manifest = JSON.parse(await readFile(MANIFEST, "utf8")) as { entries: { relativePath: string; word?: unknown }[] };
    const entry = manifest.entries.find((candidate) => candidate.relativePath === relativePath);
    const sha256 = createHash("sha256")
        .update(await readFile(path.join(FIXTURES, relativePath)))
        .digest("hex");
    return effectiveOutcome(parseWordField(entry?.word), sha256);
}
