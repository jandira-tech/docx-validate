/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Schema provenance record (src/scripts/office/schemas/SOURCES.md, issue
 * #304): every sha256 the record lists must match the bundled bytes, so a
 * regenerated or line-ending-normalised schema cannot ship behind a stale
 * record. Eight entries recorded the upstream CRLF form while the bundled
 * files are LF; the record now pins the bundled bytes.
 */

import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { defaultSchemasDir } from "../src/scripts/office/validators/base";

const SOURCES_MD = path.join(defaultSchemasDir(), "SOURCES.md");

/** Every provenance table row: file path relative to the schemas dir → sha256. */
const readRecords = async (): Promise<Record<string, string>> => {
    const text = await fs.readFile(SOURCES_MD, "utf8");
    const records: Record<string, string> = {};
    for (const line of text.split("\n")) {
        const match = line.match(/^\|\s*([^|]+?)\s*\|\s*([0-9a-f]{64})\s*\|$/);
        if (match) {
            records[match[1]!.trim()] = match[2]!;
        }
    }
    return records;
};

describe("schemas SOURCES.md provenance record", () => {
    it("lists schema files with a sha256 each, so the record is parseable", async () => {
        const records = await readRecords();
        // 22 iso-strict entries plus 3 dublincore entries; the threshold keeps
        // a markdown edit from silently turning the hash test vacuous.
        expect(Object.keys(records).length).toBeGreaterThanOrEqual(25);
        for (const name of Object.keys(records)) {
            expect(name, "records must stay relative paths under the schemas dir").toMatch(/^(?:iso-strict|dublincore)\/[\w.-]+\.xsd$/);
        }
    });

    it("records the sha256 of every listed file as bundled", async () => {
        const records = await readRecords();
        const mismatches: string[] = [];
        for (const [name, recorded] of Object.entries(records)) {
            const bytes = await fs.readFile(path.join(defaultSchemasDir(), name));
            const actual = createHash("sha256").update(bytes).digest("hex");
            if (actual !== recorded) {
                mismatches.push(`${name}: recorded ${recorded}, bundled ${actual}`);
            }
        }
        expect(mismatches).toEqual([]);
    });
});
