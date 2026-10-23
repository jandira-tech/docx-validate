/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Plan Task 13: manifest writers never discard existing Word evidence. Every
 * script that writes tests/fixtures-all.manifest.json must merge each entry's
 * Word observation through `mergeWordObservation`, and the refresh script
 * must run such a writer.
 */
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SCRIPTS = path.join(REPO, "scripts");

const manifestWriters = (): string[] =>
    readdirSync(SCRIPTS)
        .filter((name) => name.endsWith(".ts"))
        .filter((name) => {
            const source = readFileSync(path.join(SCRIPTS, name), "utf-8");
            return source.includes("fixtures-all.manifest.json") && /\bwriteFile(Sync)?\(/.test(source);
        });

describe("fixtures manifest writers", () => {
    it("merge every Word observation instead of rewriting the entry", () => {
        const writers = manifestWriters();
        expect(writers.length).toBeGreaterThan(0);
        for (const writer of writers) {
            expect(readFileSync(path.join(SCRIPTS, writer), "utf-8"), writer).toMatch(/\bmergeWordObservation\(/);
        }
    });

    it("refresh through the provenance-preserving writer", () => {
        const pkg = JSON.parse(readFileSync(path.join(REPO, "package.json"), "utf-8")) as { scripts: Record<string, string> };
        const refresh = pkg.scripts["test:fixtures:refresh"] ?? "";
        const script = /scripts\/([\w-]+\.ts)/.exec(refresh)?.[1];
        expect(script, refresh).toBeDefined();
        expect(manifestWriters()).toContain(script);
        expect(readFileSync(path.join(SCRIPTS, script!), "utf-8"), script).toMatch(/\bmergeWordObservation\(/);
    });
});
