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
 * CLI contract tests for scripts/generate-element-fixtures.ts and
 * scripts/repair-element-fixtures.ts: --limit validation must happen before
 * any output is created or cleared, --only/--resume must preserve manifest
 * rows for unselected or unfinished fixtures, and a failed fixture write must
 * produce a nonzero exit status.
 */

import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { afterEach, describe, expect, it } from "vitest";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const run = promisify(execFile);

interface RunOutcome {
    code: number;
    stdout: string;
    stderr: string;
}

const tsx = async (script: string, args: string[]): Promise<RunOutcome> => {
    try {
        const { stdout, stderr } = await run("pnpm", ["exec", "tsx", script, ...args], { cwd: REPO, maxBuffer: 1 << 24 });
        return { code: 0, stdout, stderr };
    } catch (err) {
        const e = err as { code?: number; stdout?: string; stderr?: string };
        return { code: typeof e.code === "number" ? e.code : 1, stdout: e.stdout ?? "", stderr: e.stderr ?? "" };
    }
};

const GENERATE = "scripts/generate-element-fixtures.ts";
const REPAIR = "scripts/repair-element-fixtures.ts";

// The fixture scripts confine --out/--results to the repository, so tests
// must use a temp dir inside the repo (not os.tmpdir()).
const createdTmpDirs: string[] = [];
const tmp = (): string => {
    const dir = mkdtempSync(path.join(REPO, ".tmp-element-fixture-scripts-"));
    createdTmpDirs.push(dir);
    return dir;
};
afterEach(() => {
    while (createdTmpDirs.length > 0) rmSync(createdTmpDirs.pop()!, { recursive: true, force: true });
});

const readRows = (file: string): Record<string, unknown>[] =>
    readFileSync(file, "utf8")
        .split("\n")
        .filter((l) => l.trim())
        .map((l) => JSON.parse(l) as Record<string, unknown>);

const SENTINEL = JSON.stringify({ element: "other", file: "other.docx", sha256: "sentinel", profile: "lenient", valid: true });

describe("generate-element-fixtures --limit validation", () => {
    it("rejects a non-numeric --limit before creating or clearing outputs", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(GENERATE, ["--out", path.join(dir, "out"), "--results", results, "--limit", "abc"]);
        expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
        expect(readFileSync(results, "utf8")).toContain("sentinel");
    });

    it("rejects a fractional --limit", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        const out = await tsx(GENERATE, ["--out", path.join(dir, "out"), "--results", results, "--limit", "1.5"]);
        expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
        expect(existsSync(results)).toBe(false);
    });

    it("accepts --limit 0 as a full no-op that touches no outputs", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const outDir = path.join(dir, "out");
        const out = await tsx(GENERATE, ["--out", outDir, "--results", results, "--limit", "0"]);
        expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
        expect(readFileSync(results, "utf8")).toContain("sentinel");
        expect(existsSync(outDir)).toBe(false);
    });

    it("rejects an explicitly empty --limit= before touching outputs", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(GENERATE, ["--out", path.join(dir, "out"), "--results", results, "--limit="]);
        expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
        expect(readFileSync(results, "utf8")).toContain("sentinel");
    });
});

describe("repair-element-fixtures CLI contract", () => {
    it("rejects an invalid --limit before touching the manifest", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(REPAIR, ["--out", path.join(dir, "out"), "--results", results, "--only", "zoom", "--limit", "-1"]);
        expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
        expect(readFileSync(results, "utf8")).toContain("sentinel");
    });

    it("--only preserves manifest rows for unselected fixtures", { timeout: 180_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(REPAIR, ["--out", path.join(dir, "out"), "--results", results, "--only", "zoom"]);
        expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
        const rows = readRows(results);
        expect(rows.some((r) => r.element === "other" && r.sha256 === "sentinel")).toBe(true);
        const zoom = rows.find((r) => r.element === "zoom");
        expect(zoom, "expected a fresh row for zoom").toBeTruthy();
        expect(zoom?.valid).toBe(true);
        expect(String(zoom?.sha256)).not.toBe("");
    });

    it("--resume reprocesses build-failed rows and replaces them on success", { timeout: 180_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        const failed = JSON.stringify({
            element: "zoom",
            file: "zoom.docx",
            sha256: "",
            profile: "lenient",
            valid: false,
            issueCount: 0,
            errorCodes: ["build-failed"],
            warningCodes: [],
            errorMessages: ["boom"],
        });
        writeFileSync(results, failed + "\n");
        const out = await tsx(REPAIR, ["--out", path.join(dir, "out"), "--results", results, "--resume", "--only", "zoom"]);
        expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
        const rows = readRows(results).filter((r) => r.element === "zoom");
        expect(rows.length).toBe(1);
        expect(rows[0].valid).toBe(true);
        expect(String(rows[0].sha256)).not.toBe("");
    });

    it("--resume reprocesses a row whose fixture hash no longer matches", { timeout: 240_000 }, async () => {
        const dir = tmp();
        const outDir = path.join(dir, "out");
        const results = path.join(dir, "manifest.jsonl");
        const first = await tsx(REPAIR, ["--out", outDir, "--results", results, "--only", "zoom"]);
        expect(first.code, `stderr: ${first.stderr.slice(0, 300)}`).toBe(0);
        // Corrupt the fixture on disk: the recorded hash no longer matches, so
        // the row must be treated as incomplete and the fixture rebuilt.
        writeFileSync(path.join(outDir, "zoom.docx"), "tampered");
        const second = await tsx(REPAIR, ["--out", outDir, "--results", results, "--resume", "--only", "zoom"]);
        expect(second.code, `stderr: ${second.stderr.slice(0, 300)}`).toBe(0);
        const rows = readRows(results).filter((r) => r.element === "zoom");
        expect(rows.length).toBe(1);
        expect(rows[0].valid).toBe(true);
        const onDisk = readFileSync(path.join(outDir, "zoom.docx"));
        expect(onDisk.toString("utf8")).not.toBe("tampered");
        // The manifest row must describe the fixture that is actually on disk.
        expect(String(rows[0].sha256)).toBe(createHash("sha256").update(onDisk).digest("hex"));
    });

    it("rejects an explicitly empty --limit= before touching the manifest", { timeout: 120_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(REPAIR, ["--out", path.join(dir, "out"), "--results", results, "--only", "zoom", "--limit="]);
        expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
        expect(readFileSync(results, "utf8")).toContain("sentinel");
    });

    it("a limited run without --only/--resume preserves existing manifest rows", { timeout: 180_000 }, async () => {
        const dir = tmp();
        const results = path.join(dir, "manifest.jsonl");
        writeFileSync(results, SENTINEL + "\n");
        const out = await tsx(REPAIR, ["--out", path.join(dir, "out"), "--results", results, "--limit", "1"]);
        expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
        const rows = readRows(results);
        expect(rows.some((r) => r.element === "other" && r.sha256 === "sentinel")).toBe(true);
        expect(rows.length).toBe(2);
    });

    it("--resume does not trust generator rows (no repaired flag)", { timeout: 240_000 }, async () => {
        const dir = tmp();
        const outDir = path.join(dir, "out");
        const results = path.join(dir, "manifest.jsonl");
        const first = await tsx(REPAIR, ["--out", outDir, "--results", results, "--only", "zoom"]);
        expect(first.code, `stderr: ${first.stderr.slice(0, 300)}`).toBe(0);
        // Simulate a generator-written row: same shape, but no repaired flag.
        const rows = readRows(results).map((r) => {
            const { repaired: _repaired, ...rest } = r as Record<string, unknown> & { repaired?: boolean };
            return rest;
        });
        writeFileSync(results, rows.map((r) => JSON.stringify(r)).join("\n") + "\n");
        const second = await tsx(REPAIR, ["--out", outDir, "--results", results, "--resume", "--only", "zoom"]);
        expect(second.code, `stderr: ${second.stderr.slice(0, 300)}`).toBe(0);
        expect(second.stdout).toContain("repairing 1 of");
        const zoomRows = readRows(results).filter((r) => r.element === "zoom");
        expect(zoomRows.length).toBe(1);
        expect(zoomRows[0].repaired).toBe(true);
    });

    it("attachedTemplate gets a typed internal relationship, not a hyperlink", { timeout: 180_000 }, async () => {
        const dir = tmp();
        const outDir = path.join(dir, "out");
        const results = path.join(dir, "manifest.jsonl");
        const out = await tsx(REPAIR, ["--out", outDir, "--results", results, "--only", "attachedTemplate"]);
        expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
        const JSZip = (await import("jszip")).default;
        const zip = await JSZip.loadAsync(readFileSync(path.join(outDir, "attachedTemplate.docx")));
        const relsName = Object.keys(zip.files).find((n) => n.endsWith("settings.xml.rels"));
        expect(relsName, "expected a settings sidecar").toBeTruthy();
        const rels = await zip.files[relsName!].async("string");
        expect(rels).toContain("/attachedTemplate");
        expect(rels).not.toContain('TargetMode="External"');
    });

    it("returns nonzero when a selected fixture cannot be written", { timeout: 180_000 }, async () => {
        const dir = tmp();
        const outDir = path.join(dir, "out");
        mkdirSync(path.join(outDir, "zoom.docx"), { recursive: true }); // writeFileSync will fail with EISDIR
        const results = path.join(dir, "manifest.jsonl");
        const out = await tsx(REPAIR, ["--out", outDir, "--results", results, "--only", "zoom"]);
        expect(out.code, "expected nonzero exit when the fixture write fails").not.toBe(0);
        const rows = readRows(results);
        expect(rows.some((r) => r.element === "zoom" && r.errorCodes && (r.errorCodes as string[]).includes("build-failed"))).toBe(true);
    });
});
