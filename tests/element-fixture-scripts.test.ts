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
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { afterEach, describe, expect, it } from "vitest";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const run = promisify(execFile);

const GENERATE = "scripts/generate-element-fixtures.ts";
const REPAIR = "scripts/repair-element-fixtures.ts";
const BREAK = "scripts/break-element-fixtures.ts";
const VALID_CORPUS = path.join(REPO, "word-element-fixtures");

interface RunOutcome {
    code: number;
    stdout: string;
    stderr: string;
}

interface Ctx {
    dir: string;
    outDir: string;
    results: string;
}

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

/** Fresh temp dir with a manifest seeded with `seed` (default: a sentinel row). */
const setup = (seed = SENTINEL): Ctx => {
    const dir = tmp();
    const results = path.join(dir, "manifest.jsonl");
    writeFileSync(results, seed + "\n");
    return { dir, outDir: path.join(dir, "out"), results };
};

const tsx = async (script: string, args: string[]): Promise<RunOutcome> => {
    try {
        const { stdout, stderr } = await run("pnpm", ["exec", "tsx", script, ...args], { cwd: REPO, maxBuffer: 1 << 24 });
        return { code: 0, stdout, stderr };
    } catch (err) {
        const e = err as { code?: number; stdout?: string; stderr?: string };
        return { code: typeof e.code === "number" ? e.code : 1, stdout: e.stdout ?? "", stderr: e.stderr ?? "" };
    }
};

/** Run a fixture script with the standard --out/--results wiring plus extras. */
const runScript = (script: string, ctx: Ctx, extra: string[]): Promise<RunOutcome> =>
    tsx(script, ["--out", ctx.outDir, "--results", ctx.results, ...extra]);

const readRows = (file: string): Record<string, unknown>[] =>
    readFileSync(file, "utf8")
        .split("\n")
        .filter((l) => l.trim())
        .map((l) => JSON.parse(l) as Record<string, unknown>);

const SENTINEL = JSON.stringify({ element: "other", file: "other.docx", sha256: "sentinel", profile: "lenient", valid: true });

const expectNonzero = (out: RunOutcome): void => expect(out.code, `expected nonzero exit, stderr: ${out.stderr.slice(0, 300)}`).not.toBe(0);
const expectOk = (out: RunOutcome): void => expect(out.code, `stderr: ${out.stderr.slice(0, 300)}`).toBe(0);
const sentinelKept = (ctx: Ctx): void => expect(readFileSync(ctx.results, "utf8")).toContain("sentinel");

describe("generate-element-fixtures --limit validation", () => {
    it("rejects a non-numeric --limit before creating or clearing outputs", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectNonzero(await runScript(GENERATE, ctx, ["--limit", "abc"]));
        sentinelKept(ctx);
    });

    it("rejects a fractional --limit", { timeout: 120_000 }, async () => {
        const ctx = setup(""); // no manifest file yet
        rmSync(ctx.results);
        expectNonzero(await runScript(GENERATE, ctx, ["--limit", "1.5"]));
        expect(existsSync(ctx.results)).toBe(false);
    });

    it("accepts --limit 0 as a full no-op that touches no outputs", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectOk(await runScript(GENERATE, ctx, ["--limit", "0"]));
        sentinelKept(ctx);
        expect(existsSync(ctx.outDir)).toBe(false);
    });

    it("rejects an explicitly empty --limit= before touching outputs", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectNonzero(await runScript(GENERATE, ctx, ["--limit="]));
        sentinelKept(ctx);
    });
});

describe("repair-element-fixtures CLI contract", () => {
    it("rejects an invalid --limit before touching the manifest", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectNonzero(await runScript(REPAIR, ctx, ["--only", "zoom", "--limit", "-1"]));
        sentinelKept(ctx);
    });

    it("rejects an explicitly empty --limit= before touching the manifest", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectNonzero(await runScript(REPAIR, ctx, ["--only", "zoom", "--limit="]));
        sentinelKept(ctx);
    });

    it("--only preserves manifest rows for unselected fixtures", { timeout: 180_000 }, async () => {
        const ctx = setup();
        expectOk(await runScript(REPAIR, ctx, ["--only", "zoom"]));
        const rows = readRows(ctx.results);
        expect(rows.some((r) => r.element === "other" && r.sha256 === "sentinel")).toBe(true);
        const zoom = rows.find((r) => r.element === "zoom");
        expect(zoom, "expected a fresh row for zoom").toBeTruthy();
        expect(zoom?.valid).toBe(true);
        expect(String(zoom?.sha256)).not.toBe("");
    });

    it("a limited run without --only/--resume preserves existing manifest rows", { timeout: 180_000 }, async () => {
        const ctx = setup();
        expectOk(await runScript(REPAIR, ctx, ["--limit", "1"]));
        const rows = readRows(ctx.results);
        expect(rows.some((r) => r.element === "other" && r.sha256 === "sentinel")).toBe(true);
        expect(rows.length).toBe(2);
    });

    it("--resume reprocesses build-failed rows and replaces them on success", { timeout: 180_000 }, async () => {
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
        const ctx = setup(failed);
        expectOk(await runScript(REPAIR, ctx, ["--resume", "--only", "zoom"]));
        const rows = readRows(ctx.results).filter((r) => r.element === "zoom");
        expect(rows.length).toBe(1);
        expect(rows[0].valid).toBe(true);
        expect(String(rows[0].sha256)).not.toBe("");
    });

    it("--resume reprocesses a row whose fixture hash no longer matches", { timeout: 240_000 }, async () => {
        const ctx = setup("");
        rmSync(ctx.results);
        expectOk(await runScript(REPAIR, ctx, ["--only", "zoom"]));
        // Corrupt the fixture on disk: the recorded hash no longer matches, so
        // the row must be treated as incomplete and the fixture rebuilt.
        writeFileSync(path.join(ctx.outDir, "zoom.docx"), "tampered");
        expectOk(await runScript(REPAIR, ctx, ["--resume", "--only", "zoom"]));
        const rows = readRows(ctx.results).filter((r) => r.element === "zoom");
        expect(rows.length).toBe(1);
        expect(rows[0].valid).toBe(true);
        const onDisk = readFileSync(path.join(ctx.outDir, "zoom.docx"));
        expect(onDisk.toString("utf8")).not.toBe("tampered");
        // The manifest row must describe the fixture that is actually on disk.
        expect(String(rows[0].sha256)).toBe(createHash("sha256").update(onDisk).digest("hex"));
    });

    it("--resume does not trust generator rows (no repaired flag)", { timeout: 240_000 }, async () => {
        const ctx = setup("");
        rmSync(ctx.results);
        expectOk(await runScript(REPAIR, ctx, ["--only", "zoom"]));
        // Simulate a generator-written row: same shape, but no repaired flag.
        const rows = readRows(ctx.results).map((r) => {
            const { repaired: _repaired, ...rest } = r as Record<string, unknown> & { repaired?: boolean };
            return rest;
        });
        writeFileSync(ctx.results, rows.map((r) => JSON.stringify(r)).join("\n") + "\n");
        const second = await runScript(REPAIR, ctx, ["--resume", "--only", "zoom"]);
        expectOk(second);
        expect(second.stdout).toContain("repairing 1 of");
        const zoomRows = readRows(ctx.results).filter((r) => r.element === "zoom");
        expect(zoomRows.length).toBe(1);
        expect(zoomRows[0].repaired).toBe(true);
    });

    it("attachedTemplate gets a typed internal relationship, not a hyperlink", { timeout: 180_000 }, async () => {
        const ctx = setup("");
        rmSync(ctx.results);
        expectOk(await runScript(REPAIR, ctx, ["--only", "attachedTemplate"]));
        const JSZip = (await import("jszip")).default;
        const zip = await JSZip.loadAsync(readFileSync(path.join(ctx.outDir, "attachedTemplate.docx")));
        const relsName = Object.keys(zip.files).find((n) => n.endsWith("settings.xml.rels"));
        expect(relsName, "expected a settings sidecar").toBeTruthy();
        const rels = await zip.files[relsName!].async("string");
        expect(rels).toContain("/attachedTemplate");
        expect(rels).not.toContain('TargetMode="External"');
    });

    it("returns nonzero when a selected fixture cannot be written", { timeout: 180_000 }, async () => {
        const ctx = setup("");
        rmSync(ctx.results);
        mkdirSync(path.join(ctx.outDir, "zoom.docx"), { recursive: true }); // writeFileSync will fail with EISDIR
        const out = await runScript(REPAIR, ctx, ["--only", "zoom"]);
        expectNonzero(out);
        const rows = readRows(ctx.results);
        expect(rows.some((r) => r.element === "zoom" && r.errorCodes && (r.errorCodes as string[]).includes("build-failed"))).toBe(true);
    });
});

describe("break-element-fixtures CLI contract", () => {
    it("rejects an invalid --limit before touching the manifest", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectNonzero(await runScript(BREAK, ctx, ["--src", VALID_CORPUS, "--limit", "-1"]));
        sentinelKept(ctx);
    });

    it("accepts --limit 0 as a full no-op that touches no outputs", { timeout: 120_000 }, async () => {
        const ctx = setup();
        expectOk(await runScript(BREAK, ctx, ["--src", VALID_CORPUS, "--limit", "0"]));
        sentinelKept(ctx);
    });

    it("--only preserves manifest rows for unselected fixtures and breaks the selected one", { timeout: 180_000 }, async () => {
        const ctx = setup();
        expectOk(await runScript(BREAK, ctx, ["--src", VALID_CORPUS, "--only", "zoom"]));
        const rows = readRows(ctx.results);
        expect(rows.some((r) => r.element === "other" && r.sha256 === "sentinel")).toBe(true);
        const zoom = rows.find((r) => r.element === "zoom");
        expect(zoom, "expected a fresh row for zoom").toBeTruthy();
        expect(zoom?.valid).toBe(false);
        expect(String(zoom?.sha256)).not.toBe("");
    });
});
