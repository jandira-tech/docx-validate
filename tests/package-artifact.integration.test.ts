/*
 * Package-artifact smoke test (#274, Task 13 part C).
 *
 * Gated on PACKAGE_SMOKE=1 — skipped otherwise (the name says so). Building,
 * packing and installing into a fresh project is far too slow for the regular
 * suite, so it runs only when explicitly requested:
 *
 *   PACKAGE_SMOKE=1 pnpm exec vitest run tests/package-artifact.integration.test.ts
 *
 * It must:
 *   1. build (`pnpm run build`) and `pnpm pack` into a temp dir;
 *   2. install the tarball into a fresh temp project with no access to this
 *      repo's sources;
 *   3. in that project, import the public entrypoints, validate one small
 *      fixture (copied in), and assert the bundled schemas resolve — both the
 *      Transitional tree and the XSD root strict uses;
 *   4. run the packaged CLI's --help (exit 0).
 */

import { execFile } from "node:child_process";
import { copyFileSync, mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";

const execFileAsync = promisify(execFile);

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const FIXTURE = path.join(REPO, "tests/fixtures/working/sample-document.afterword-repaired-word-repaired.docx");

const run = (command: string, args: readonly string[], options: { cwd: string; timeout?: number }) =>
    execFileAsync(command, args, { cwd: options.cwd, timeout: options.timeout ?? 120_000, maxBuffer: 16 * 1024 * 1024 });

/**
 * Executed inside the consumer project with `node`. Imports the installed
 * package only — nothing from this repo — and exits non-zero on any failed
 * expectation, printing a JSON summary for the test to pin.
 */
const CONSUMER_DRIVER = `
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { defaultSchemasDir, jubarte, Validate, validate } from "docx-validate";

const assert = (condition, message) => {
    if (!condition) {
        throw new Error(message);
    }
};

assert(typeof validate === "function", "public entrypoint validate is not a function");
assert(typeof Validate === "function", "public class Validate is missing");
assert(typeof jubarte.validate === "function", "jubarte.validate helper is missing");

const fixture = path.resolve("fixture.docx");

const lenient = await validate(fixture, { profile: "lenient" });
assert(lenient.valid === true, "lenient validation of the strict-clean fixture failed: " + JSON.stringify(lenient.issues));

const strict = await validate(fixture, { profile: "strict" });
assert(strict.valid === true, "strict validation of the strict-clean fixture failed: " + JSON.stringify(strict.issues));
assert(strict.xsdWalk?.started === true, "strict profile never started the XSD walk");
assert(
    !strict.issues.some((issue) => issue.code === "xsd-schema-load-failed"),
    "bundled schemas failed to load from the installed package",
);
assert(strict.xsdWalk.eligible > 0, "no parts were eligible for XSD validation");

const viaBytes = await jubarte.validate(new Uint8Array(readFileSync(fixture)));
assert(viaBytes.valid === true, "jubarte.validate(bytes) failed on the same fixture");

const schemasDir = defaultSchemasDir();
assert(schemasDir.includes("node_modules"), "defaultSchemasDir() does not resolve inside the installed package: " + schemasDir);
assert(
    existsSync(path.join(schemasDir, "ISO-IEC29500-4_2016", "wml.xsd")),
    "Transitional wml.xsd is missing from the bundled schemas",
);
assert(existsSync(path.join(schemasDir, "mce", "mc.xsd")), "mce/mc.xsd (strict import closure) is missing from the bundled schemas");

console.log(JSON.stringify({ ok: true, eligible: strict.xsdWalk.eligible, schemasDir }));
`;

const gated = process.env.PACKAGE_SMOKE === "1" ? describe : describe.skip;

gated("package artifact smoke (skipped unless PACKAGE_SMOKE=1)", { timeout: 600_000 }, () => {
    it("builds, packs, installs into a fresh project, validates with bundled schemas, and the CLI --help exits 0", async () => {
        await withTempDir(async (tmp) => {
            // 1. Build + pack into the temp dir.
            await run("pnpm", ["run", "build"], { cwd: REPO, timeout: 240_000 });
            const pack = await run("pnpm", ["pack", "--pack-destination", tmp], { cwd: REPO });
            // `pnpm pack` prints the tarball path under "Tarball Details" as
            // an absolute path already.
            const tarballLine = pack.stdout
                .trim()
                .split(/\r?\n/)
                .filter((line) => line.endsWith(".tgz"))
                .at(-1);
            expect(tarballLine, `pnpm pack did not report a tarball: ${pack.stdout}`).toBeTruthy();
            const tarball = path.isAbsolute(tarballLine!) ? tarballLine! : path.join(tmp, tarballLine!);

            // 2. Fresh consumer project — a directory under the system
            //    temp dir, outside this repo, so module resolution cannot
            //    reach the repo's sources.
            const consumer = path.join(tmp, "consumer");
            mkdirSync(consumer, { recursive: true });
            writeFileSync(
                path.join(consumer, "package.json"),
                `${JSON.stringify({ name: "docx-validate-smoke-consumer", private: true, type: "module" }, null, 2)}\n`,
            );
            await run("pnpm", ["add", tarball], { cwd: consumer, timeout: 300_000 });

            // 3. Copy the fixture in and run the driver against the
            //    installed package only.
            copyFileSync(FIXTURE, path.join(consumer, "fixture.docx"));
            writeFileSync(path.join(consumer, "smoke.mjs"), CONSUMER_DRIVER);
            const driver = await run(process.execPath, ["smoke.mjs"], { cwd: consumer, timeout: 180_000 });
            const summary = JSON.parse(driver.stdout.trim().split(/\r?\n/).at(-1)!) as {
                ok: boolean;
                eligible: number;
                schemasDir: string;
            };
            expect(summary.ok).toBe(true);
            expect(summary.eligible).toBeGreaterThan(0);
            expect(summary.schemasDir).toContain(path.join("node_modules", "docx-validate"));

            // 4. The packaged CLI's --help exits 0 (execFile rejects on any
            //    non-zero exit) and prints usage text.
            const cli = path.join(consumer, "node_modules", "docx-validate", "dist", "index.mjs");
            const help = await run(process.execPath, [cli, "--help"], { cwd: consumer, timeout: 60_000 });
            expect(help.stdout).toContain("Usage:");
        });
    });
});
