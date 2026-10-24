/// <reference types="vitest" />
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { defineConfig } from "vite-plus";

const here = path.dirname(fileURLToPath(import.meta.url));
const SCHEMAS_SRC = path.resolve(here, "src/scripts/office/schemas");
const SCHEMAS_DST = path.resolve(here, "dist/schemas");

// vite-plus's pack pipeline collapses TypeScript to dist/index.mjs and does
// not copy non-TS assets. The bundled OOXML XSDs at src/scripts/office/schemas/
// must ship inside dist/ so defaultSchemasDir() in validators/base.ts resolves
// against the published bundle (not the source tree, which is absent on npm).
const copySchemasPlugin = {
    name: "docx-validate:copy-schemas",
    async closeBundle() {
        await fs.rm(SCHEMAS_DST, { recursive: true, force: true });
        await fs.cp(SCHEMAS_SRC, SCHEMAS_DST, { recursive: true });
    },
};

const CLI_ENTRY = path.resolve(here, "src/cli.ts");
const SRC_DIR = path.resolve(here, "src") + path.sep;
// A file: URL that can never canonicalize to a real argv[1] — the bundle-only
// replacement for `import.meta.url` inside the `runCli` shims.
const INERT_CLI_META_URL = '"file:///docx-validate.cli-shim-inert"';

// Every src/scripts/** CLI module ends with `runCli(import.meta.url, …)` so it
// doubles as a `tsx` entry point. Bundled into dist/index.mjs, those modules
// all share ONE import.meta.url — the guard can no longer tell
// `node dist/index.mjs` apart from a library import, so every shim fires at
// once (#314: `node dist/index.mjs --help` printed several usages). Rewrite
// the shim's metaUrl to an inert file: URL so only src/cli.ts (exempt below —
// it IS the packaged validator CLI) keeps a live entry-point guard.
const inertCliShimsPlugin = {
    name: "docx-validate:inert-cli-shims",
    transform(code: string, id: string) {
        const file = id.split("?")[0]!;
        if (file === CLI_ENTRY || !file.startsWith(SRC_DIR) || !file.endsWith(".ts")) return null;
        if (!code.includes("runCli(import.meta.url")) return null;
        return {
            code: code.replaceAll("runCli(import.meta.url", `runCli(${INERT_CLI_META_URL}`),
            map: null,
        };
    },
};

export default defineConfig({
    // Format settings live here — NOT in `.oxfmtrc.json`. vite-plus's oxfmt
    // invocation (vp fmt / vp check) only reads the `fmt` block and ignores
    // `.oxfmtrc.json`, so a standalone rc file silently desyncs `vp check`
    // from `pnpm run fmt` (they formatted with different settings). These
    // values are the historical `.oxfmtrc.json` contents, verbatim.
    fmt: {
        trailingComma: "all",
        printWidth: 140,
        tabWidth: 4,
        arrowParens: "always",
        bracketSpacing: true,
        proseWrap: "preserve",
        sortPackageJson: false,
    },
    pack: {
        // Two entry points: the library bundle (dist/index.mjs) and the
        // packaged validator CLI (dist/cli.mjs, package.json's `bin` target).
        entry: {
            index: "src/index.ts",
            cli: "src/cli.ts",
        },
        // Emit dist/index.d.mts so consumers get the full TypeScript surface
        // (every public re-export from src/index.ts is fully typed). The
        // package.json `exports.types` condition points at this file.
        dts: true,
        // Build-time gates against publish-mistakes:
        //   publint  — verifies package.json conforms to npm's publish rules
        //              (exports map shape, file references, types ordering).
        //   attw     — Are The Types Wrong: cross-checks that the published
        //              types load correctly under every resolution mode
        //              (node10, node16, bundler, ESM, CJS).
        publint: true,
        attw: true,
        plugins: [inertCliShimsPlugin, copySchemasPlugin],
    },
    test: {
        // Only run the package's own test suite. Vendored third-party
        // snapshots (docx/, docx-templates/) carry their own tests against
        // their own deps and must not be picked up by vitest here.
        include: ["tests/**/*.test.ts"],
        exclude: ["node_modules/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        // The summary reporter prints `Coverage: N% lines, M% branches` on
        // every --coverage run, failed ones included; `reportOnFailure` below
        // keeps the provider's own reports on failed runs too.
        reporters: ["default", "./tests/support/coverage-summary-reporter.ts"],
        coverage: {
            provider: "v8",
            // lcov is required for codecov upload in ci.yml; text/html keep
            // the local `pnpm test --coverage` workflow human-readable;
            // json-summary feeds the per-module threshold gate.
            reporter: ["text", "json-summary", "html", "lcov"],
            reportOnFailure: true,
            include: ["src/**/*.ts"],
            exclude: ["src/**/*.d.ts", "tests/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        },
    },
});
