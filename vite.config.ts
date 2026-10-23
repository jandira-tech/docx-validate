/// <reference types="vitest" />
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { dts } from "rolldown-plugin-dts";
import { defineConfig } from "vitest/config";

const here = path.dirname(fileURLToPath(import.meta.url));
const SCHEMAS_SRC = path.resolve(here, "src/scripts/office/schemas");
const SCHEMAS_DST = path.resolve(here, "dist/schemas");

// The vite build collapses TypeScript to dist/index.mjs and does not copy
// non-TS assets. The bundled OOXML XSDs at src/scripts/office/schemas/
// must ship inside dist/ so defaultSchemasDir() in validators/base.ts resolves
// against the published bundle (not the source tree, which is absent on npm).
const copySchemasPlugin = {
    name: "docx-validate:copy-schemas",
    async closeBundle() {
        await fs.rm(SCHEMAS_DST, { recursive: true, force: true });
        await fs.cp(SCHEMAS_SRC, SCHEMAS_DST, { recursive: true });
    },
};

// Every runtime dependency stays external: the published bundle must import
// them (single-dependency rule for consumers), not inline them.
const EXTERNAL_DEPENDENCIES = [/^@xmldom\/xmldom($|\/)/, /^commander($|\/)/, /^jszip($|\/)/, /^libxml2-wasm($|\/)/, /^tmp($|\/)/, /^xpath($|\/)/];

export default defineConfig({
    build: {
        lib: {
            entry: path.resolve(here, "src/index.ts"),
            formats: ["es"],
            // package.json `exports["."].import` points here; the declaration
            // twin (dist/index.d.mts) is emitted by the dts() plugin below.
            fileName: () => "index.mjs",
        },
        rollupOptions: {
            external: EXTERNAL_DEPENDENCIES,
        },
        // npm consumers get the readable bundle; minification buys nothing
        // here and obscifies stack traces from bug reports.
        minify: false,
    },
    plugins: [
        // Bundles the TypeScript declarations into a single dist/index.d.mts
        // so consumers get the full typed surface of every public re-export
        // from src/index.ts (package.json `exports.types` points at it).
        dts(),
        copySchemasPlugin,
    ],
    test: {
        // Only run the package's own test suite. Vendored third-party
        // snapshots (docx/, docx-templates/) carry their own tests against
        // their own deps and must not be picked up by vitest here.
        include: ["tests/**/*.test.ts"],
        exclude: ["node_modules/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        coverage: {
            provider: "v8",
            // lcov is required for codecov upload in ci.yml; text/html keep
            // the local `pnpm test --coverage` workflow human-readable.
            reporter: ["text", "html", "lcov"],
            include: ["src/**/*.ts"],
            exclude: ["src/**/*.d.ts", "tests/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        },
    },
});
