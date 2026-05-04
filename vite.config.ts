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

export default defineConfig({
    pack: {
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
        plugins: [copySchemasPlugin],
    },
    test: {
        // Only run the package's own test suite. Vendored third-party
        // snapshots (docx/, docx-templates/) carry their own tests against
        // their own deps and must not be picked up by vitest here.
        include: ["tests/**/*.test.ts"],
        exclude: ["node_modules/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        coverage: {
            provider: "v8",
            // lcov is required for codecov upload in ci.yml; text/html keep
            // the local `bun run test --coverage` workflow human-readable.
            reporter: ["text", "html", "lcov"],
            include: ["src/**/*.ts"],
            exclude: ["src/**/*.d.ts", "tests/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
        },
    },
});
