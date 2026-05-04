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
        plugins: [copySchemasPlugin],
    },
    test: {
        // Only run the package's own test suite. Vendored third-party
        // snapshots (docx/, docx-templates/) carry their own tests against
        // their own deps and must not be picked up by vitest here.
        include: ["tests/**/*.test.ts"],
        exclude: ["node_modules/**", "dist/**", "docx/**", "docx-templates/**", "dotgithubtoport/**"],
    },
});
