import { readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { MemoryPartFS } from "../src/lib/part-fs";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { createBrowserEngine, createNodeEngine } from "../src/lib/xsd-engine/index";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SCHEMAS = path.resolve(HERE, "..", "src/scripts/office/schemas");

// Build the in-memory schema bundle the WASM engine needs: every .xsd keyed by
// its path relative to the schema root (so xs:include/import resolve).
function schemaBundle(): Record<string, string> {
    const bundle: Record<string, string> = {};
    const walk = (dir: string) => {
        for (const entry of readdirSync(dir)) {
            const full = path.join(dir, entry);
            if (statSync(full).isDirectory()) walk(full);
            else if (full.endsWith(".xsd")) bundle[path.relative(SCHEMAS, full).split(path.sep).join("/")] = readFileSync(full, "utf-8");
        }
    };
    walk(SCHEMAS);
    return bundle;
}

async function partsOf(docxPath: string): Promise<Array<[string, Uint8Array]>> {
    const zip = await JSZip.loadAsync(readFileSync(docxPath));
    const parts: Array<[string, Uint8Array]> = [];
    for (const name of Object.keys(zip.files)) {
        const e = zip.files[name];
        if (e.dir) continue;
        parts.push([name, await e.async("uint8array")]);
    }
    return parts;
}

// Paths that the validator flags with an XSD error, under a given engine.
async function xsdErrorPaths(docxPath: string, engine: "node" | "wasm", bundle: Record<string, string>): Promise<string[]> {
    const parts = await partsOf(docxPath);
    const xsdEngine = engine === "node" ? createNodeEngine({ schemasDir: SCHEMAS }) : createBrowserEngine({ schemaBundle: bundle });
    const v = new DOCXSchemaValidator({ partFS: new MemoryPartFS(parts), profile: "strict", schemasDir: SCHEMAS, xsdEngine });
    const res = await v.validateAgainstXsd();
    return [...new Set(res.issues.filter((i) => i.code === "xsd-summary" && i.path).map((i) => i.path as string))].sort();
}

describe("XSD engine parity: libxml2-wasm (browser) vs libxmljs2 (node)", () => {
    const bundle = schemaBundle();
    const FIX = path.resolve(HERE, "fixtures");
    const cases = [
        "working/sample-document.afterword-repaired-word-repaired.docx",
        "broken/sample-document.broken-tables.docx",
    ].map((p) => path.join(FIX, p));

    it("the schema bundle has the full OOXML set", () => {
        expect(Object.keys(bundle).length).toBeGreaterThanOrEqual(35);
        expect(bundle["ISO-IEC29500-4_2016/wml.xsd"]).toMatch(/<xs(d)?:schema/);
    });

    for (const docx of cases) {
        it(`agrees on which parts are XSD-invalid: ${path.basename(docx)}`, async () => {
            const [nodePaths, wasmPaths] = await Promise.all([
                xsdErrorPaths(docx, "node", bundle),
                xsdErrorPaths(docx, "wasm", bundle),
            ]);
            expect(wasmPaths).toEqual(nodePaths);
        });
    }
});
