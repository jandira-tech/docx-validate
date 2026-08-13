/**
 * The Desired Conversion Methodology (DCM) definition is hand-mirrored across
 * AGENTS.md / README.md / CLAUDE.md (clause (ii) of the definition itself
 * names them). Mirrors drift; this test stops that: the definition must stay
 * byte-identical across every surface that carries it.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, test } from "vitest";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SURFACES = ["AGENTS.md", "README.md", "CLAUDE.md"];
const START = "**Desired Conversion Methodology** means";
const END = "`CLAUDE.md`.";

function dcmSection(file: string): string {
    const text = readFileSync(path.join(ROOT, file), "utf8").replace(/\r\n/g, "\n");
    const i = text.indexOf(START);
    expect(i, `${file} must contain the DCM definition`).toBeGreaterThanOrEqual(0);
    expect(text.lastIndexOf(START), `${file} must contain the DCM start marker exactly once`).toBe(i);
    const j = text.indexOf(END, i);
    expect(j, `${file} DCM definition must end with the clause (ii) sentence`).toBeGreaterThan(i);
    expect(text.lastIndexOf(END), `${file} must contain the DCM end marker exactly once`).toBe(j);
    return text.slice(i, j + END.length);
}

describe("Desired Conversion Methodology — doc-surface consistency", () => {
    test("the definition is byte-identical across all doc surfaces", () => {
        const canonical = dcmSection(SURFACES[0]);
        for (const file of SURFACES.slice(1)) {
            expect(dcmSection(file), `${file} drifted from ${SURFACES[0]}`).toBe(canonical);
        }
    });
});
