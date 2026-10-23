/*
 * Lane integration coverage: every jubarte 0.13.0 lane specimen, table-driven (#274).
 *
 * For each of the 24 rows of tests/lane/jubarte-0.13.0/manifest.csv this pins:
 *
 *   1. the specimen's sha256 against the manifest row — a changed file
 *      invalidates the lane's historical Word observation, so it must fail
 *      loudly instead of silently re-basing;
 *   2. `validate()` validity under the `lenient` and `strict` profiles
 *      (word-valid is deliberately NOT pinned: Task 10 flips the default and
 *      Task 7 changes word-valid semantics; the orchestrator adds that
 *      column later);
 *   3. the exact set of `word-*` advisory codes present under `strict`.
 *
 * Expected values were derived from the specimens' XML and GitHub issues
 * #208–#263 (one issue per finding; the derivation is cited per row), never
 * pasted from validator output. Cross-checked against the per-check
 * integration suites in tests/checks-*.integration.test.ts.
 *
 * Kinds (from manifest.csv / issues): k01 Override-declared media (#208,
 * #237, #249 — validator false positives, since fixed, so no word-* advisory);
 * k02 w:ins inside m:r (#235, #256, #257, #245, #252, #263); k03/k05 files are
 * hash-identical to the k02 specimens (same bench redlines, other findings:
 * MCE positionH #239/#251 and chart c:ext #230/#243 are validator false
 * positives, since fixed); k04 w:del inside m:r (#213, #259); k06 empty
 * w:fill (#221, #215, and #210/#225/#220/#238/#242/#229 for the k07/k08/k09
 * files, which are hash-identical to k06-carrying rows); k07 w:gridCol
 * w:type (#209, #226, #219, #224, #233, #240); k08/k09 tblCellMar order in
 * document.xml/styles.xml (#211/#216/#222/#223/#227/#228/#231/#232/#234/
 * #236/#241/#246); k10 pgMar missing attributes (#214); k11/k13 missing
 * w:tblPr/w:tblGrid in header tables (#217, #218, #212 — the header tables'
 * gridCols also carry w:type, hence the grid-column advisory); k12
 * wne:txbxContent (#244).
 */

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

/** Minimal quoted-field CSV parser for manifest.csv (no embedded newlines). */
const parseCsvRow = (line: string): string[] => {
    const fields: string[] = [];
    let current = "";
    let quoted = false;
    for (const char of line) {
        if (quoted) {
            current += char;
            if (char === '"') quoted = false;
        } else if (char === '"') {
            quoted = true;
        } else if (char === ",") {
            fields.push(current);
            current = "";
        } else {
            current += char;
        }
    }
    fields.push(current);
    return fields;
};

interface LaneRow {
    kindId: string;
    file: string;
    sha256: string;
}

const laneRows = (): LaneRow[] => {
    const lines = readFileSync(path.join(LANE, "manifest.csv"), "utf-8")
        .split(/\r?\n/)
        .filter((line) => line.trim().length > 0);
    const header = parseCsvRow(lines[0]!);
    const kindIdx = header.indexOf("kind_id");
    const fileIdx = header.indexOf("file");
    const shaIdx = header.indexOf("sha256");
    return lines.slice(1).map((line) => {
        const cells = parseCsvRow(line);
        const file = cells[fileIdx]!;
        return { kindId: cells[kindIdx]!, file: path.basename(file), sha256: cells[shaIdx]! };
    });
};

/**
 * Derived expectations per specimen stem. Sixteen distinct files: k02_1/k03_1/k05_1
 * share bytes, as do k02_2/k03_2/k05_2 and each of {k07,k08,k09}_{1,2}.
 */
interface LaneExpectation {
    lenientValid: boolean;
    strictValid: boolean;
    wordCodes: readonly string[];
}

const EXPECTED: Readonly<Record<string, LaneExpectation>> = {
    // k01: ct-undeclared-ext was a validator false positive (now fixed) — the
    // packages' only own defect. #208/#237/#249. Lenient clean; strict still
    // fails on paragraphs without w14:paraId (62 paragraphs, 20 with paraId).
    k01_1: { lenientValid: true, strictValid: false, wordCodes: [] },
    k01_2: { lenientValid: true, strictValid: false, wordCodes: [] },
    // k02 (+hash-identical k03/k05): 11 × w:ins as direct child of m:r (#235
    // et al.) — genuine jubarte defect, XSD-invalid in every profile that
    // walks the schema; advisory under strict. The k03/k05 findings (MCE
    // positionH, chart c:ext) were validator false positives and are fixed,
    // so they contribute no word-* code and no rejection.
    k02_1: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    k02_2: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    k03_1: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    k03_2: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    k05_1: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    k05_2: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-insertion"] },
    // k04: w:del as direct child of m:r — 12 sites in k04_1 (#213), 16 in
    // k04_2 (#259).
    k04_1: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-deletion"] },
    k04_2: { lenientValid: false, strictValid: false, wordCodes: ["word-math-run-deletion"] },
    // k06: 60 × empty w:fill in document.xml (#221 k06_1, #215 k06_2).
    k06_1: { lenientValid: false, strictValid: false, wordCodes: ["word-shading-color"] },
    k06_2: { lenientValid: false, strictValid: false, wordCodes: ["word-shading-color"] },
    // k07/k08/k09 (two distinct files, three manifest kinds each): gridCol
    // w:type (38 sites, #209/#226), tblCellMar bottom-after-right in
    // document.xml (31) and styles.xml (1) (#211/#216/#231/#232 et al.), and
    // the same empty-w:fill shading defect as k06 (#210/#225 name k07_1/k07_2
    // directly; #220/#238/#242/#229 the same bytes as k08/k09).
    k07_1: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    k07_2: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    k08_1: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    k08_2: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    k09_1: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    k09_2: {
        lenientValid: false,
        strictValid: false,
        wordCodes: ["word-cell-margin-order", "word-grid-column-attributes", "word-shading-color"],
    },
    // k10: k10_1's live pgMar lacks w:header/w:footer/w:gutter (#214); k10_2's
    // sectPrChange copy lacks w:header and w:gutter, and its header1/footer1
    // tables omit the required w:tblGrid (CT_Tbl sequence).
    k10_1: { lenientValid: false, strictValid: false, wordCodes: ["word-page-margin-attributes"] },
    k10_2: { lenientValid: false, strictValid: false, wordCodes: ["word-page-margin-attributes", "word-table-properties-required"] },
    // k11/k13: header tables with no w:tblPr before w:tblGrid (#217/#218/#212);
    // their gridCols carry w:type="dxa", hence the grid-column advisory too.
    k11_1: { lenientValid: false, strictValid: false, wordCodes: ["word-grid-column-attributes", "word-table-properties-required"] },
    k11_2: { lenientValid: false, strictValid: false, wordCodes: ["word-grid-column-attributes", "word-table-properties-required"] },
    // k12: five wne:txbxContent sites in document.xml plus one in header2.xml (#244).
    k12_1: { lenientValid: false, strictValid: false, wordCodes: ["word-textbox-content-namespace"] },
    k13_1: { lenientValid: false, strictValid: false, wordCodes: ["word-grid-column-attributes", "word-table-properties-required"] },
};

const stem = (file: string): string => file.replace(/\.redline\.docx$/, "");

describe("jubarte 0.13.0 lane integration (manifest.csv)", { timeout: 120_000 }, () => {
    const rows = laneRows();

    it("manifest.csv lists all 24 specimens with derived expectations", () => {
        expect(rows).toHaveLength(24);
        for (const row of rows) {
            expect(EXPECTED[stem(row.file)], `no derived expectation for ${row.file}`).toBeDefined();
        }
    });

    for (const row of rows) {
        const file = row.file;
        const expected = EXPECTED[stem(file)]!;
        it(
            `${file} (${row.kindId}): sha256 pinned, lenient/strict validity and strict word-* advisories`,
            { timeout: 60_000 },
            async () => {
                // 1. Hash pin — a changed specimen invalidates the lane's Word
                //    evidence and every count bound to it.
                const bytes = readFileSync(path.join(LANE, file));
                const sha256 = createHash("sha256").update(bytes).digest("hex");
                expect(sha256, `${file} changed on disk; manifest.csv row is stale`).toBe(row.sha256);

                // 2. Profile validity.
                const lenient = await validate(path.join(LANE, file), { profile: "lenient" });
                expect(lenient.valid, `${file} lenient validity (derived)`).toBe(expected.lenientValid);
                const strict = await validate(path.join(LANE, file), { profile: "strict" });
                expect(strict.valid, `${file} strict validity (derived; k01 rows fail on strict-only paraid-missing-element)`).toBe(
                    expected.strictValid,
                );

                // 3. Exact set of word-* advisory codes under strict.
                const wordCodes = Array.from(
                    new Set(strict.issues.map((issue) => issue.code ?? "").filter((code) => code.startsWith("word-"))),
                ).sort();
                expect(wordCodes, `${file} strict word-* advisory set (derived)`).toEqual([...expected.wordCodes].sort());
            },
        );
    }
});
