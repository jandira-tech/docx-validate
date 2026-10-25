/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { RELATIONSHIPS_MCE_CASES, RELATIONSHIPS_MCE_CODE } from "./support/relationships-markup-compatibility-specimens";

const CODE = "opc-relationships-structure";
const PACKAGE = "_rels/.rels";
const SIDECAR = "word/_rels/document.xml.rels";
const OUTSIDE_WORD = "customXml/_rels/item1.xml.rels";
const extra = '<x:Extra xmlns:x="urn:x"/>';
const append = (xml: string, child: string) => xml.replace("</Relationships>", `${child}</Relationships>`);
const main = MINIMAL_PARTS[PACKAGE]!;
const nested = main.replace(/(<Relationship\s[^>]*\/>)/, '<x:Wrap xmlns:x="urn:x">$1</x:Wrap>');
const cases = [
    { label: "minimal control", part: PACKAGE, xml: main, broken: false, missingMain: false },
    { label: "foreign sibling", part: PACKAGE, xml: append(main, extra), broken: true, missingMain: false },
    { label: "wrapped main", part: PACKAGE, xml: nested, broken: true, missingMain: true },
    { label: "wrong root", part: PACKAGE, xml: main.replaceAll("Relationships", "Foo"), broken: true, missingMain: true },
    { label: "foreign sidecar child", part: SIDECAR, xml: append(MINIMAL_PARTS[SIDECAR]!, extra), broken: true, missingMain: false },
    {
        label: "foreign child in a sidecar outside word/",
        part: OUTSIDE_WORD,
        xml: append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>', extra),
        extraParts: { "customXml/item1.xml": "<root/>" },
        broken: true,
        missingMain: false,
        advisory: true,
    },
    {
        label: "nested non-main",
        part: PACKAGE,
        xml: append(main, '<x:Wrap xmlns:x="urn:x"><Relationship Id="other" Type="urn:other" Target="word/styles.xml"/></x:Wrap>'),
        broken: true,
        missingMain: false,
    },
] as const;
const selections = [{ quick: true }, { checks: ["relationships"] }, { checks: [CODE] }];
const errors = (issues: readonly ValidationIssue[]) =>
    new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`));
const structures = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === CODE);

describe("relationships markup compatibility profile and selection parity (#632, #381)", () => {
    for (const suffix of ["docx", "docm"] as const) {
        for (const profile of ["word-valid", "strict"] as const) {
            it.each(RELATIONSHIPS_MCE_CASES)(`${suffix} ${profile}: $label`, async ({ parts, part, clean }) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `specimen.${suffix}`);
                    const bytes = await buildDocx(parts);
                    await writeFile(file, bytes);
                    const full = await validate(file, { profile });
                    expect.soft(full.valid).toBe(clean && profile === "word-valid");
                    expect
                        .soft(full.issues.filter((issue) => issue.code === RELATIONSHIPS_MCE_CODE))
                        .toEqual(
                            clean ? [expect.objectContaining({ path: part, severity: profile === "strict" ? "error" : "warning" })] : [],
                        );
                    expect
                        .soft(structures(full.issues))
                        .toEqual(
                            clean
                                ? []
                                : [expect.objectContaining({ path: part, severity: "error", message: expect.stringContaining("Wrap") })],
                        );
                    expect.soft(full.issues.filter((issue) => issue.code === "word-main-document-part")).toEqual([]);
                    expect.soft(full.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
                    if (profile === "word-valid") {
                        if (clean) expect.soft(full.issues.filter((issue) => issue.severity === "error")).toEqual([]);
                        for (const selection of [
                            { checks: [CODE, "word-main-document-part", "word-document-root", "word-header-footer-relationship"] },
                            { quick: true },
                        ]) {
                            const partial = await validate(file, { profile, ...selection });
                            expect.soft(partial.valid).toBe(full.valid);
                            expect
                                .soft(partial.issues.filter((issue) => issue.code === RELATIONSHIPS_MCE_CODE))
                                .toEqual(full.issues.filter((issue) => issue.code === RELATIONSHIPS_MCE_CODE));
                            expect.soft(structures(partial.issues)).toEqual(structures(full.issues));
                            expect.soft(partial.issues.filter((issue) => issue.code === "word-main-document-part")).toEqual([]);
                            expect.soft(partial.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
                            expect.soft([...errors(partial.issues)].filter((pair) => !errors(full.issues).has(pair))).toEqual([]);
                        }
                    }
                    expect(await readFile(file)).toEqual(bytes);
                });
            });
        }
    }
});

describe("relationships structure profile and selection parity (#433, #434, #381)", () => {
    for (const suffix of ["docx", "docm"] as const) {
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            it.each(cases)(`${suffix} ${profile}: $label`, async (specimen) => {
                await withTempDir(async (dir) => {
                    const parts = {
                        ...MINIMAL_PARTS,
                        ...("extraParts" in specimen ? specimen.extraParts : {}),
                        [specimen.part]: specimen.xml,
                    };
                    const file = path.join(dir, `specimen.${suffix}`);
                    const bytes = await buildDocx(parts);
                    await writeFile(file, bytes);
                    const full = await validate(file, { profile });
                    // Lenient keeps its existing XSD errors, without running the new Word rule.
                    // Word-valid blocks the package part and relationships parts under word/,
                    // which Word repairs (sessions 3c and 4c); other sidecars only warn.
                    const blocks = specimen.broken && (profile !== "word-valid" || !("advisory" in specimen));
                    expect(full.valid).toBe(!blocks);
                    expect(structures(full.issues)).toEqual(
                        profile === "lenient" || !specimen.broken
                            ? []
                            : [expect.objectContaining({ path: specimen.part, severity: blocks ? "error" : "warning" })],
                    );
                    // No main part blocks under Word-valid as under strict (#380, #435).
                    expect(full.issues.filter((issue) => issue.code === "word-main-document-part")).toEqual(
                        profile !== "lenient" && specimen.missingMain
                            ? [expect.objectContaining({ path: PACKAGE, severity: "error" })]
                            : [],
                    );
                    if (profile !== "lenient") expect(full.metadata?.executedChecks).toContain(CODE);
                    for (const selection of selections) {
                        if (profile === "strict" || (profile === "lenient" && !selection.quick)) {
                            await expect(validate(file, { profile, ...selection })).rejects.toThrow(ModeConflictError);
                            continue;
                        }
                        const partial = await validate(file, { profile, ...selection });
                        expect(partial.valid).toBe(profile === "lenient" || !blocks);
                        expect(structures(partial.issues)).toEqual(structures(full.issues));
                        if (full.valid) expect(partial.valid).toBe(true);
                        expect([...errors(partial.issues)].filter((pair) => !errors(full.issues).has(pair))).toEqual([]);
                    }
                    expect(await readFile(file)).toEqual(bytes);
                });
            });
        }
    }
});
