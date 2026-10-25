/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";
import {
    RELATIONSHIPS_MCE_CASES,
    documentRelsChoiceUndeclaredPrefix,
    documentRelsRootMustUnderstand,
    packageRelsRootMustUnderstand,
} from "./support/relationships-markup-compatibility-specimens";

const CODE = "word-markup-compatibility";
const PACKAGE = "_rels/.rels";
const SIDECAR = "word/_rels/document.xml.rels";
const mce = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code?.startsWith("mce-"));
const errors = (issues: readonly ValidationIssue[]) =>
    new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`));

/**
 * Session 34 (#636): Word honours markup-compatibility errors in
 * relationships parts, so word-valid and strict report them at the stored
 * part name, and quick and named runs agree with the full run (#381).
 */
describe("markup compatibility in relationships parts (#636, #381)", () => {
    const specimens = [
        {
            label: "mce-rels-root-mustunderstand",
            parts: packageRelsRootMustUnderstand,
            // s34c: recovery offer, then "Word experienced an error trying to
            // open the file." after it was declined.
            code: "mce-must-understand",
            part: PACKAGE,
            wordValidBlocks: true,
            // Lenient walks the OPC schema over the stored root and runs no
            // Word rule, so its findings stay as they were (#636 changes
            // nothing there; this pins them).
            lenient: ["error:xsd-error@_rels/.rels", "error:xsd-summary@_rels/.rels", "info:latent-styles-missing@word/styles.xml"],
        },
        {
            label: "mce-rels-choice-undeclared-prefix",
            parts: documentRelsChoiceUndeclaredPrefix,
            // s34d: recovery offer. The footer relationship disappears, so
            // word-header-footer-relationship already rejected the file.
            code: "mce-malformed",
            part: SIDECAR,
            wordValidBlocks: true,
            lenient: [
                "error:mce-malformed@word/_rels/document.xml.rels",
                "info:latent-styles-missing@word/styles.xml",
                "warning:paraid-missing-element@word/footer1.xml",
            ],
        },
        {
            label: "mce-rels-sidecar-root-mustunderstand",
            parts: documentRelsRootMustUnderstand,
            // Unobserved: session 34 probed mc:MustUnderstand only on
            // _rels/.rels (s34c), so this finding keeps its word-valid
            // warning instead of inheriting that part's blocking rule.
            code: "mce-must-understand",
            part: SIDECAR,
            wordValidBlocks: false,
            lenient: [
                "error:mce-must-understand@word/_rels/document.xml.rels",
                "info:latent-styles-missing@word/styles.xml",
                "warning:paraid-missing-element@word/footer1.xml",
            ],
        },
    ] as const;

    for (const suffix of ["docx", "docm"] as const) {
        it.each(specimens)(`${suffix}: $label`, async ({ parts, code, part, wordValidBlocks, lenient }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `specimen.${suffix}`);
                const bytes = await buildDocx(parts());
                await writeFile(file, bytes);

                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity: wordValidBlocks ? "error" : "warning", path: part }),
                ]);
                expect(wordValid.valid).toBe(!wordValidBlocks);

                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity: "error", path: part }),
                ]);
                expect(strict.valid).toBe(false);

                // Session 34 s34d also loses the footer relationship, so the
                // existing finding stays beside the new one. The unobserved
                // sidecar shape keeps its relationship, so it has none.
                if (code === "mce-malformed") {
                    expect(wordValid.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([
                        expect.objectContaining({ severity: "error", path: "word/document.xml" }),
                    ]);
                } else {
                    expect(wordValid.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
                }

                for (const selection of [{ quick: true }, { checks: [CODE] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect.soft(partial.valid, JSON.stringify(selection)).toBe(wordValid.valid);
                    expect.soft(mce(partial.issues), JSON.stringify(selection)).toEqual(mce(wordValid.issues));
                    expect
                        .soft(
                            [...errors(partial.issues)].filter((pair) => !errors(wordValid.issues).has(pair)),
                            JSON.stringify(selection),
                        )
                        .toEqual([]);
                }

                // Lenient runs no Word rule, so its findings stay as they were.
                const lenientRun = await validate(file, { profile: "lenient" });
                expect(lenientRun.issues.map((issue) => `${issue.severity}:${issue.code}@${issue.path ?? ""}`).sort()).toEqual([
                    ...lenient,
                ]);
                expect(await readFile(file)).toEqual(bytes);
            });
        });
    }

    it.each(RELATIONSHIPS_MCE_CASES)(
        "session 33's $label keeps no mce-* finding and its word-valid verdict (#632)",
        async ({ parts, clean }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts));
                for (const profile of ["word-valid", "strict"] as const) {
                    const result: { valid: boolean; issues: readonly ValidationIssue[] } = await validate(file, { profile });
                    expect.soft(mce(result.issues), profile).toEqual([]);
                    if (profile === "word-valid") expect(result.valid).toBe(clean);
                }
            });
        },
    );
});
