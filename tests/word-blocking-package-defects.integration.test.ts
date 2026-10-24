/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word 16.115 offers to recover a package without `[Content_Types].xml`, an
 * empty ZIP, and a package whose main document part is gone (#435, #380;
 * evidence in docs/plans/evidence/word-probes-2026-10-24.md, session 3c). So
 * `ct-missing` and `word-main-document-part` are Word-blocking: errors under
 * word-valid DOCX/DOCM in full runs and in every partial run that reaches
 * them (#381). `rels-broken`, `rels-target-missing` and `rels-unreferenced`
 * keep their tolerated policy, and `main-relationship-missing` is
 * Word-inconclusive — it is the `word-main-document-part` promotion that
 * rejects it.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const MAIN_PART = "word-main-document-part";

const withoutParts = (...names: string[]) => Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => !names.includes(name)));

// The session 3c `main-relationship-missing` shape: the officeDocument
// relationship removed, the main part kept (`rels-unreferenced`,
// `rels-empty-part`, `word-main-document-part` findings).
const NO_MAIN_RELATIONSHIP = {
    ...MINIMAL_PARTS,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR}"/>`,
};

interface Specimen {
    readonly name: string;
    readonly bytes: () => Promise<Buffer>;
    /** Codes that block under word-valid, with the path the finding carries. */
    readonly blocking: readonly (readonly [string, string | undefined])[];
    /** Codes that stay warnings under word-valid. */
    readonly tolerated?: readonly string[];
    /** Error codes the full lenient run reports. */
    readonly lenientErrors: readonly string[];
    /** The named check that reports the blocking code, if one is registered. */
    readonly check?: string;
}

const SPECIMENS: readonly Specimen[] = [
    {
        name: "no-content-types",
        bytes: () => buildDocx(withoutParts("[Content_Types].xml")),
        blocking: [["ct-missing", undefined]],
        lenientErrors: ["ct-missing"],
        // `ct-missing` is emitted by the validator's own content-types pass:
        // no registered check reports it, so no partial run can (#435).
    },
    {
        name: "empty-archive",
        bytes: async () => buildRawZip([]),
        blocking: [
            ["ct-missing", undefined],
            [MAIN_PART, undefined],
        ],
        lenientErrors: ["ct-missing"],
        check: MAIN_PART,
    },
    {
        name: "main-part-missing",
        bytes: () => buildDocx(withoutParts("word/document.xml")),
        blocking: [[MAIN_PART, "_rels/.rels"]],
        tolerated: ["rels-broken", "rels-target-missing"],
        lenientErrors: ["rels-broken", "rels-target-missing"],
        check: MAIN_PART,
    },
    {
        name: "main-relationship-missing",
        bytes: () => buildDocx(NO_MAIN_RELATIONSHIP),
        blocking: [[MAIN_PART, "_rels/.rels"]],
        tolerated: ["rels-unreferenced"],
        lenientErrors: ["rels-unreferenced"],
        check: MAIN_PART,
    },
];

const specimen = async (dir: string, name: string, bytes: Buffer): Promise<string> => {
    const file = path.join(dir, name);
    writeFileSync(file, bytes);
    return file;
};

const issuesWith = (result: { issues: readonly ValidationIssue[] }, code: string, severity: string): ValidationIssue[] =>
    result.issues.filter((issue) => issue.code === code && issue.severity === severity);

/** An error finding at `code`, with `path` asserted only when it carries one. */
const blockingAt = (code: string, issuePath: string | undefined) =>
    expect.objectContaining({ code, severity: "error", ...(issuePath !== undefined ? { path: issuePath } : {}) });

/** Error findings as `code@path`, for the #381 subset invariant. */
const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    [...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code ?? ""}@${issue.path ?? ""}`))].sort();

describe("package-defect specimens Word offers to recover (#435, #380)", { timeout: 120_000 }, () => {
    describe.each(SPECIMENS)("$name", (spec) => {
        it.each(["docx", "docm"] as const)(
            "is rejected by word-valid .%s, with each blocking code an error at its path",
            async (suffix) => {
                await withTempDir(async (dir) => {
                    const file = await specimen(dir, `${spec.name}.${suffix}`, await spec.bytes());
                    const result = await validate(file, { profile: "word-valid" });

                    expect(result.metadata?.profile).toBe("word-valid");
                    expect(result.valid).toBe(false);
                    for (const [code, issuePath] of spec.blocking) {
                        expect(result.issues).toContainEqual(blockingAt(code, issuePath));
                    }
                    for (const code of spec.tolerated ?? []) {
                        expect(issuesWith(result, code, "warning"), code).not.toEqual([]);
                        expect(issuesWith(result, code, "error"), code).toEqual([]);
                    }
                });
            },
        );

        it.each(["docx", "docm"] as const)("keeps each blocking code an error under strict .%s", async (suffix) => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${spec.name}.${suffix}`, await spec.bytes());
                const result = await validate(file, { profile: "strict" });

                expect(result.valid).toBe(false);
                for (const [code, issuePath] of spec.blocking) {
                    expect(result.issues).toContainEqual(blockingAt(code, issuePath));
                }
            });
        });

        it.each(["docx", "docm"] as const)("keeps the existing lenient verdict on .%s", async (suffix) => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${spec.name}.${suffix}`, await spec.bytes());
                const result = await validate(file, { profile: "lenient" });

                // Lenient never runs the Word rules, so word-main-document-part
                // cannot appear; these packages fail on the validator's own
                // findings only.
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.code === MAIN_PART)).toEqual([]);
                for (const code of spec.lenientErrors) {
                    expect(result.issues).toContainEqual(expect.objectContaining({ code, severity: "error" }));
                }
            });
        });

        it("reports the same blocking errors in word-valid quick and named runs (#381)", async () => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${spec.name}.docx`, await spec.bytes());
                const full = await validate(file, { profile: "word-valid" });
                const selections = [{ quick: true }, { checks: ["relationships"] }, { checks: [MAIN_PART] }] as const;

                for (const selection of selections) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    const tag = `${spec.name} ${JSON.stringify(selection)}`;

                    // #381: a valid full run means a valid partial run, and
                    // partial errors are a subset of full errors as code@path.
                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    if (spec.check !== undefined) {
                        expect(partial.metadata?.executedChecks, tag).toContain(spec.check);
                        for (const [code, issuePath] of spec.blocking.filter(([code]) => code === spec.check)) {
                            expect(partial.issues, tag).toContainEqual(blockingAt(code, issuePath));
                        }
                    }
                }

                if (spec.check === undefined) {
                    // ct-missing has no registered check: a partial run cannot
                    // report it, so word-valid partial runs stay valid here.
                    const quick = await validate(file, { profile: "word-valid", quick: true });
                    expect(quick.valid).toBe(true);
                    expect(quick.issues.filter((issue) => issue.code === "ct-missing")).toEqual([]);
                }
            });
        });

        it("keeps lenient partial runs unchanged and refuses named Word rules", async () => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${spec.name}.docx`, await spec.bytes());
                const quick = await validate(file, { profile: "lenient", quick: true });

                // A lenient quick run executes neither the Word rules nor the
                // ZIP-envelope checks, so it sees none of these defects.
                expect(quick.valid).toBe(true);
                expect(quick.metadata?.omittedChecks).toContain(MAIN_PART);
                for (const selection of [{ checks: ["relationships"] }, { checks: [MAIN_PART] }] as const) {
                    await expect(validate(file, { profile: "lenient", ...selection }), JSON.stringify(selection)).rejects.toBeInstanceOf(
                        ModeConflictError,
                    );
                }
            });
        });

        it("keeps strict refusing partial runs", async () => {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `${spec.name}.docx`, await spec.bytes());
                for (const selection of [{ quick: true }, { checks: ["relationships"] }] as const) {
                    await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toBeInstanceOf(
                        ModeConflictError,
                    );
                }
            });
        });
    });

    it.each(["word-valid", "strict", "lenient"] as const)("keeps the minimal package valid under %s", async (profile) => {
        await withTempDir(async (dir) => {
            for (const suffix of ["docx", "docm"] as const) {
                const file = await specimen(dir, `control.${suffix}`, await buildDocx());
                const result = await validate(file, { profile });
                expect(result.valid, `${profile} .${suffix}: ${JSON.stringify(result.issues)}`).toBe(true);
            }
        });
    });

    it("keeps the minimal package valid in a word-valid quick run", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "control.docx", await buildDocx());
            const result = await validate(file, { quick: true });
            expect(result.valid).toBe(true);
        });
    });
});
