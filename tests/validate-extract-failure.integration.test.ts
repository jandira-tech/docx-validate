/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * A full run that cannot extract still reports what it can read from the
 * archive (#454).
 *
 * A packed file whose stored names collide on the host filesystem — two
 * spellings of one name, as `[Content_Types].xml` and `[CONTENT_TYPES].XML`
 * are on APFS — opens as an archive and builds an in-memory package view, but
 * cannot be staged on disk. The full run stopped there, so every finding a
 * partial run computes on the archive context was dropped: the partial run's
 * errors were not a subset of the full run's, which breaks #381.
 *
 * The full run must therefore merge the archive-backed findings into the
 * staging failure, under the profile, without changing the verdict: the
 * package still cannot be staged, so it stays invalid.
 */

import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

/**
 * Whether the host filesystem folds case. Two stored names that differ only
 * in case are one file there, so the extraction fails; on a case-sensitive
 * host it does not, and the specimens cannot reproduce the defect at all.
 */
const CASE_INSENSITIVE = await withTempDir(async (dir) => {
    writeFileSync(path.join(dir, "a"), "x");
    return existsSync(path.join(dir, "A"));
});

/** The content-types stream with an `Override` whose PartName names no part. */
const BAD_OVERRIDE_TYPES = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
    "</Types>",
    `<Override PartName="/word/no%2Fthere.xml" ContentType="application/xml"/></Types>`,
);

/**
 * The specimens: a case-only collision in the stored names, beside the
 * malformed `Override` the archive holds and a partial run reports.
 */
const SPECIMENS = [
    {
        name: "content-types-stream",
        entries: [
            ...Object.entries(MINIMAL_PARTS)
                .filter(([name]) => name !== "[Content_Types].xml")
                .map(([name, text]) => ({ name, content: text })),
            { name: "[Content_Types].xml", content: BAD_OVERRIDE_TYPES },
            { name: "[CONTENT_TYPES].XML", content: MINIMAL_PARTS["[Content_Types].xml"]! },
        ],
    },
    {
        name: "word-part",
        entries: [
            ...Object.entries(MINIMAL_PARTS).map(([name, text]) => ({
                name,
                content: name === "[Content_Types].xml" ? BAD_OVERRIDE_TYPES : text,
            })),
            { name: "word/ZZ.xml", content: "<x/>" },
            { name: "word/zz.xml", content: "<x/>" },
        ],
    },
] as const;

const PROFILES = ["word-valid", "lenient", "strict"] as const;
const PARTIAL_SELECTIONS = [{ quick: true }, { checks: ["content-types"] }, { checks: ["ct-override-partname-invalid"] }] as const;

/** The error findings of a run, as `code@path`. */
const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code ?? ""}@${issue.path ?? ""}`);

/** The findings a run reports twice, if any. */
const duplicates = (issues: readonly ValidationIssue[]): string[] => {
    const seen = new Set<string>();
    const repeated: string[] = [];
    for (const issue of issues) {
        const key = `${issue.code ?? ""}@${issue.path ?? ""}@${issue.message}`;
        if (seen.has(key)) repeated.push(key);
        else seen.add(key);
    }
    return repeated;
};

/** The severity the profile gives the finding. */
const severityUnder = (profile: (typeof PROFILES)[number]): "error" | "warning" => (profile === "lenient" ? "warning" : "error");

describe("a full run that cannot extract still reports what the archive holds (#454)", { timeout: 120_000 }, () => {
    describe.each(SPECIMENS)("$name", ({ name, entries }) => {
        // On a case-sensitive host the two spellings are two files, so the
        // extraction succeeds and the defect cannot be reproduced.
        it.runIf(CASE_INSENSITIVE).each(["docx", "docm"] as const)(
            "keeps the partial runs a subset of the full %s run, under every profile",
            async (suffix) => {
                await withTempDir(async (dir) => {
                    const bytes = buildRawZip(entries);
                    const file = path.join(dir, `specimen.${suffix}`);
                    writeFileSync(file, bytes);

                    for (const profile of PROFILES) {
                        const full = await validate(file, { profile });
                        const tag = `${name} ${suffix} ${profile}`;

                        // The package cannot be staged, so the verdict does not
                        // change: it stays invalid.
                        expect(full.valid, tag).toBe(false);
                        expect(full.issues, tag).toContainEqual(
                            expect.objectContaining({ code: "package-extract-failed", severity: "error" }),
                        );
                        // The finding the archive holds, which every partial
                        // run of it reports.
                        expect(full.issues, tag).toContainEqual(
                            expect.objectContaining({
                                code: "ct-override-partname-invalid",
                                path: "[Content_Types].xml",
                                severity: severityUnder(profile),
                            }),
                        );
                        // The merged findings come from checks that ran.
                        expect(full.metadata?.executedChecks, tag).toContain("ct-override-partname-invalid");
                        // Nothing is reported twice.
                        expect(duplicates(full.issues), tag).toEqual([]);

                        if (profile === "strict") {
                            // Strict is the complete mode: it has no partial
                            // selection to compare with.
                            await expect(validate(file, { profile, quick: true }), tag).rejects.toBeInstanceOf(ModeConflictError);
                            await expect(validate(file, { profile, checks: ["content-types"] }), tag).rejects.toBeInstanceOf(
                                ModeConflictError,
                            );
                            continue;
                        }

                        for (const selection of PARTIAL_SELECTIONS) {
                            const partial = await validate(file, { profile, ...selection });
                            const selectionTag = `${tag} ${JSON.stringify(selection)}`;
                            // #381: a valid full run implies a valid partial run.
                            if (full.valid) expect(partial.valid, selectionTag).toBe(true);
                            // #381: the partial run's errors are a subset of the
                            // full run's, compared as `code@path`.
                            expect(errorPairs(full.issues), selectionTag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        }
                    }

                    // The input is never modified.
                    expect(readFileSync(file)).toEqual(bytes);
                });
            },
        );
    });
});

describe("the paths that are already consistent stay unchanged (#454)", { timeout: 120_000 }, () => {
    it("keeps package-open-failed for bytes that are not a ZIP", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "notzip.docx");
            writeFileSync(file, Buffer.from("this is not a zip archive at all"));
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                const tag = `not-a-zip ${profile}`;

                // Nothing could be read, so nothing else is reported.
                expect(
                    full.issues.map((issue) => issue.code),
                    tag,
                ).toEqual(["package-open-failed"]);
                // Only the ZIP-envelope checks read the bytes, and lenient
                // runs neither them nor a Word rule.
                expect(full.metadata?.executedChecks, tag).toEqual(
                    profile === "word-valid" ? ["zip-bytes-after-eocd", "zip-eocd-comment-overrun"] : [],
                );
                const partial = await validate(file, { profile, quick: true });
                expect(
                    partial.issues.map((issue) => issue.code),
                    tag,
                ).toEqual(["package-open-failed"]);
                expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            }
        });
    });

    it("keeps the refused package view as the stored-name findings and the extract failure", async () => {
        // A stored name that places no part refuses the package view (#378):
        // the archive context holds no part, so no part-based check can run.
        const bytes = buildRawZip([
            ...Object.entries(MINIMAL_PARTS).map(([name, text]) => ({ name, content: text })),
            { name: "word/a/../b.xml", content: "<x/>" },
        ]);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "refused.docx");
            writeFileSync(file, bytes);
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                const tag = `refused ${profile}`;

                expect(full.valid, tag).toBe(false);
                expect(
                    full.issues.map((issue) => issue.code),
                    tag,
                ).toContain("package-extract-failed");
                // The stored-name diagnosis ran; no part-based check did.
                expect(full.metadata?.executedChecks, tag).toContain("opc-equivalent-part-names");
                expect(full.metadata?.executedChecks, tag).not.toContain("ct-override-partname-invalid");
                const partial = await validate(file, { profile, checks: ["ct-override-partname-invalid"] });
                expect(
                    partial.issues.map((issue) => issue.code),
                    tag,
                ).toEqual(["package-extract-failed"]);
                expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});
