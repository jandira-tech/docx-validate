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
 * Stored-name collisions refuse the package view on every host (#441).
 * Later filesystem failures still merge readable archive findings (#454).
 */

import { chmodSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

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
        refusedView: true,
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
        refusedView: true,
        entries: [
            ...Object.entries(MINIMAL_PARTS).map(([name, text]) => ({
                name,
                content: name === "[Content_Types].xml" ? BAD_OVERRIDE_TYPES : text,
            })),
            { name: "word/ZZ.xml", content: "<x/>" },
            { name: "word/zz.xml", content: "<x/>" },
        ],
    },
    {
        name: "long-file-name",
        refusedView: false,
        entries: [
            ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({
                name,
                content: name === "[Content_Types].xml" ? BAD_OVERRIDE_TYPES : content,
            })),
            { name: `word/${"a".repeat(300)}.xml`, content: "<x/>" },
        ],
    },
] as const;

const PROFILES = ["word-valid", "lenient", "strict"] as const;
const PARTIAL_SELECTIONS = [
    { quick: true },
    { checks: ["package"] },
    { checks: ["content-types"] },
    { checks: ["ct-override-partname-invalid"] },
] as const;

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
    describe.each(SPECIMENS)("$name", ({ name, entries, refusedView }) => {
        it.each(["docx", "docm"] as const)(
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
                        // A refused view runs only stored-name/envelope checks.
                        // A later filesystem failure can still judge its parts.
                        if (refusedView) {
                            expect(
                                full.issues.some((issue) => issue.code === "ct-override-partname-invalid"),
                                tag,
                            ).toBe(false);
                            expect(full.metadata?.executedChecks, tag).not.toContain("ct-override-partname-invalid");
                        } else {
                            expect(full.issues, tag).toContainEqual(
                                expect.objectContaining({
                                    code: "ct-override-partname-invalid",
                                    path: "[Content_Types].xml",
                                    severity: severityUnder(profile),
                                }),
                            );
                            expect(full.metadata?.executedChecks, tag).toContain("ct-override-partname-invalid");
                        }
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
            60_000,
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
                    profile === "word-valid" ? ["zip-bytes-after-eocd", "zip-bytes-before-archive", "zip-eocd-comment-overrun"] : [],
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

describe("portable extraction failures (#475, #476)", { timeout: 120_000 }, () => {
    it.each(PROFILES)("preserves comment-threading severity under %s", async (profile) => {
        const parts = {
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                '<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>',
            ),
            "word/comments.xml":
                '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
                'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">' +
                '<w:comment w:id="0" w:author="A"><w:p w14:paraId="11111111" w14:textId="11111111">' +
                "<w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>",
            "word/commentsExtended.xml": '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
        };
        await withTempDir(async (dir) => {
            for (const unstageable of [false, true]) {
                const file = path.join(dir, `${unstageable ? "unstageable" : "stageable"}.docx`);
                const bytes = buildRawZip([
                    ...Object.entries(parts).map(([name, content]) => ({ name, content })),
                    ...(unstageable ? [{ name: "word", content: "file blocks word directory" }] : []),
                ]);
                writeFileSync(file, bytes);
                const result = await validate(file, { profile });
                const tag = `${profile} ${unstageable ? "unstageable" : "stageable"}`;
                expect(
                    result.issues.some((issue) => issue.code === "package-extract-failed"),
                    tag,
                ).toBe(unstageable);
                expect(result.metadata?.executedChecks, tag).toContain("comment-threading");
                expect(
                    result.issues.filter((issue) => issue.code === "comment-thread-paraid-missing"),
                    tag,
                ).toEqual([
                    expect.objectContaining({
                        code: "comment-thread-paraid-missing",
                        severity: profile === "strict" ? "error" : "warning",
                        path: "word/commentsExtended.xml",
                    }),
                ]);
                expect(readFileSync(file)).toEqual(bytes);
            }
        });
    });

    it.each(PROFILES)("keeps a reached part's XML error as the full run reports it under %s (#437)", async (profile) => {
        // word/styles.xml is reached through the document's styles
        // relationship, so word-valid blocks its malformed XML; the fallback
        // must judge it with the same Word reachability set.
        const parts = { ...MINIMAL_PARTS, "word/styles.xml": "<w:styles" };
        await withTempDir(async (dir) => {
            const found = new Map<boolean, ValidationIssue[]>();
            for (const unstageable of [false, true]) {
                const file = path.join(dir, `${unstageable ? "unstageable" : "stageable"}.docx`);
                const bytes = buildRawZip([
                    ...Object.entries(parts).map(([name, content]) => ({ name, content })),
                    ...(unstageable ? [{ name: "word", content: "file blocks word directory" }] : []),
                ]);
                writeFileSync(file, bytes);
                const result = await validate(file, { profile });
                expect(
                    result.issues.some((issue) => issue.code === "package-extract-failed"),
                    `${profile} ${unstageable}`,
                ).toBe(unstageable);
                found.set(
                    unstageable,
                    result.issues.filter((issue) => issue.code === "xml-syntax" && issue.path === "word/styles.xml"),
                );
                expect(readFileSync(file)).toEqual(bytes);
            }
            expect(found.get(false), `${profile} stageable`).toEqual([expect.objectContaining({ severity: "error" })]);
            expect(
                found.get(true)!.map((issue) => issue.severity),
                `${profile} unstageable`,
            ).toEqual(found.get(false)!.map((issue) => issue.severity));
        });
    });

    it.each(PROFILES)("merges archive findings beside a file/folder collision under %s", async (profile) => {
        const bytes = buildRawZip([
            ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({
                name,
                content: name === "[Content_Types].xml" ? BAD_OVERRIDE_TYPES : content,
            })),
            { name: "word", content: "file blocks word directory" },
        ]);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "portable.docx");
            writeFileSync(file, bytes);
            const full = await validate(file, { profile });
            expect(full.valid).toBe(false);
            expect(full.issues).toContainEqual(expect.objectContaining({ code: "package-extract-failed", severity: "error" }));
            expect(full.issues).toContainEqual(
                expect.objectContaining({
                    code: "ct-override-partname-invalid",
                    path: "[Content_Types].xml",
                    severity: severityUnder(profile),
                }),
            );
            expect(full.metadata?.executedChecks).toContain("ct-override-partname-invalid");
            expect(duplicates(full.issues)).toEqual([]);
            for (const selection of PARTIAL_SELECTIONS) {
                const tag = `${profile} ${JSON.stringify(selection)}`;
                if (profile === "strict") {
                    await expect(validate(file, { profile, ...selection }), tag).rejects.toBeInstanceOf(ModeConflictError);
                } else {
                    const partial = await validate(file, { profile, ...selection });
                    if (full.valid) expect(partial.valid, tag).toBe(true);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});

describe("math parse diagnostics agree across validation paths (#500)", { timeout: 120_000 }, () => {
    describe.each(["docx", "docm"] as const)("%s", (suffix) => {
        it.each(PROFILES)("keeps the positioned parser findings under %s", async (profile) => {
            await withTempDir(async (dir) => {
                const findings = new Map<boolean, ValidationIssue[]>();
                for (const unstageable of [false, true]) {
                    const bytes = buildRawZip([
                        ...Object.entries({ ...MINIMAL_PARTS, "word/document.xml": "<w:document" }).map(([name, content]) => ({
                            name,
                            content,
                        })),
                        ...(unstageable ? [{ name: "word", content: "file blocks word directory" }] : []),
                    ]);
                    const file = path.join(dir, `${unstageable}.${suffix}`);
                    writeFileSync(file, bytes);
                    const full = await validate(file, { profile });
                    const tag = `${suffix} ${profile} ${unstageable}`;
                    expect(
                        full.issues.some((issue) => issue.code === "package-extract-failed"),
                        tag,
                    ).toBe(unstageable);
                    findings.set(
                        unstageable,
                        full.issues.filter((issue) => issue.code?.startsWith("xml-") || issue.code === "word-math-parse"),
                    );
                    if (profile !== "strict") {
                        const checks = profile === "lenient" ? ["xml-syntax"] : ["word-math-parse"];
                        const partial = await validate(file, { profile, checks });
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                        expect.soft(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
                    expect(readFileSync(file)).toEqual(bytes);
                }
                expect(findings.get(false)).toEqual([
                    expect.objectContaining({ code: "xml-undeclared-prefix", path: "word/document.xml", severity: "error", line: 1 }),
                    expect.objectContaining({ code: "xml-syntax", path: "word/document.xml", severity: "error", line: 1 }),
                ]);
                expect(findings.get(true)).toEqual(findings.get(false));
            });
        });

        it.each(["word-valid", "lenient"] as const)("keeps valid full and partial controls valid under %s", async (profile) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `valid.${suffix}`);
                writeFileSync(file, buildRawZip(Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content }))));
                const full = await validate(file, { profile });
                const checks = profile === "lenient" ? ["xml-syntax"] : ["word-math-parse"];
                const partial = await validate(file, { profile, checks });
                expect(full.valid).toBe(true);
                expect(partial.valid).toBe(true);
                expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            });
        });
    });
});

/**
 * A packed file that cannot be read fails as an archive that does not open,
 * in every mode: one `package-open-failed` error, and nothing else ran
 * (#422). `fs.stat` still answers for a mode-000 file, so the unguarded read
 * must fail as the archive's own open failure does — with a result, never a
 * throw.
 */
describe.skipIf(process.platform === "win32")("a packed input that cannot be read (#422)", { timeout: 120_000 }, () => {
    // chmod has no effect for root.
    const runsAsRoot = process.getuid?.() === 0;

    it.skipIf(runsAsRoot)(
        "describes a read failure in every mode without blaming the archive (#547)",
        async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "packed-ro.docx");
                writeFileSync(file, await buildDocx());
                chmodSync(file, 0o000);
                const modes: readonly (readonly [string, Parameters<typeof validate>[1]])[] = [
                    ["word-valid", {}],
                    ["lenient", { profile: "lenient" }],
                    ["strict", { profile: "strict" }],
                    ["quick", { quick: true }],
                    ["named xml-syntax", { checks: ["xml-syntax"] }],
                ];
                const messages: string[] = [];
                const wordVerifier = vi.fn();
                try {
                    for (const [tag, options] of modes) {
                        const result = await validate(file, { ...options, verifyWord: true, wordVerifier }).catch((error: unknown) => {
                            throw new Error(
                                `${tag}: validate() rejected instead of resolving: ${error instanceof Error ? error.message : String(error)}`,
                            );
                        });
                        expect(result.valid, tag).toBe(false);
                        expect(result.issues, tag).toEqual([expect.objectContaining({ severity: "error", code: "package-open-failed" })]);
                        expect(result.metadata?.executedChecks, tag).toEqual([]);
                        expect(result.issues[0]!.message, tag).toMatch(/^Cannot read .*packed-ro\.docx: EACCES:/);
                        expect(result.issues[0]!.message, tag).not.toMatch(/encrypted|corrupted|truncated/);
                        expect(result.metadata?.wordVerification?.outcome, tag).toBe("unavailable");
                        messages.push(result.issues[0]!.message);
                    }
                } finally {
                    chmodSync(file, 0o644);
                }
                // Every mode reports the same failure message.
                for (const message of messages) expect(message).toBe(messages[0]);
                expect(wordVerifier).not.toHaveBeenCalled();
            });
        },
        60_000,
    );
});

it("keeps the archive explanation for JSZip refusals in full and partial runs (#547)", async () => {
    await withTempDir(async (dir) => {
        const file = path.join(dir, "corrupt.docx");
        writeFileSync(file, "not a ZIP");
        for (const options of [{}, { quick: true }, { checks: ["xml-syntax"] }]) {
            const result = await validate(file, options);
            expect(result.issues).toEqual([
                expect.objectContaining({
                    code: "package-open-failed",
                    severity: "error",
                    message: expect.stringContaining("not a valid .docx/zip — encrypted, corrupted, or truncated"),
                }),
            ]);
        }
    });
}, 60_000);

describe("root Ignorable diagnostic identity survives staging failures (#499)", { timeout: 120_000 }, () => {
    const document = MINIMAL_PARTS["word/document.xml"]!;
    const specimens = [
        { name: "root", xml: document.replace('mc:Ignorable="w14"', 'mc:Ignorable="w14 ghost"'), codes: ["ignorable-undeclared"] },
        { name: "descendant", xml: document.replace("<w:body>", '<w:body mc:Ignorable="ghost">'), codes: ["mce-unbound-prefix"] },
        {
            name: "root-and-descendant",
            xml: document.replace('mc:Ignorable="w14"', 'mc:Ignorable="w14 ghost"').replace("<w:body>", '<w:body mc:Ignorable="ghost">'),
            codes: ["ignorable-undeclared", "mce-unbound-prefix"],
        },
        {
            name: "bound-only-on-descendant",
            xml: document
                .replace('mc:Ignorable="w14"', 'mc:Ignorable="w14 ghost"')
                .replace("<w:body>", '<w:body xmlns:ghost="urn:future">'),
            codes: ["mce-unbound-prefix"],
        },
    ];
    const namespaceFindings = (issues: readonly ValidationIssue[]) =>
        issues.filter((issue) => issue.code === "ignorable-undeclared" || issue.code === "mce-unbound-prefix");

    describe.each(["docx", "docm"] as const)("%s", (suffix) => {
        describe.each(specimens)("$name", ({ name, xml, codes }) => {
            it.each(PROFILES)("preserves namespace findings under %s", async (profile) => {
                await withTempDir(async (dir) => {
                    const found = new Map<boolean, ValidationIssue[]>();
                    for (const unstageable of [false, true]) {
                        const bytes = buildRawZip([
                            ...Object.entries({ ...MINIMAL_PARTS, "word/document.xml": xml }).map(([name, content]) => ({ name, content })),
                            ...(unstageable ? [{ name: "word", content: "file blocks word directory" }] : []),
                        ]);
                        const file = path.join(dir, `${unstageable}.${suffix}`);
                        writeFileSync(file, bytes);
                        const full = await validate(file, { profile });
                        const tag = `${suffix} ${profile} ${name} ${unstageable}`;
                        expect(
                            full.issues.some((issue) => issue.code === "package-extract-failed"),
                            tag,
                        ).toBe(unstageable);
                        found.set(unstageable, namespaceFindings(full.issues));
                        if (profile !== "strict") {
                            const checks = profile === "lenient" ? ["xml-syntax"] : ["word-markup-compatibility"];
                            const partial = await validate(file, { profile, checks });
                            if (full.valid) expect.soft(partial.valid, tag).toBe(true);
                            expect.soft(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                            if (profile === "word-valid") {
                                expect.soft(namespaceFindings(partial.issues), `${tag} partial`).toEqual(found.get(false));
                            }
                        }
                        expect(readFileSync(file)).toEqual(bytes);
                    }
                    // Lenient does not run the registered Word rule in the fallback.
                    if (profile === "lenient") return;
                    expect(found.get(false)).toEqual(
                        codes.map((code) =>
                            expect.objectContaining({
                                code,
                                path: "word/document.xml",
                                severity: profile === "strict" ? "error" : "warning",
                            }),
                        ),
                    );
                    expect(found.get(true)).toEqual(found.get(false));
                });
            });
        });
    });
});
