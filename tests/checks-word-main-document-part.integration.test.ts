/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { DUPLICATE_MAIN_DOCUMENT_PARTS, ORPHAN_MAIN_DOCUMENT_PARTS } from "./support/main-document-specimens";

const CODE = "word-main-document-part";
const CORE_ONLY = {
    "[Content_Types].xml":
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        "</Types>",
    "_rels/.rels":
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        "</Relationships>",
    "docProps/core.xml": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
};

describe("multiple internal officeDocument relationships (#515, Word probes session 15)", () => {
    for (const extension of ["docx", "docm"] as const) {
        for (const profile of ["word-valid", "strict"] as const) {
            const formatParts = (parts: Readonly<Record<string, string>>) => ({
                ...parts,
                "[Content_Types].xml":
                    extension === "docm"
                        ? parts["[Content_Types].xml"]!.replaceAll(
                              "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                              "application/vnd.ms-word.document.macroEnabled.main+xml",
                          )
                        : parts["[Content_Types].xml"]!,
            });

            it.each(DUPLICATE_MAIN_DOCUMENT_PARTS)(
                `${extension} ${profile} rejects $label`,
                async ({ parts }) => {
                    await withTempDir(async (dir) => {
                        const file = path.join(dir, `specimen.${extension}`);
                        await writeFile(file, await buildDocx(formatParts(parts)));
                        const result = await validate(file, { profile });
                        expect(result.valid).toBe(false);
                        expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([
                            expect.objectContaining({
                                severity: "error",
                                path: "_rels/.rels",
                                message: expect.stringMatching(/rId1.*rId9/),
                            }),
                        ]);
                        expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code)).toEqual([CODE]);
                    });
                },
                60_000,
            );

            it.each([
                { label: "single relationship", parts: MINIMAL_PARTS },
                { label: "orphan main-typed part", parts: ORPHAN_MAIN_DOCUMENT_PARTS },
            ])(
                `${extension} ${profile} preserves the $label verdict`,
                async ({ label, parts }) => {
                    await withTempDir(async (dir) => {
                        const file = path.join(dir, `control.${extension}`);
                        await writeFile(file, await buildDocx(formatParts(parts)));
                        const result = await validate(file, { profile });
                        expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                        expect(result.valid).toBe(label === "single relationship" || profile === "word-valid");
                        if (!result.valid) {
                            expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code)).toEqual([
                                "rels-unreferenced",
                            ]);
                        }
                    });
                },
                60_000,
            );
        }
    }
});

describe("missing main document part across profiles (#380)", () => {
    for (const profile of ["strict", "word-valid", "lenient"] as const) {
        it.each([
            ["empty ZIP", {}],
            ["content-types only", { "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]! }],
            ["core-properties only", CORE_ONLY],
        ])(`${profile} handles %s`, async (_name, parts) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts));
                const result = await validate(file, { profile });
                const issues = result.issues.filter((issue) => issue.code === CODE);
                if (profile === "lenient") {
                    expect(issues).toEqual([]);
                } else {
                    // Word 16.115 offers to recover these packages, so the
                    // finding is blocking under strict and word-valid (#380;
                    // session 3c in docs/plans/evidence/word-probes-2026-10-24.md).
                    expect(issues).toEqual([expect.objectContaining({ severity: "error" })]);
                    expect(result.valid).toBe(false);
                }
            });
        });

        it(`${profile} accepts the minimal package without a main-part finding`, async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.docx");
                await writeFile(file, await buildDocx());
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    }
});

describe("a quick run judges a missing main part as the full run does (#380, #381)", () => {
    it("word-valid quick reports it as an error, and lenient quick does not run it", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "core-only.docx");
            await writeFile(file, await buildDocx(CORE_ONLY));
            const quick = await validate(file, { quick: true });
            expect(quick.issues.filter((issue) => issue.code === CODE)).toEqual([expect.objectContaining({ severity: "error" })]);
            expect(quick.metadata?.executedChecks).toContain(CODE);

            const lenient = await validate(file, { profile: "lenient", quick: true });
            expect(lenient.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            expect(lenient.metadata?.omittedChecks).toContain(CODE);
        });
    });
});

describe("an officeDocument target naming a folder (#962, Word session 70 s70a/s70c)", () => {
    /** The minimal parts with the main document stored as word/main.xml (its sidecar renamed, its Override moved), named by `target`. */
    const mainAt = (target: string): Record<string, string> => {
        const parts: Record<string, string> = {
            ...MINIMAL_PARTS,
            "word/main.xml": MINIMAL_PARTS["word/document.xml"]!,
            "word/_rels/main.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!,
            "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', `Target="${target}"`),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                'PartName="/word/document.xml"',
                'PartName="/word/main.xml"',
            ),
        };
        delete parts["word/document.xml"];
        delete parts["word/_rels/document.xml.rels"];
        return parts;
    };

    // Word offered recovery for each, in both runs: a final `.` or `..`
    // segment leaves a folder, which names no main document. The
    // file-reference walk reads the folder target as broken too (#975).
    it.each(["word/main.xml/.", "/word/main.xml/.", "word/main.xml/x/..", "word/main.xml/"])(
        "rejects the main document the target %s names no part for",
        async (target) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "folder-target.docx");
                await writeFile(file, await buildDocx(mainAt(target)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                    expect.objectContaining({
                        code: "rels-broken",
                        path: "_rels/.rels",
                        message: `Broken reference to ${target}: it names a folder, not a part`,
                    }),
                    expect.objectContaining({ code: CODE, path: "_rels/.rels" }),
                ]);
            });
        },
    );

    // Word opened each clean, in both runs, and rendered the main
    // document's text: an empty segment collapses and an inner `.` resolves.
    it.each(["word/main.xml", "word//main.xml", "/word//main.xml", "word/./main.xml"])(
        "accepts the main document the target %s names",
        async (target) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "main-target.docx");
                await writeFile(file, await buildDocx(mainAt(target)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
                expect(result.valid).toBe(true);
            });
        },
    );
});

describe("an officeDocument target climbing past the package root (#986, Word session 78 s78a)", () => {
    /** The minimal parts with the main document named from the package relationships by `target`. */
    const mainAt = (target: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', `Target="${target}"`),
    });
    const judged = async (target: string, profile: "word-valid" | "strict" | "lenient") =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "climbing-main.docx");
            await writeFile(file, await buildDocx(mainAt(target)));
            return validate(file, { profile });
        });

    // Word opened each clean, in both runs, and rendered "Hello": the climb
    // stops at the package root, so each names the stored word/document.xml.
    it.each(["../word/document.xml", "../../word/document.xml", "/../word/document.xml", "word/document.xml"])(
        "accepts the main document %s names under every profile",
        async (target) => {
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const result = await judged(target, profile);
                expect(
                    result.issues.filter((issue) => issue.severity === "error"),
                    profile,
                ).toEqual([]);
                expect(result.valid, profile).toBe(true);
            }
        },
    );

    // Unprobed for officeDocument: the climb stops at the root and names the
    // absent /document.xml, as session 75's styles climb named the absent
    // /styles.xml.
    it("rejects a climb whose stopped part is absent", async () => {
        const result = await judged("../document.xml", "word-valid");
        expect(result.valid).toBe(false);
        expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([expect.objectContaining({ path: "_rels/.rels" })]);
    });
});
