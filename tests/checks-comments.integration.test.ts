/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, readdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const broken = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures/broken");
const markerFixture = path.join(broken, "comments.unmatched-comment-marker.docx");
const commentIssues = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code?.startsWith("comment-"));

describe("named comment validation (#384)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "lenient"] as const)("selects comments with full-run findings under %s", async (profile) => {
        const selected = await validate(markerFixture, { profile, checks: ["comments"] });
        const full = await validate(markerFixture, { profile });
        expect(selected.issues.map((issue) => issue.code)).toEqual(
            expect.arrayContaining(["comment-marker-missing", "comment-orphan-start"]),
        );
        expect(selected.issues).toEqual(commentIssues(full.issues));
        expect(selected.metadata?.executedChecks).toEqual(["comment-markers", "comment-threading"]);
        expect(selected.xsdWalk?.started).toBe(false);
        expect(full.metadata?.executedChecks).toEqual(expect.arrayContaining(["comment-markers", "comment-threading"]));
        expect(new Set(full.metadata?.executedChecks).size).toBe(full.metadata?.executedChecks?.length);
    });

    it("keeps the full run's Word-blocking threading errors fatal", async () => {
        const file = path.join(broken, "sample-document.broken-tables.docx");
        const full = await validate(file, {});
        const selected = await validate(file, { checks: ["comments"] });
        expect(selected.issues).toEqual(commentIssues(full.issues));
        expect(selected.issues).toContainEqual(
            expect.objectContaining({ code: "comment-thread-commentid-paraid-orphan", severity: "error" }),
        );
        expect(selected.issues).toContainEqual(expect.objectContaining({ code: "comment-thread-durableid-orphan", severity: "error" }));
        expect(selected.valid).toBe(false);
    });

    it("keeps quick selection unchanged", async () => {
        const result = await validate(markerFixture, { quick: true });
        expect(commentIssues(result.issues)).toEqual([]);
        expect(result.metadata?.executedChecks).not.toContain("comment-markers");
        expect(result.metadata?.omittedChecks).toEqual(expect.arrayContaining(["comment-markers", "comment-threading"]));
    });

    it.each(["lenient", "word-valid", "strict"] as const)("preserves the soft threading rule's severity under %s", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "thread.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/comments.xml":
                        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0"><w:p w14:paraId="11111111"/></w:comment></w:comments>',
                    "word/commentsExtended.xml": '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
                }),
            );
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === "comment-thread-paraid-missing")).toEqual([
                expect.objectContaining({ severity: profile === "strict" ? "error" : "warning", path: "word/commentsExtended.xml" }),
            ]);
        });
    });

    // The orchestrated RED/GREEN runs supply the pre-change snapshot outside
    // the repository. Ordinary suite runs do not depend on that local artifact.
    it.skipIf(!process.env.COMMENTS_BASELINE)("preserves every full issue list for broken/comments.* under all profiles", async () => {
        const baseline = JSON.parse(await readFile(process.env.COMMENTS_BASELINE!, "utf8")) as Record<
            string,
            Record<string, ValidationIssue[]>
        >;
        const names = (await readdir(broken)).filter((name) => name.startsWith("comments.")).sort();
        expect(names).toEqual(Object.keys(baseline));
        for (const name of names) {
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                expect((await validate(path.join(broken, name), { profile })).issues, `${name}: ${profile}`).toEqual(
                    baseline[name]![profile],
                );
            }
        }
    });
});
