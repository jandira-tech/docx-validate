/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFile } from "node:fs/promises";
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
        expect(selected.metadata?.executedChecks).toEqual(["xml-syntax", "comment-markers", "comment-threading"]);
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

    // A part name OPC treats as equal to the literal one (ASCII case,
    // percent-encoded unreserved characters) is the same part, as for every
    // other check. The legacy comment rules matched literal basenames only,
    // so before #384 these spellings went unchecked.
    it.each(["word/comments.XML", "word/%63omments.xml"])("judges %s as the comments part under every profile", async (name) => {
        await withTempDir(async (dir) => {
            const document = MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                '<w:commentRangeStart w:id="0"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="7"/></w:r>',
            );
            const comments =
                '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p/></w:comment></w:comments>';
            const build = async (part: string, file: string) => {
                await writeFile(
                    path.join(dir, file),
                    await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": document, [part]: comments }),
                );
                return path.join(dir, file);
            };
            const literal = await build("word/comments.xml", "literal.docx");
            const equivalent = await build(name, "equivalent.docx");
            const findings = (issues: readonly ValidationIssue[]) =>
                commentIssues(issues).map(({ code, severity }) => ({ code, severity }));
            for (const profile of ["lenient", "word-valid", "strict"] as const) {
                const expected = findings((await validate(literal, { profile })).issues);
                expect(expected, profile).toContainEqual(expect.objectContaining({ code: "comment-marker-missing" }));
                expect(findings((await validate(equivalent, { profile })).issues), profile).toEqual(expected);
            }
        });
    });
});
