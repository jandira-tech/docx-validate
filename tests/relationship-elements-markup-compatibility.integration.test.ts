/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The legacy relationship walks process markup compatibility for DOCX/DOCM
 * only (#643): every observation behind #632–#636 is Word 16.115 opening
 * `.docx` files (probe sessions 33–35 in
 * docs/plans/evidence/word-probes-2026-10-24.md). PowerPoint was never shown
 * a relationships part with markup compatibility, so a PPTX keeps the stored
 * relationships view under every profile — f166c0a1's verdicts — while the
 * DOCX specimens stay processed.
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";
import { MCE_DECLARATIONS, fallback, wrapRelationship } from "./support/relationships-markup-compatibility-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PPTX_RELS = "ppt/_rels/presentation.xml.rels";

/** The PPTX specimen of #643: the slide relationship behind an unselected Choice's Fallback. */
const PPTX_SPECIMEN: Record<string, string> = {
    ...MINIMAL_PPTX_PARTS,
    [PPTX_RELS]:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" ` +
        `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:foo="urn:foo">` +
        `<Relationship Id="rId1" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<mc:AlternateContent><mc:Choice Requires="foo"/><mc:Fallback>` +
        `<Relationship Id="rId2" Type="${R}/slide" Target="slides/slide1.xml"/>` +
        `</mc:Fallback></mc:AlternateContent></Relationships>`,
};

const pinned = (issues: readonly ValidationIssue[]) =>
    issues.map((issue) => ({ severity: issue.severity, code: issue.code, path: issue.path, message: issue.message }));
const matching = (issues: readonly ValidationIssue[], ...codes: string[]): ValidationIssue[] =>
    issues.filter((issue) => codes.includes(issue.code ?? ""));

describe("PPTX relationships parts keep the stored view under every profile (#643)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "strict"] as const)(
        "%s: a slide relationship behind mc:Fallback stays unread — f166c0a1's verdict",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.pptx");
                await writeFile(file, await buildPptx(PPTX_SPECIMEN));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(matching(result.issues, "rels-id-missing")).toEqual([
                    expect.objectContaining({ severity: "error", path: "ppt/presentation.xml" }),
                ]);
                expect(matching(result.issues, "rels-unreferenced")).toEqual([
                    expect.objectContaining({ severity: "error", path: "ppt/slides/slide1.xml" }),
                ]);
            });
        },
    );

    it("lenient keeps its byte-identical output: the descendant scan reads the Fallback relationship", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.pptx");
            await writeFile(file, await buildPptx(PPTX_SPECIMEN));
            const result = await validate(file, { profile: "lenient" });
            expect(result.valid).toBe(true);
            expect(pinned(result.issues)).toEqual([]);
        });
    });

    it("DOCX keeps the processed view: session 34 s34a's split duplicate still blocks", async () => {
        await withTempDir(async (dir) => {
            const added = `<Relationship Id="rId1" Type="${R}/styles" Target="word/styles.xml"/>`;
            const parts = {
                ...MINIMAL_PARTS,
                "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace("<Relationships ", `<Relationships ${MCE_DECLARATIONS} `).replace(
                    "</Relationships>",
                    `${fallback(added)}</Relationships>`,
                ),
            };
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(false);
                expect(matching(result.issues, "rels-id-duplicate"), profile).toEqual([
                    expect.objectContaining({ severity: "error", path: "_rels/.rels" }),
                ]);
            }
        });
    });

    it("DOCX keeps the processed view: session 33 s33d's Fallback footer stays clean", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(wrapRelationship(MINIMAL_PARTS, "_rels/.rels", "rId1", fallback)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.valid).toBe(true);
        });
    });
});
