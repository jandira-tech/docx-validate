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

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN = '<x:foo xmlns:x="urn:x"/>';
const IGNORED = `<x:document xmlns:x="urn:future:document" xmlns:w="${W}" xmlns:mc="${MC}" mc:Ignorable="x"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></x:document>`;

const withMainTarget = (target: string): string =>
    MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', `Target="${target}"`);

describe("main document root behind an OPC-equivalent target (#344)", () => {
    for (const profile of ["word-valid", "strict"] as const) {
        it.each(["WORD/document.xml", "word/%64ocument.xml"])(`${profile} rejects a foreign root behind %s`, async (target) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(
                    file,
                    await buildDocx({ ...MINIMAL_PARTS, "_rels/.rels": withMainTarget(target), "word/document.xml": FOREIGN }),
                );
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.code === "word-document-root")).toEqual([
                    expect.objectContaining({ severity: "error", path: "word/document.xml" }),
                ]);
            });
        });

        it(`${profile} draws no root finding for the control behind WORD/document.xml`, async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.docx");
                await writeFile(file, await buildDocx({ ...MINIMAL_PARTS, "_rels/.rels": withMainTarget("WORD/document.xml") }));
                const result = await validate(file, { profile });
                expect(result.issues.filter((issue) => issue.code === "word-document-root")).toEqual([]);
                // Strict still fails the relationship checks, which ignore OPC
                // equivalence (#346); word-valid only warns about them.
                if (profile === "word-valid") expect(result.valid).toBe(true);
            });
        });
    }
});

describe("main document root validation (#317)", () => {
    for (const profile of [undefined, "word-valid", "strict"] as const) {
        it.each([
            ["foreign root", FOREIGN, /refuses to open/],
            ["ignored root", IGNORED, /drops all content/],
        ] as const)(`${profile ?? "default"} rejects %s`, async (_name, xml, message) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": xml }));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.code === "word-document-root")).toEqual([
                    expect.objectContaining({ severity: "error", path: "word/document.xml", message: expect.stringMatching(message) }),
                ]);
                expect(result.metadata?.executedChecks).toContain("word-document-root");
                if (profile !== "strict") expect(result.xsdWalk?.started).toBe(false);
            });
        });

        it(`${profile ?? "default"} accepts the normal control`, async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.docx");
                await writeFile(file, await buildDocx());
                const result = await validate(file, { profile });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === "word-document-root")).toEqual([]);
            });
        });
    }
});
