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
                    expect(issues).toEqual([expect.objectContaining({ severity: profile === "strict" ? "error" : "warning" })]);
                    if (profile === "strict") expect(result.valid).toBe(false);
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
    it("word-valid quick reports it as a warning, and lenient quick does not run it", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "core-only.docx");
            await writeFile(file, await buildDocx(CORE_ONLY));
            const quick = await validate(file, { quick: true });
            expect(quick.issues.filter((issue) => issue.code === CODE)).toEqual([expect.objectContaining({ severity: "warning" })]);
            expect(quick.metadata?.executedChecks).toContain(CODE);

            const lenient = await validate(file, { profile: "lenient", quick: true });
            expect(lenient.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            expect(lenient.metadata?.omittedChecks).toContain(CODE);
        });
    });
});
