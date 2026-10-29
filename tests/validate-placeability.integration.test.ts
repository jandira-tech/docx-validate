/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdirSync, symlinkSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? "-"}`);

const selections = [{ quick: true }, { checks: ["xml-syntax"] }, { checks: ["package"] }] as const;
const rawPackage = (parts: Readonly<Record<string, string>>) =>
    buildRawZip(Object.entries(parts).map(([name, content]) => ({ name, content, utf8: true })));

// Raw central names, including NUL and both Unicode spellings, must survive
// the builder. The minimal package is the same control the Word probes use.
const HOST_NAME_PARTS = [
    { name: "case", parts: { ...MINIMAL_PARTS, "word/header1.xml": "<x/>", "word/HEADER1.xml": "<w:hdr" } },
    { name: "normalization", parts: { ...MINIMAL_PARTS, "word/caf\u00e9.xml": "<x/>", "word/cafe\u0301.xml": "<w:hdr" } },
    { name: "nul", parts: { ...MINIMAL_PARTS, "word/x\0y.xml": "<x/>" } },
];

describe("partial/full placeability parity (#423)", () => {
    for (const profile of ["word-valid", "lenient"] as const) {
        it.each(HOST_NAME_PARTS)(
            `${profile}: $name preserves the error subset for every selection`,
            async ({ parts }) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, "specimen.docx");
                    writeFileSync(file, rawPackage(parts));
                    const full = await validate(file, { profile });
                    for (const selection of selections) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = JSON.stringify(selection);
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                    }
                });
            },
            60_000,
        );

        it(`${profile}: a valid control stays valid for every selection`, async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "control.docx");
                writeFileSync(file, rawPackage(MINIMAL_PARTS));
                expect((await validate(file, { profile })).valid).toBe(true);
                for (const selection of selections) expect((await validate(file, { profile, ...selection })).valid).toBe(true);
            });
        }, 60_000);
    }
});

const directorySpecimen = (dir: string, stored: string): string => {
    const root = path.join(dir, "directory.docx");
    mkdirSync(root);
    for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
        writeFileSync(path.join(root, name.replaceAll("/", "\\")), content);
    }
    writeFileSync(path.join(root, stored), "<w:document");
    return root;
};

describe.skipIf(process.platform === "win32")("directory placeability parity (#424)", () => {
    for (const profile of ["word-valid", "lenient"] as const) {
        it.each(["word\\.\\x.xml", "word\\\\x.xml", "word\\a\\..\\x.xml"])(
            `${profile}: refuses the stored name %j before judging its XML`,
            async (stored) => {
                await withTempDir(async (dir) => {
                    const root = directorySpecimen(dir, stored);
                    const full = await validate(root, { profile });
                    const failure = full.issues.find((issue) => issue.code === "package-extract-failed");
                    expect(failure).toBeDefined();
                    for (const selection of selections) {
                        const partial = await validate(root, { profile, ...selection });
                        const tag = JSON.stringify(selection);
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        expect(partial.valid, tag).toBe(false);
                        expect(
                            partial.issues.find((issue) => issue.code === "package-extract-failed"),
                            tag,
                        ).toEqual(failure);
                        expect(partial.metadata?.executedChecks, tag).not.toContain("xml-syntax");
                    }
                });
            },
            60_000,
        );
    }

    // A refused stored name fails the directory before any stored file is
    // read, so it wins over an unreadable entry in the same tree (#426) —
    // in the full run and in every partial run alike.
    it("refuses the stored name first when the tree also holds an unreadable entry (#424, #426)", async () => {
        await withTempDir(async (dir) => {
            const stored = "word\\.\\x.xml";
            const root = directorySpecimen(dir, stored);
            symlinkSync(path.join(dir, "missing-target.xml"), path.join(root, "word\\footer9.xml"));
            const full = await validate(root, { profile: "word-valid" });
            const failures = full.issues.filter((issue) => issue.code === "package-extract-failed");
            expect(failures).toHaveLength(1);
            expect(failures[0]!.message).toContain(JSON.stringify(stored));
            expect(failures[0]!.message).not.toContain("Cannot read");
            for (const selection of selections) {
                const partial = await validate(root, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(
                    partial.issues.filter((issue) => issue.code === "package-extract-failed"),
                    tag,
                ).toEqual(failures);
            }
        });
    }, 60_000);
});

const unclosed = (xml: string): string => xml.slice(0, xml.lastIndexOf("</"));
const COLLISION_PARTS = [
    {
        name: "core",
        parts: {
            ...MINIMAL_PARTS,
            "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
            ),
            "docProps/core.xml": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">',
            "DOCPROPS/CORE.XML": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">',
        },
    },
    {
        name: "main",
        parts: {
            ...MINIMAL_PARTS,
            "word/document.xml": unclosed(MINIMAL_PARTS["word/document.xml"]!),
            "WORD/DOCUMENT.XML": unclosed(MINIMAL_PARTS["word/document.xml"]!),
        },
    },
];
const collisionSelections = [{ quick: true }, { checks: ["xml"] }, { checks: ["package", "xml"] }] as const;

describe("host-independent collision refusal (#441)", () => {
    it.each(COLLISION_PARTS)(
        "$name: full and partial views refuse before reading parts",
        async ({ parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "collision.docx");
                writeFileSync(file, rawPackage(parts));
                for (const profile of ["word-valid", "lenient", "strict"] as const) {
                    const full = await validate(file, { profile });
                    const failure = full.issues.find((issue) => issue.code === "package-extract-failed");
                    expect(failure).toBeDefined();
                    // This decision must precede extraction, even on a case-sensitive host.
                    expect(full.issues.some((issue) => issue.code === "xml-syntax")).toBe(false);
                    expect(full.metadata?.executedChecks).not.toContain("xml-syntax");
                    if (profile === "strict") {
                        for (const selection of collisionSelections) {
                            await expect(validate(file, { profile, ...selection })).rejects.toBeInstanceOf(ModeConflictError);
                        }
                        continue;
                    }
                    for (const selection of collisionSelections) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(partial.valid).toBe(false);
                        expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        expect(partial.issues.find((issue) => issue.code === "package-extract-failed")).toEqual(failure);
                        expect(partial.metadata?.executedChecks).not.toContain("xml-syntax");
                    }
                }
            });
        },
        60_000,
    );
});
