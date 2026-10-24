/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue, ValidationResult } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const EQUIVALENT = "opc-equivalent-part-names";
const ESCAPED = "opc-part-name-escaped-slash";
const EXTRACT_FAILED = "package-extract-failed";
const DOC = "word/document.xml";
const PROFILES = ["word-valid", "lenient", "strict"] as const;

const withCode = (issues: readonly ValidationIssue[], code: string) => issues.filter((issue) => issue.code === code);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const withTrailing = (bytes: Buffer, count: number, fake?: Buffer): Buffer =>
    Buffer.concat(fake ? [bytes, fake, Buffer.alloc(count - fake.length)] : [bytes, Buffer.alloc(count)]);
// An EOCD signature with no readable central directory behind it.
const fakeEocd = (): Buffer => Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(18)]);
// A prepended archive. Its record is structural when its offsets are read as
// absolute and the live archive's is not; JSZip loads the live one, treating
// the prefix as prepended data (`reader.zero`), so the stored names must be
// the live archive's too (#462).
const prefixed = (prefixEntries: readonly { name: string; content: string }[], live: Buffer): Buffer =>
    Buffer.concat([buildRawZip([{ name: "pad.bin", content: Buffer.alloc(6000) }, ...prefixEntries]), live]);

// Partial runs agree with the full run (#381): a valid full run implies a
// valid partial run, and partial errors are a subset of full errors.
const expectParity = (full: ValidationResult, partial: ValidationResult, tag: string) => {
    if (full.valid) expect(partial.valid, tag).toBe(true);
    expect(
        errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair)),
        tag,
    ).toBe(true);
};

describe("stored names come from the archive JSZip loads (#462)", { timeout: 300_000 }, () => {
    it.each([0, 70_000])(
        "does not refuse the package for a name only a prepended archive stores, behind %i trailing bytes",
        async (count) => {
            await withTempDir(async (dir) => {
                const bytes = withTrailing(prefixed([{ name: "../evil.xml", content: "x" }], await buildDocx()), count);
                const file = path.join(dir, "prefixed.docx");
                writeFileSync(file, bytes);
                for (const profile of PROFILES) {
                    const tag = `${count} ${profile}`;
                    const full = await validate(file, { profile });
                    expect(withCode(full.issues, EXTRACT_FAILED), tag).toEqual([]);
                    expect(
                        full.issues.filter((issue) => issue.message.includes("../evil.xml")),
                        tag,
                    ).toEqual([]);
                    // The loaded archive is the minimal package; lenient runs no
                    // ZIP-envelope check, so nothing else is wrong with it.
                    if (profile === "lenient") expect(full.valid, tag).toBe(true);
                    if (profile === "strict") continue;
                    for (const selection of [{ checks: [EQUIVALENT] }, { checks: ["package"] }, { quick: true }] as const) {
                        const partialTag = `${tag} ${JSON.stringify(selection)}`;
                        const partial = await validate(file, { profile, ...selection });
                        expect(withCode(partial.issues, EXTRACT_FAILED), partialTag).toEqual([]);
                        expectParity(full, partial, partialTag);
                    }
                }
                expect(readFileSync(file)).toEqual(bytes);
            });
        },
    );

    it.each([0, 70_000])(
        "reports an escaped slash the loaded archive stores behind a clean prepended archive, behind %i trailing bytes",
        async (count) => {
            await withTempDir(async (dir) => {
                const live = await buildDocx({ ...MINIMAL_PARTS, "word%2Fxtra.xml": "<x/>" });
                const bytes = withTrailing(prefixed([{ name: "ok.xml", content: "x" }], live), count);
                const file = path.join(dir, "prefixed.docx");
                writeFileSync(file, bytes);
                for (const profile of PROFILES) {
                    const tag = `${count} ${profile}`;
                    const full = await validate(file, { profile });
                    expect(withCode(full.issues, ESCAPED), tag).toEqual([
                        expect.objectContaining({ code: ESCAPED, severity: profile === "lenient" ? "warning" : "error" }),
                    ]);
                    if (profile === "strict") continue;
                    for (const selection of [{ checks: [ESCAPED] }, { checks: ["package"] }] as const) {
                        const partialTag = `${tag} ${JSON.stringify(selection)}`;
                        const partial = await validate(file, { profile, ...selection });
                        expect(withCode(partial.issues, ESCAPED), partialTag).toEqual(withCode(full.issues, ESCAPED));
                        expect(partial.valid, partialTag).toBe(profile === "lenient");
                        expectParity(full, partial, partialTag);
                    }
                }
                expect(readFileSync(file)).toEqual(bytes);
            });
        },
    );

    it.each([0, 70_000])("does not report a duplicate only a prepended archive stores, behind %i trailing bytes", async (count) => {
        await withTempDir(async (dir) => {
            const duplicates = [
                { name: DOC, content: "<a/>" },
                { name: DOC, content: "<b/>" },
            ];
            const bytes = withTrailing(prefixed(duplicates, await buildDocx()), count);
            const file = path.join(dir, "prefixed.docx");
            writeFileSync(file, bytes);
            for (const profile of PROFILES) {
                const tag = `${count} ${profile}`;
                const full = await validate(file, { profile });
                expect(withCode(full.issues, EQUIVALENT), tag).toEqual([]);
                if (profile === "lenient") expect(full.valid, tag).toBe(true);
                if (profile === "strict") continue;
                for (const selection of [{ checks: [EQUIVALENT] }, { checks: ["package"] }, { quick: true }] as const) {
                    const partialTag = `${tag} ${JSON.stringify(selection)}`;
                    const partial = await validate(file, { profile, ...selection });
                    expect(withCode(partial.issues, EQUIVALENT), partialTag).toEqual([]);
                    expectParity(full, partial, partialTag);
                }
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    // Behind a fake signature JSZip loads an empty archive: the fake is the last
    // signature, and its zero offsets rebase onto itself. The stored names are
    // that archive's — none — so the duplicate behind it is not reported, and
    // the staged run reports the missing parts instead.
    it("reports the empty archive JSZip loads behind a fake EOCD signature, not the duplicate it shadows", async () => {
        await withTempDir(async (dir) => {
            const duplicate = buildRawZip([
                ...Object.entries(MINIMAL_PARTS)
                    .filter(([part]) => part !== DOC)
                    .map(([part, content]) => ({ name: part, content })),
                { name: DOC, content: MINIMAL_PARTS[DOC]! },
                { name: DOC, content: MINIMAL_PARTS[DOC]! },
            ]);
            const bytes = withTrailing(duplicate, 70_000, fakeEocd());
            const file = path.join(dir, "fake.docx");
            writeFileSync(file, bytes);
            for (const profile of PROFILES) {
                const full = await validate(file, { profile });
                expect(withCode(full.issues, EQUIVALENT), profile).toEqual([]);
                expect(withCode(full.issues, "ct-missing"), profile).toHaveLength(1);
                expect(full.valid, profile).toBe(false);
                if (profile === "strict") continue;
                for (const selection of [{ checks: [EQUIVALENT] }, { checks: ["package"] }, { quick: true }] as const) {
                    const partialTag = `${profile} ${JSON.stringify(selection)}`;
                    const partial = await validate(file, { profile, ...selection });
                    expect(withCode(partial.issues, EQUIVALENT), partialTag).toEqual([]);
                    expectParity(full, partial, partialTag);
                }
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});
