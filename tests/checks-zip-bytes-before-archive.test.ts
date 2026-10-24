/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `zip-bytes-before-archive` (#464): bytes in front of the archive JSZip
 * loads make Word 16.115 offer recovery — a whole prepended ZIP, junk with
 * relative offsets, and junk with the absolute offsets a self-extractor
 * writes (probe session 14b, hash-bound evidence in
 * /Users/arthrod/temp/T/wfv-scratch/lane-464/evidence.md; session 14 of
 * docs/plans/evidence/word-probes-2026-10-24.md). It follows
 * `zip-bytes-after-eocd`'s policy: a blocking error under word-valid and
 * strict, never run by lenient, unavailable for directories.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { absoluteOffsetsAfter, buildRawZip } from "./support/raw-zip";

const ID = "zip-bytes-before-archive";
const AFTER_EOCD = "zip-bytes-after-eocd";
const JUNK = Buffer.alloc(4096, 0x20);
const PREFIX_ZIP = () =>
    buildRawZip([
        { name: "pad.bin", content: Buffer.alloc(600) },
        { name: "ok.xml", content: "<x/>" },
    ]);
const DOCX = () => buildDocx(MINIMAL_PARTS);

// The session-14b #464 specimens: the minimal DOCX behind a prefix.
const SPECIMENS = [
    ["464-prefix-zip-docx", async () => Buffer.concat([PREFIX_ZIP(), await DOCX()]), () => PREFIX_ZIP().length],
    ["464-prefix-junk-docx-relative", async () => Buffer.concat([JUNK, await DOCX()]), () => JUNK.length],
    [
        "464-prefix-junk-docx-absolute",
        async () => Buffer.concat([JUNK, absoluteOffsetsAfter(await DOCX(), JUNK.length)]),
        () => JUNK.length,
    ],
] as const;

const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorCodes = (issues: readonly ValidationIssue[]) => [
    ...new Set(issues.filter((i) => i.severity === "error").map((i) => i.code ?? "")),
];

describe("bytes before the ZIP archive (#464)", { timeout: 300_000 }, () => {
    it("keeps the schema-clean control valid under word-valid and strict", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await DOCX());
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(true);
                expect(findings(result.issues), profile).toEqual([]);
            }
        });
    });

    it.each(SPECIMENS)("word-valid and strict report $name once, through zip-bytes-before-archive", async (_name, make, prefix) => {
        await withTempDir(async (dir) => {
            const bytes = await make();
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, bytes);
            for (const profile of ["word-valid", "strict"] as const) {
                const tag = `${profile}`;
                const result = await validate(file, { profile });
                expect(result.valid, tag).toBe(false);
                expect(errorCodes(result.issues), tag).toEqual([ID]);
                expect(
                    result.issues.some((issue) => issue.code === AFTER_EOCD),
                    `${tag}: reported once, not also as trailing data`,
                ).toBe(false);
                expect(findings(result.issues), tag).toHaveLength(1);
                expect(findings(result.issues)[0]?.message, tag).toContain(`${prefix()} byte(s) precede the ZIP archive`);
                expect(findings(result.issues)[0]?.message, tag).toContain("Word offers to recover the document");
                expect(result.metadata?.executedChecks, tag).toContain(ID);
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("reports a prefix of one byte", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "one-byte.docx");
            writeFileSync(file, Buffer.concat([Buffer.from([0x0a]), await DOCX()]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(findings(result.issues)[0]?.message).toContain("1 byte(s) precede the ZIP archive");
        });
    });

    it("reports junk whose first bytes fake a local-header signature", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "fake-sig.docx");
            const stub = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.alloc(4092)]);
            writeFileSync(file, Buffer.concat([stub, await DOCX()]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result.issues)).toEqual([ID]);
        });
    });

    it("reports a ZIP64 archive behind a prefix, and opens the unprefixed control", async () => {
        await withTempDir(async (dir) => {
            const live = buildRawZip(
                Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                { zip64: true },
            );
            const control = path.join(dir, "zip64-control.docx");
            writeFileSync(control, live);
            const clean = await validate(control, { profile: "word-valid" });
            expect(clean.valid).toBe(true);
            expect(findings(clean.issues)).toEqual([]);

            const file = path.join(dir, "zip64-prefixed.docx");
            writeFileSync(file, Buffer.concat([JUNK, live]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result.issues)).toEqual([ID]);
            expect(findings(result.issues)[0]?.message).toContain("4096 byte(s) precede the ZIP archive");
        });
    });

    it("reports a prefixed .docm under word-valid and strict", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docm");
            writeFileSync(file, Buffer.concat([JUNK, await DOCX()]));
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.valid).toBe(false);
            expect(errorCodes(wordValid.issues)).toEqual([ID]);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            expect(errorCodes(strict.issues)).toEqual([ID]);
        });
    });

    it.each(SPECIMENS)("runs no envelope check on lenient for $name", async (_name, make) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await make());
            const result = await validate(file, { profile: "lenient" });
            expect(result.valid).toBe(true);
            expect(result.issues.some((issue) => issue.code === ID || issue.code === AFTER_EOCD)).toBe(false);
            expect(result.metadata?.executedChecks).not.toContain(ID);
        });
    });

    it("reports a package with a whole archive appended after it once", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "appended.docx");
            const docx = await DOCX();
            writeFileSync(file, Buffer.concat([docx, PREFIX_ZIP()]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(findings(result.issues)).toHaveLength(1);
            expect(findings(result.issues)[0]?.message).toContain(`${docx.length} byte(s) precede the ZIP archive`);
            // The live archive (the appended copy) is not trailing data (#464).
            expect(result.issues.some((issue) => issue.code === AFTER_EOCD)).toBe(false);
        });
    });

    it("still reports genuine trailing data behind a prefix through zip-bytes-after-eocd", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "prefix-and-trailing.docx");
            const trailing = Buffer.alloc(120);
            writeFileSync(file, Buffer.concat([PREFIX_ZIP(), await DOCX(), trailing]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result.issues).sort()).toEqual([AFTER_EOCD, ID]);
            expect(findings(result.issues)[0]?.message).toContain(`${PREFIX_ZIP().length} byte(s) precede the ZIP archive`);
            const afterEocd = result.issues.find((issue) => issue.code === AFTER_EOCD);
            // Measured from the loaded archive's own record, not the prefix's.
            expect(afterEocd?.message).toContain("120 undeclared byte(s)");
        });
    });

    it("keeps reporting trailing data that ends in a fake EOCD signature", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "fake-eocd.docx");
            const fakeEocd = Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(18)]);
            const junk = Buffer.concat([Buffer.alloc(100), fakeEocd]);
            writeFileSync(file, Buffer.concat([await DOCX(), junk]));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const afterEocd = result.issues.find((issue) => issue.code === AFTER_EOCD);
            expect(afterEocd?.severity).toBe("error");
            expect(afterEocd?.message).toContain(`${junk.length} undeclared byte(s)`);
        });
    });

    it("partial selections report exactly the full run's finding (#381)", async () => {
        await withTempDir(async (dir) => {
            const bytes = Buffer.concat([JUNK, await DOCX()]);
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, bytes);
            const full = await validate(file, { profile: "word-valid" });
            expect(full.valid).toBe(false);
            const fullErrors = errorCodes(full.issues);
            for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }] as const) {
                const partial = await validate(file, { ...selection });
                const tag = JSON.stringify(selection);
                expect(findings(partial.issues), tag).toEqual(findings(full.issues));
                expect(partial.valid, tag).toBe(false);
                expect(
                    errorCodes(partial.issues).every((code) => fullErrors.includes(code)),
                    `${tag}: partial errors ⊆ full errors`,
                ).toBe(true);
            }
            // Validation never modifies its input.
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("keeps a valid full run's partial runs valid (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await DOCX());
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                expect(full.valid, profile).toBe(true);
                for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }] as const) {
                    if (profile === "lenient" && selection.checks?.[0] === ID) continue; // refused, tested below
                    const partial = await validate(file, { profile, ...selection });
                    expect(partial.valid, `${profile} ${JSON.stringify(selection)}`).toBe(true);
                    expect(findings(partial.issues)).toEqual([]);
                }
            }
        });
    });

    it("refuses a lenient selection of the check, as for the other envelope checks", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await DOCX());
            await expect(validate(file, { profile: "lenient", checks: [ID] })).rejects.toThrow(ModeConflictError);
        });
    });
});
