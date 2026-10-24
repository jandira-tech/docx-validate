/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile, ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const ID = "opc-equivalent-part-names";
const AFTER_EOCD = "zip-bytes-after-eocd";
const DOC = "word/document.xml";
const OTHER_DOCUMENT = MINIMAL_PARTS[DOC]!.replace(">Hello<", ">Other<");
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const severityUnder = (profile: Profile) => (profile === "lenient" ? "warning" : "error");
const identical = (severity: string) => ({
    code: ID,
    severity,
    path: DOC,
    message: `OPC [M1.12] forbids equivalent stored names: ${JSON.stringify(DOC)}, ${JSON.stringify(DOC)}. The last stored entry is the one read.`,
});

// The minimal package with word/document.xml stored twice, as raw ZIP bytes:
// JSZip cannot write two entries under one name (#411).
const duplicate = (): Buffer =>
    buildRawZip([
        ...Object.entries(MINIMAL_PARTS)
            .filter(([part]) => part !== DOC)
            .map(([part, content]) => ({ name: part, content })),
        { name: DOC, content: MINIMAL_PARTS[DOC]! },
        { name: DOC, content: OTHER_DOCUMENT },
    ]);
const withTrailing = (archive: Buffer, count: number, fake?: Buffer): Buffer =>
    Buffer.concat(fake ? [archive, fake, Buffer.alloc(count - fake.length)] : [archive, Buffer.alloc(count)]);
// An EOCD signature with no readable central directory behind it.
const fakeEocd = (): Buffer => Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(18)]);

// `lastStructuralEocdOffset` scans the whole buffer, as JSZip does, so the
// stored names stay readable however much undeclared data follows the record
// (#447). 100 trailing bytes sit inside the old 64 KiB comment window; 70 000
// do not. Behind a fake signature the real record is still found, because the
// structural check rejects the fake. The garbage fake leaves JSZip holding an
// empty archive (it trusts the last record it finds), so the staged run also
// reports the missing parts.
const KINDS = [
    { label: "no trailing bytes", bytes: duplicate(), trailing: 0, lenientValid: true },
    { label: "100 trailing bytes", bytes: withTrailing(duplicate(), 100), trailing: 100, lenientValid: true },
    { label: "70000 trailing bytes", bytes: withTrailing(duplicate(), 70_000), trailing: 70_000, lenientValid: true },
    {
        label: "70000 trailing bytes behind a fake EOCD signature",
        bytes: withTrailing(duplicate(), 70_000, fakeEocd()),
        trailing: 70_000,
        lenientValid: false,
    },
] as const;

describe("identical stored names behind trailing data (#447)", { timeout: 300_000 }, () => {
    it.each(KINDS)("reports the duplicate behind $label under every profile", async (kind) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "duplicate.docx");
            writeFileSync(file, kind.bytes);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const tag = `${kind.label} ${profile}`;
                const full = await validate(file, { profile });
                expect(findings(full.issues), tag).toEqual([identical(severityUnder(profile))]);
                expect(full.valid, tag).toBe(profile === "lenient" ? kind.lenientValid : false);
                expect(full.metadata?.executedChecks, tag).toContain(ID);
                const afterEocd = full.issues.filter((issue) => issue.code === AFTER_EOCD);
                if (profile === "lenient") {
                    // Lenient runs no ZIP-envelope check.
                    expect(afterEocd, tag).toEqual([]);
                    expect(full.metadata?.executedChecks, tag).not.toContain(AFTER_EOCD);
                } else if (kind.trailing === 0) {
                    // The check ran and found nothing to follow the record.
                    expect(afterEocd, tag).toEqual([]);
                    expect(full.metadata?.executedChecks, tag).toContain(AFTER_EOCD);
                } else {
                    expect(afterEocd, tag).toEqual([expect.objectContaining({ severity: "error", code: AFTER_EOCD })]);
                    expect(afterEocd[0]?.message, tag).toContain(`${kind.trailing} undeclared byte(s)`);
                    expect(full.metadata?.executedChecks, tag).toContain(AFTER_EOCD);
                }
                if (profile === "strict") continue;
                for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }] as const) {
                    const partialTag = `${tag} ${JSON.stringify(selection)}`;
                    const partial = await validate(file, { profile, ...selection });
                    expect(findings(partial.issues), partialTag).toEqual(findings(full.issues));
                    if (full.valid) expect(partial.valid, partialTag).toBe(true); // valid full => valid partial (#381)
                    expect(
                        errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair)),
                        partialTag,
                    ).toBe(true); // partial errors are a subset of full errors (#381)
                }
            }
            expect(readFileSync(file)).toEqual(kind.bytes);
        });
    });

    it("reports the duplicate behind 70000 trailing bytes in a .docm under word-valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "duplicate.docm");
            const bytes = withTrailing(duplicate(), 70_000);
            writeFileSync(file, bytes);
            const full = await validate(file);
            expect(findings(full.issues)).toEqual([identical("error")]);
            expect(full.issues.filter((issue) => issue.code === AFTER_EOCD)).toHaveLength(1);
            expect(full.valid).toBe(false);
            for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }] as const) {
                const partial = await validate(file, selection);
                expect(findings(partial.issues), JSON.stringify(selection)).toEqual([identical("error")]);
                expect(errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair))).toBe(true);
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    // A later record that still points at the real central directory is
    // structural, so the stored names are readable — but JSZip rebases on it
    // and refuses the archive, so every profile reports the open failure and
    // no stored-name finding is invented (#447).
    it("reports the open failure, not a duplicate, when a later structural record hides the directory", async () => {
        await withTempDir(async (dir) => {
            const archive = duplicate();
            const file = path.join(dir, "duplicate.docx");
            writeFileSync(file, withTrailing(archive, 70_000, archive.subarray(archive.length - 22)));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const tag = `structural ${profile}`;
                const full = await validate(file, { profile });
                expect(
                    full.issues.map((issue) => issue.code),
                    tag,
                ).toContain("package-open-failed");
                expect(findings(full.issues), tag).toEqual([]);
                expect(full.valid, tag).toBe(false);
                expect(full.issues.filter((issue) => issue.code === AFTER_EOCD)).toHaveLength(profile === "lenient" ? 0 : 1);
                if (profile === "strict") continue;
                for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }] as const) {
                    const partial = await validate(file, { profile, ...selection });
                    const partialTag = `${tag} ${JSON.stringify(selection)}`;
                    expect(
                        partial.issues.map((issue) => issue.code),
                        partialTag,
                    ).toContain("package-open-failed");
                    expect(findings(partial.issues), partialTag).toEqual([]);
                    expect(
                        errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair)),
                        partialTag,
                    ).toBe(true);
                }
            }
            expect(readFileSync(file)).toEqual(withTrailing(archive, 70_000, archive.subarray(archive.length - 22)));
        });
    });
});
