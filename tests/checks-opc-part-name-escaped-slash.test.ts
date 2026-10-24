/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildRawZip } from "./support/raw-zip";

const ID = "opc-part-name-escaped-slash";
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const expected = (name: string, severity = "error") => ({
    code: ID,
    severity,
    path: name,
    message: `OPC [M1.7] forbids percent-encoded '/' or '\\' in stored name: ${JSON.stringify(name)}.`,
});
const STYLES = MINIMAL_PARTS["word/styles.xml"]!;
/** An unreferenced extra item, as in the Word probes for #436. */
const extra = (name: string) => ({ ...MINIMAL_PARTS, [name]: STYLES });
/** A referenced part whose stored name and relationship target carry the escape. */
const referenced = (escape: string): Record<string, string> => {
    const parts: Record<string, string> = {};
    for (const [part, text] of Object.entries(MINIMAL_PARTS)) {
        if (part === "word/styles.xml") parts[`word/sty${escape}les.xml`] = text;
        else if (part === "word/_rels/document.xml.rels") parts[part] = text.replace('Target="styles.xml"', `Target="sty${escape}les.xml"`);
        else parts[part] = text;
    }
    return parts;
};
const CASES = [
    { name: "word%2Fxtra.xml", parts: extra("word%2Fxtra.xml"), control: extra("word/xtra.xml") },
    { name: "word%2fxtra.xml", parts: extra("word%2fxtra.xml"), control: extra("word/xtra.xml") },
    { name: "word%5Cxtra.xml", parts: extra("word%5Cxtra.xml"), control: extra("word/xtra.xml") },
    { name: "word%5cxtra.xml", parts: extra("word%5cxtra.xml"), control: extra("word/xtra.xml") },
    { name: "word/sty%2Fles.xml", parts: referenced("%2F"), control: { ...MINIMAL_PARTS } },
];
const CONTROLS = ["%20", "%25", "%5B", "%78", "%73", "%2E", "%GG", "%2", "%g5", "%252F", "plain"].map((value) => `word/a${value}-b.xml`);
const diagnose = async (ctx: ValidationContext) => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check, "M1.7 check is registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};

// Deterministic, memory-only unit tests: names need no XML parsing or I/O.
describe("OPC M1.7 stored-name check", () => {
    it("registers for every package format with only the stored names", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["package"], formats: ["docx", "pptx", "xlsx"] });
        expect(check?.archiveOnly).not.toBe(true);
        expect(check?.storedNamesOnly).toBe(true);
        expect(check?.requires ?? []).toEqual([]);
    });

    it("reports both separator escapes in either hex case, once per name", async () => {
        for (const escape of ["%2F", "%2f", "%5C", "%5c"]) {
            const name = `word/a${escape}b.xml`;
            expect(await diagnose(ValidationContext.fromEntries({ [name]: "unused" })), escape).toEqual([expected(name)]);
        }
        // An escaped separator inside a later segment too, not just a leading one.
        const nested = "word/a%2Fb/c%5cd.xml";
        expect(await diagnose(ValidationContext.fromEntries({ [nested]: "unused" }))).toEqual([expected(nested)]);
        // Both escapes in one name still report the spelling once.
        const both = "word/a%2Fb%5Cc.xml";
        expect(await diagnose(ValidationContext.fromEntries({ [both]: "unused" }))).toEqual([expected(both)]);
    });

    it("reports each stored spelling once in code-unit order, without reading any part", async () => {
        const names = ["word/z%2Fx.xml", "WORD%5Ca.xml", "word/a%2fz.xml", "word/plain.xml", "word%5Ca.xml"];
        const ctx = ValidationContext.fromEntries([
            ...names.map((name) => [name, "unused"] as const),
            ["word%5Ca.xml", "duplicate spelling"],
        ]);
        const view = ctx.withUnparseable(new Map([["word/plain.xml", "rejected"]]));
        const issues = await diagnose(view);
        expect(issues).toEqual(
            names
                .filter((name) => name !== "word/plain.xml")
                .sort()
                .map((name) => expected(name)),
        );
        expect(issues.every((issue) => isWordBlockingIssue(issue, new Set()))).toBe(true);
        expect(ctx.partNames.every((name) => ctx.readCount(name) === 0)).toBe(true);
    });

    it("ignores other escapes, malformed escapes and plain names", async () => {
        expect(await diagnose(ValidationContext.fromEntries(Object.fromEntries(CONTROLS.map((name) => [name, "unused"]))))).toEqual([]);
    });
});

// Real filesystem/archive operations belong to integration tests only.
describe.each(["packed", "directory"] as const)("M1.7 %s integration", { timeout: 120_000 }, (input) => {
    const writePackage = async (dir: string, label: string, parts: Readonly<Record<string, string>>) => {
        const target = path.join(dir, label);
        if (input === "packed") writeFileSync(target, await buildDocx(parts));
        else
            for (const [name, text] of Object.entries(parts)) {
                mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                writeFileSync(path.join(target, name), text);
            }
        return target;
    };

    it.each(CASES)("reports $name under every profile and selection without modifying input", async ({ name, parts, control }) => {
        await withTempDir(async (dir) => {
            const file = await writePackage(dir, "escaped.docx", parts);
            const before = input === "packed" ? readFileSync(file) : undefined;
            const clean = await writePackage(dir, "control.docx", control);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                // Word 16.115 offers recovery for `word%2Fxtra.xml`, refuses
                // `word%5Cxtra.xml`, and opens the `word/xtra.xml` control
                // clean (#436; docs/plans/evidence/word-probes-2026-10-24.md).
                const controlRun = await validate(clean, { profile });
                expect(findings(full.issues), profile).toEqual([expected(name, profile === "lenient" ? "warning" : "error")]);
                expect(full.metadata?.executedChecks, profile).toContain(ID);
                if (profile === "word-valid") expect(controlRun.valid).toBe(true);
                expect(findings(controlRun.issues), profile).toEqual([]);
                expect(full.valid, profile).toBe(profile === "lenient" ? controlRun.valid : false);
                if (profile !== "strict")
                    for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }]) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(findings(partial.issues), `${profile} ${JSON.stringify(selection)}`).toEqual(findings(full.issues));
                        if (full.valid) expect(partial.valid).toBe(true);
                        expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
            }
            if (before) expect(readFileSync(file)).toEqual(before);
            else for (const [stored, text] of Object.entries(parts)) expect(readFileSync(path.join(file, stored), "utf8")).toBe(text);
        });
    });

    it("does not report the well-formed control in full or partial runs", async () => {
        await withTempDir(async (dir) => {
            const file = await writePackage(dir, "control.docx", { ...MINIMAL_PARTS, "word/xtra.xml": STYLES });
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                // Fails on an assertion in RED, before any unknown-selector error.
                expect(full.metadata?.executedChecks, profile).toContain(ID);
                expect(findings(full.issues), profile).toEqual([]);
                if (profile === "word-valid") expect(full.valid).toBe(true);
                if (profile !== "strict")
                    for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }]) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(findings(partial.issues)).toEqual([]);
                        if (full.valid) expect(partial.valid).toBe(true);
                        expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
            }
        });
    });

    it.each(["docm", "pptx", "xlsx"])("applies format policy to %s", async (extension) => {
        expect(createDefaultCheckRegistry().get(ID)).toBeDefined();
        await withTempDir(async (dir) => {
            const name = "word%2Fxtra.xml";
            const file = await writePackage(dir, `escaped.${extension}`, extra(name));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                // Full XLSX validation is unsupported; registered checks still apply.
                if (extension === "xlsx" && profile === "strict") continue;
                const result = await validate(file, { profile, ...(extension === "xlsx" ? { checks: [ID] } : {}) });
                const severity = profile === "strict" || (profile === "word-valid" && extension === "docm") ? "error" : "warning";
                expect(findings(result.issues), `${extension} ${profile}`).toEqual([expected(name, severity)]);
                if (extension === "docm" && profile === "word-valid") expect(result.valid).toBe(false);
            }
        });
    });
});

describe("an escaped name beside a name the package view refuses (#420)", { timeout: 120_000 }, () => {
    // `word/a/../b.xml` refuses the package view, but the stored names still
    // carry the escaped spelling, so the M1.7 finding survives beside the
    // extraction failure, in full and partial runs alike.
    const refusedPackage = (): Buffer =>
        buildRawZip([
            ...Object.entries(MINIMAL_PARTS).map(([part, text]) => ({ name: part, content: text })),
            { name: "word%2Fxtra.xml", content: STYLES },
            { name: "word/a/../b.xml", content: "<x/>" },
        ]);

    it("still reports the escaped name, in full and partial runs", async () => {
        const bytes = refusedPackage();
        await withTempDir(async (dir) => {
            const file = path.join(dir, "refused.docx");
            writeFileSync(file, bytes);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                expect(findings(full.issues), profile).toEqual([expected("word%2Fxtra.xml", profile === "lenient" ? "warning" : "error")]);
                expect(
                    full.issues.map((issue) => issue.code),
                    profile,
                ).toContain("package-extract-failed");
                expect(full.valid, profile).toBe(false);
            }
            for (const profile of ["word-valid", "lenient"] as const)
                for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }]) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    expect(findings(partial.issues), tag).toEqual([
                        expected("word%2Fxtra.xml", profile === "lenient" ? "warning" : "error"),
                    ]);
                    expect(
                        partial.issues.map((issue) => issue.code),
                        tag,
                    ).toContain("package-extract-failed");
                    expect(partial.metadata?.executedChecks, tag).toContain(ID);
                    expect(partial.valid, tag).toBe(false);
                }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});
