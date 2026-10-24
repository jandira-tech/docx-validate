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

const ID = "opc-part-name-escaped-unreserved";
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const expected = (name: string, severity = "error") => ({
    code: ID,
    severity,
    path: name,
    message: `OPC [M1.8] forbids percent-encoded unreserved characters in stored name: ${JSON.stringify(name)}.`,
});
const renamed = (from: string, to: string) =>
    Object.fromEntries(Object.entries(MINIMAL_PARTS).map(([name, text]) => [name === from ? to : name, text]));
const CASES = [
    { name: "word/%78tra.xml", parts: { ...MINIMAL_PARTS, "word/%78tra.xml": "<root/>" } },
    { name: "word/%73tyles.xml", parts: renamed("word/styles.xml", "word/%73tyles.xml") },
    { name: "word/%64ocument.xml", parts: renamed("word/document.xml", "word/%64ocument.xml") },
    { name: "word/_rels/document.xml.%72els", parts: renamed("word/_rels/document.xml.rels", "word/_rels/document.xml.%72els") },
    { name: "word/%73tyles.xml", parts: { ...MINIMAL_PARTS, "word/%73tyles.xml": MINIMAL_PARTS["word/styles.xml"]! } },
    ...["%6a", "%41", "%30", "%2D", "%2E", "%5F", "%7E"].map((escape) => ({
        name: `word/a${escape}b.xml`,
        parts: { ...MINIMAL_PARTS, [`word/a${escape}b.xml`]: "<root/>" },
    })),
];
const CONTROLS = ["%20", "%2F", "%25", "%5B", "%", "%6", "%GG", "%6G", "%G6", "plain"].map((value) => `word/a${value}-b.xml`);
const diagnose = async (ctx: ValidationContext) => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check, "M1.8 check is registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};

// Deterministic, memory-only unit tests: names need no XML parsing or I/O.
describe("OPC M1.8 stored-name check", () => {
    it("registers for every package format without archive or XML prerequisites", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["package"], formats: ["docx", "pptx", "xlsx"] });
        expect(check?.archiveOnly).not.toBe(true);
        expect(check?.requires ?? []).toEqual([]);
    });

    it("classifies every escaped ASCII octet, including lower-case hex, without decoding recursively", async () => {
        for (let octet = 0; octet < 256; octet += 1) {
            const name = `word/a%${octet.toString(16).padStart(2, "0")}b.xml`;
            const issues = await diagnose(ValidationContext.fromEntries({ [name]: "unused" }));
            expect(issues, name).toEqual(/^[A-Za-z0-9._~-]$/.test(String.fromCharCode(octet)) ? [expected(name)] : []);
        }
        expect(await diagnose(ValidationContext.fromEntries({ "word/%2578.xml": "unused" }))).toEqual([]);
    });

    it("reports each stored name once in code-unit order, including dropped aliases and gated views", async () => {
        const names = ["word/z%78%79.xml", "word/%73tyles.xml", "word/styles.xml", "WORD/%41.xml", "word\\a%30.xml"];
        const ctx = ValidationContext.fromEntries(Object.fromEntries(names.map((name) => [name, "unused"])));
        const view = ctx.withUnparseable(new Map([["word/styles.xml", "rejected"]]));
        const issues = await diagnose(view);
        expect(issues).toEqual(
            names
                .filter((name) => name !== "word/styles.xml")
                .sort()
                .map((name) => expected(name)),
        );
        expect(issues.every((issue) => isWordBlockingIssue(issue, new Set()))).toBe(true);
        expect(ctx.partNames.every((name) => ctx.readCount(name) === 0)).toBe(true);
        expect(
            await diagnose(ValidationContext.fromEntries(Object.fromEntries([...names].reverse().map((name) => [name, "unused"])))),
        ).toEqual(issues);
    });

    it("ignores reserved escapes, malformed escapes and plain names", async () => {
        expect(await diagnose(ValidationContext.fromEntries(Object.fromEntries(CONTROLS.map((name) => [name, "unused"]))))).toEqual([]);
    });
});

// Real filesystem/archive operations belong to integration tests only.
describe.each(["packed", "directory"] as const)("M1.8 %s integration", { timeout: 120_000 }, (input) => {
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

    it.each(CASES)("reports $name under every profile and selection without modifying input", async ({ name, parts }) => {
        await withTempDir(async (dir) => {
            const file = await writePackage(dir, "escaped.docx", parts);
            const before = input === "packed" ? readFileSync(file) : undefined;
            // Preserve unrelated defects in the plain-name control, including
            // an unreferenced extra part (which lenient already rejects).
            const controlParts = Object.fromEntries(
                Object.entries(parts).map(([stored, text]) => {
                    const decoded = decodeURIComponent(stored);
                    const plain = decoded !== stored && decoded in parts ? `${decoded}.control.xml` : decoded;
                    return [plain, text];
                }),
            );
            const control = await writePackage(dir, "control.docx", controlParts);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                const clean = await validate(control, { profile });
                expect(findings(full.issues)).toEqual([expected(name, profile === "lenient" ? "warning" : "error")]);
                expect(full.metadata?.executedChecks).toContain(ID);
                if (profile === "word-valid") expect(clean.valid).toBe(true);
                expect(full.valid).toBe(profile === "lenient" ? clean.valid : false);
                if (profile !== "strict")
                    for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }]) {
                        const partial = await validate(file, { profile, ...selection });
                        expect(findings(partial.issues)).toEqual(findings(full.issues));
                        if (full.valid) expect(partial.valid).toBe(true);
                        expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
            }
            if (before) expect(readFileSync(file)).toEqual(before);
            else for (const [stored, text] of Object.entries(parts)) expect(readFileSync(path.join(file, stored), "utf8")).toBe(text);
        });
    });

    it.each(CONTROLS)("does not report %s in full or partial runs", async (name) => {
        await withTempDir(async (dir) => {
            const file = await writePackage(dir, "control.docx", { ...MINIMAL_PARTS, [name]: "<root/>" });
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                // Fails on an assertion in RED, before any unknown-selector error.
                expect(full.metadata?.executedChecks).toContain(ID);
                expect(findings(full.issues)).toEqual([]);
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
            const name = "word/%78tra.xml";
            const file = await writePackage(dir, `escaped.${extension}`, { ...MINIMAL_PARTS, [name]: "<root/>" });
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                // Full XLSX validation is unsupported; registered checks still apply.
                if (extension === "xlsx" && profile === "strict") continue;
                const result = await validate(file, { profile, ...(extension === "xlsx" ? { checks: [ID] } : {}) });
                const severity = profile === "strict" || (profile === "word-valid" && extension === "docm") ? "error" : "warning";
                expect(findings(result.issues)).toEqual([expected(name, severity)]);
                if (extension === "docm" && profile === "word-valid") expect(result.valid).toBe(false);
            }
        });
    });
});

describe("an escaped name hidden by ZIP loading (#411)", { timeout: 120_000 }, () => {
    // Two entries stored under the identical escaped name: JSZip keeps only
    // the last one, so the stored-name list must come from the central
    // directory, not the loaded view.
    it("still reports the name both entries carry, once", async () => {
        const name = "word/%78tra.xml";
        const bytes = buildRawZip([
            ...Object.entries(MINIMAL_PARTS).map(([part, text]) => ({ name: part, content: text })),
            { name, content: "<a/>" },
            { name, content: "<b/>" },
        ]);
        const ctx = await ValidationContext.fromArchive(bytes);
        expect(ctx.storedNames.filter((stored) => stored === name)).toHaveLength(2);
        expect(await diagnose(ctx)).toEqual([expected(name)]);
        await withTempDir(async (dir) => {
            const file = path.join(dir, "escaped.docx");
            writeFileSync(file, bytes);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = await validate(file, { profile });
                expect(findings(full.issues), profile).toEqual([expected(name, profile === "lenient" ? "warning" : "error")]);
            }
        });
    });
});
