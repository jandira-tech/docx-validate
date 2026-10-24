/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { canonicalPartName } from "../src/lib/content-types";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile, ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { contextFromDirectory } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const ID = "opc-equivalent-part-names";
const PAIRS = [
    ["word/document.xml", "WORD/document.xml"],
    ["word/styles.xml", "word/%73tyles.xml"],
    ["word/document.xml", "word\\document.xml"],
    ["[Content_Types].xml", "[CONTENT_TYPES].XML"],
] as const;
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const duplicateParts = (a: string, b: string) => ({ ...MINIMAL_PARTS, [b]: MINIMAL_PARTS[a]! });
const diagnose = async (ctx: ValidationContext) => {
    const check = createDefaultCheckRegistry().get(ID);
    // Guard a new export through the existing registry: RED is an assertion failure.
    expect(check, "the equivalent-name check must be registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};
// Word 16.115 offers recovery for every duplicate spelling probed, so the
// diagnosis is Word-blocking for DOCX/DOCM; lenient keeps the earlier policy
// and runs no Word rule, so it only warns (#351).
const severityUnder = (profile: Profile, format: "docx" | "pptx" = "docx") =>
    profile === "strict" || (profile === "word-valid" && format === "docx") ? "error" : "warning";
const expected = (names: readonly string[], kept: string, severity = "warning") => ({
    code: ID,
    severity,
    path: kept,
    message: `OPC [M1.12] forbids equivalent stored names: ${[...names]
        .sort()
        .map((name) => JSON.stringify(name))
        .join(", ")}.`,
});

// Memory-only unit tests: no disk, network, or clock dependencies.
describe("equivalent stored names", () => {
    it("registers a package check with no XML prerequisites for every package format", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["package"], formats: ["docx", "pptx", "xlsx"] });
        expect(check?.requires ?? []).toEqual([]);
        expect(check?.archiveOnly).not.toBe(true);
    });

    it("retains names lost by normalized indexing, including in gated views", async () => {
        const ctx = ValidationContext.fromEntries([
            ["word\\document.xml", "first"],
            ["/word/document.xml", "second"],
            ["word/document.xml", "kept"],
        ]);
        expect(ctx.text("word/document.xml")).toBe("kept");
        const view = ctx.withUnparseable(new Map([["word/document.xml", "rejected"]]));
        expect(await diagnose(view)).toEqual([
            expected(["word\\document.xml", "/word/document.xml", "word/document.xml"], "word/document.xml", "error"),
        ]);
    });

    it("reports each group once, in deterministic order, without reading content", async () => {
        const ctx = ValidationContext.fromEntries({ ...duplicateParts(...PAIRS[1]), "word\\document.xml": "unused" });
        const issues = await diagnose(ctx);
        expect(issues).toEqual([expected(PAIRS[2], "word/document.xml", "error"), expected(PAIRS[1], "word/styles.xml", "error")]);
        expect(ctx.partNames.every((name) => ctx.readCount(name) === 0)).toBe(true);
        expect(issues.every((issue) => isWordBlockingIssue(issue))).toBe(true);
    });

    it("keeps the first equivalent spelling when no canonical exact spelling is stored", async () => {
        const names = ["WORD/%64OCUMENT.XML", "Word/DOCUMENT.xml"];
        expect(await diagnose(ValidationContext.fromEntries(Object.fromEntries(names.map((name) => [name, "unused"]))))).toEqual([
            expected(names, names[0]!, "error"),
        ]);
    });

    it("keeps the first content-types spelling when the usual spelling is absent", async () => {
        const names = ["[content_types].XML", "[CONTENT_TYPES].xml"];
        expect(await diagnose(ValidationContext.fromEntries(Object.fromEntries(names.map((name) => [name, "unused"]))))).toEqual([
            expected(names, [...names].sort()[0]!, "error"),
        ]);
    });
});

// Integration tests build every package in code. Case variants are ZIP-only:
// case-insensitive APFS cannot hold both file/folder spellings in a directory.
describe("equivalent-name package integration", { timeout: 120_000 }, () => {
    it.each(PAIRS)("reports packed %s + %s and leaves input bytes intact", async (a, b) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "duplicate.docx");
            const bytes = await buildDocx(duplicateParts(a, b));
            writeFileSync(file, bytes);
            const ctx = await ValidationContext.fromArchive(bytes);
            expect(await diagnose(ctx)).toEqual([expected([a, b], a, "error")]);
            const result = await validate(file, { checks: [ID] });
            expect(findings(result.issues)).toEqual([expected([a, b], a, severityUnder("word-valid"))]);
            expect(readFileSync(file)).toEqual(bytes);
            const control = path.join(dir, "control.docx");
            writeFileSync(control, await buildDocx());
            expect(findings((await validate(control, { checks: [ID] })).issues)).toEqual([]);
        });
    });

    it.each([
        ["word/a%2Fb.xml", "word/a/b.xml"],
        ["word/café.xml", "word/cafe%CC%81.xml"],
        ["%5BContent_Types%5D.xml", "[Content_Types].xml"],
    ])("does not conflate %s and %s", async (a, b) => {
        if (a !== "%5BContent_Types%5D.xml") expect(canonicalPartName(a)).not.toBe(canonicalPartName(b));
        else expect(canonicalPartName(a)).toBe(canonicalPartName(b)); // The stream is not an OPC part.
        await withTempDir(async (dir) => {
            const file = path.join(dir, "distinct.docx");
            writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, [a]: "<root/>", [b]: MINIMAL_PARTS[b] ?? "<root/>" }));
            expect(await diagnose(await ValidationContext.fromArchive(readFileSync(file)))).toEqual([]);
            expect(findings((await validate(file, { checks: [ID] })).issues)).toEqual([]);
        });
    });

    it.each([PAIRS[1], PAIRS[2]])("reports directory %s + %s without modifying it", async (a, b) => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "directory.docx");
            const parts = duplicateParts(a, b);
            const control = path.join(dir, "control.docx");
            for (const [target, entries] of [
                [root, parts],
                [control, MINIMAL_PARTS],
            ] as const) {
                for (const [name, text] of Object.entries(entries)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            }
            expect(await diagnose(contextFromDirectory(control))).toEqual([]);
            expect(findings((await validate(control)).issues)).toEqual([]);
            expect(await diagnose(contextFromDirectory(root))).toEqual([expected([a, b], a, "error")]);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(root, { profile });
                expect(findings(result.issues)).toEqual([expected([a, b], a, severityUnder(profile))]);
            }
            for (const [name, text] of Object.entries(parts)) expect(readFileSync(path.join(root, name), "utf8")).toBe(text);
        });
    });

    it.each(["word-valid", "lenient", "strict"] as const)("applies %s severity and verdict", async (profile: Profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "duplicate.docx");
            const control = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx(duplicateParts(...PAIRS[2])));
            writeFileSync(control, await buildDocx());
            const full = await validate(file, { profile });
            const clean = await validate(control, { profile });
            expect(findings(full.issues)).toEqual([expected(PAIRS[2], PAIRS[2][0], severityUnder(profile))]);
            expect(clean.valid).toBe(true);
            expect(full.valid).toBe(profile === "lenient" ? clean.valid : false);
            if (profile !== "strict") {
                const partial = await validate(file, { profile, checks: [ID] });
                expect(findings(partial.issues)).toEqual(findings(full.issues));
                expect(partial.valid).toBe(full.valid);
                if (full.valid) expect(partial.valid).toBe(true); // valid full => valid partial (#381).
                expect(errorPairs(partial.issues).every((pair) => errorPairs(full.issues).includes(pair))).toBe(true);
            }
            expect(full.metadata?.executedChecks).toContain(ID);
            const quick = profile === "strict" ? undefined : await validate(file, { profile, quick: true });
            if (quick) expect(findings(quick.issues)).toEqual(findings(full.issues));
        });
    });

    it.each(PAIRS)("reports %s + %s before extraction and regardless of ZIP order", async (a, b) => {
        await withTempDir(async (dir) => {
            const parts = duplicateParts(a, b);
            const results = [];
            for (const [order, entries] of [Object.entries(parts), Object.entries(parts).reverse()].entries()) {
                const file = path.join(dir, `order-${order}.docx`);
                const bytes = await buildDocx(Object.fromEntries(entries));
                writeFileSync(file, bytes);
                const full = await validate(file, { profile: "strict" });
                expect(findings(full.issues)).toEqual([expected([a, b], a, "error")]);
                for (const profile of ["word-valid", "lenient"] as const) {
                    const run = await validate(file, { profile });
                    const partial = await validate(file, { profile, checks: [ID] });
                    expect(findings(run.issues)).toEqual([expected([a, b], a, severityUnder(profile))]);
                    if (profile === "word-valid") expect(run.valid).toBe(false);
                    expect(findings(partial.issues)).toEqual(findings(run.issues));
                    if (run.valid) expect(partial.valid).toBe(true);
                    expect(errorPairs(partial.issues).every((pair) => errorPairs(run.issues).includes(pair))).toBe(true);
                }
                // On this host, case-only collisions still stop safe extraction (#373).
                if (a.toLowerCase() === b.toLowerCase()) expect(full.issues.map((issue) => issue.code)).toContain("package-extract-failed");
                expect(readFileSync(file)).toEqual(bytes);
                results.push(findings(full.issues));
            }
            expect(results[0]).toEqual(results[1]);
        });
    });

    it.each(["docm", "pptx"])("runs on full and partial %s inputs", async (extension) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `duplicate.${extension}`);
            writeFileSync(file, await buildDocx(duplicateParts(...PAIRS[0])));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                expect(findings((await validate(file, { profile })).issues)).toEqual([
                    expected(PAIRS[0], PAIRS[0][0], severityUnder(profile, extension === "pptx" ? "pptx" : "docx")),
                ]);
            }
            expect(findings((await validate(file, { profile: "lenient", checks: [ID] })).issues)).toEqual([
                expected(PAIRS[0], PAIRS[0][0]),
            ]);
        });
    });
});

describe("M1.8 and M1.12 diagnose independent stored-name defects (#408)", { timeout: 120_000 }, () => {
    it("reports both checks for an escaped spelling beside its decoded twin", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "both.docx");
            writeFileSync(file, await buildDocx(duplicateParts(...PAIRS[1])));
            const result = await validate(file);
            expect(result.issues).toContainEqual(
                expect.objectContaining({ code: "opc-part-name-escaped-unreserved", path: PAIRS[1][1], severity: "error" }),
            );
            expect(findings(result.issues)).toEqual([expected(PAIRS[1], PAIRS[1][0], "error")]);
        });
    });

    it("retains the M1.8 finding when a case collision prevents extraction", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "collision.docx");
            const bytes = await buildDocx({ ...duplicateParts(...PAIRS[0]), "word/%78tra.xml": "<root/>" });
            writeFileSync(file, bytes);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.issues).toContainEqual(
                    expect.objectContaining({
                        code: "opc-part-name-escaped-unreserved",
                        path: "word/%78tra.xml",
                        severity: profile === "lenient" ? "warning" : "error",
                    }),
                );
                expect(result.issues.map((issue) => issue.code)).toContain("package-extract-failed");
                expect(findings(result.issues)).toHaveLength(1);
            }
            expect(readFileSync(file)).toEqual(bytes);
        });
    });
});
