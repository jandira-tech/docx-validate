/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { Validate } from "../src/validate";
import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";
import { buildRawZip } from "./support/raw-zip";

const ID = "opc-part-name-word-refused";
const REFUSED = [
    "word/a%21b.xml",
    ...["27", "28", "29", "2A", "2a"].map((hex) => `word/a%${hex}b.xml`),
    "word/caf%C3%A9.xml",
    "word/a%80b.xml",
    "caf%E9.xml",
    "café.xml",
    "word/café.xml",
];
const CLEAN = [
    "word/a%26b.xml",
    "word/a%3Ab.xml",
    "word/a%5Bb.xml",
    "word/a%7Fb.xml",
    "word/a!b.xml",
    "caf%20.xml",
    "cafe.xml",
    "word/a%2521b.xml",
    "word/a%41b.xml",
    "word/a%GGb.xml",
    "word/a%2.xml",
];
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errors = (issues: readonly ValidationIssue[]) => issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`);

describe("Word-refused stored names: deterministic check", () => {
    it("registers a stored-names-only package check for all formats", () => {
        expect(createDefaultCheckRegistry().get(ID)).toMatchObject({
            groups: ["package"],
            formats: ["docx", "pptx", "xlsx"],
            storedNamesOnly: true,
        });
    });
    it.each(REFUSED)("reports %s once without reading parts", async (name) => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        const ctx = ValidationContext.fromEntries([
            [name, "unused"],
            [name, "duplicate"],
        ]);
        const issues = await check!.run(ctx);
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ code: ID, path: name });
        expect(issues[0]!.message).toMatch(/Word 16\.115.*recovery/);
        expect(isWordBlockingIssue(issues[0]!, new Set())).toBe(true);
        expect(ctx.readCount(name)).toBe(0);
    });
    it("does not flag clean, unreserved or double-escaped controls", async () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        expect(await check!.run(ValidationContext.fromEntries(Object.fromEntries(CLEAN.map((n) => [n, "unused"]))))).toEqual([]);
    });
});

describe.each(["packed", "directory"] as const)("Word-refused stored names: %s integration", { timeout: 120_000 }, (input) => {
    it("keeps clean controls valid across full and partial runs, including lowercase hex in isolation", async () => {
        await withTempDir(async (dir) => {
            for (const name of [...CLEAN, "word/a%2ab.xml"]) {
                const parts = { ...MINIMAL_PARTS, [name]: "<x/>" };
                const target = path.join(dir, `case-${CLEAN.indexOf(name)}.docx`);
                if (input === "packed") writeFileSync(target, await buildDocx(parts));
                else
                    for (const [stored, text] of Object.entries(parts)) {
                        mkdirSync(path.dirname(path.join(target, stored)), { recursive: true });
                        writeFileSync(path.join(target, stored), text);
                    }
                const full = await validate(target);
                expect(findings(full.issues)).toHaveLength(name === "word/a%2ab.xml" ? 1 : 0);
                if (!name.includes("%41")) expect(full.valid).toBe(name !== "word/a%2ab.xml");
                for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }]) {
                    const partial = await validate(target, selection);
                    expect(findings(partial.issues)).toEqual(findings(full.issues));
                    if (full.valid) expect(partial.valid).toBe(true);
                    expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                }
            }
        });
    });

    it.each(["docx", "docm", "pptx", "xlsx"] as const)("profile and selection parity for %s, preserving input", async (extension) => {
        if (extension === "xlsx") expect(createDefaultCheckRegistry().get(ID)).toBeDefined();
        await withTempDir(async (dir) => {
            // A case-insensitive filesystem aliases %2A and %2a. The lowercase
            // spelling has its own directory specimen below, rather than overwriting.
            const refused = input === "directory" ? REFUSED.filter((name) => name !== "word/a%2ab.xml") : REFUSED;
            const parts = {
                ...(extension === "pptx" ? MINIMAL_PPTX_PARTS : MINIMAL_PARTS),
                ...Object.fromEntries([...refused, ...CLEAN].map((n) => [n, "<x/>"])),
            };
            const target = path.join(dir, `names.${extension}`);
            const bytes = await buildDocx(parts);
            const snapshot = Buffer.from(bytes);
            if (input === "packed") writeFileSync(target, bytes);
            else
                for (const [name, text] of Object.entries(parts)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const selections = [{ checks: [ID] }, { checks: ["package"] }, { quick: true }];
                if (profile === "strict")
                    for (const selection of selections)
                        await expect(validate(target, { profile, ...selection })).rejects.toBeInstanceOf(ModeConflictError);
                if (extension === "xlsx" && profile === "strict") continue;
                const full = await validate(target, { profile, ...(extension === "xlsx" ? { checks: [ID] } : {}) });
                expect(findings(full.issues), `${extension} ${profile}`).toHaveLength(refused.length);
                const severity = extension === "docx" || extension === "docm" ? (profile === "lenient" ? "warning" : "error") : "warning";
                expect(findings(full.issues).every((i) => i.severity === severity)).toBe(true);
                expect(
                    findings(full.issues)
                        .map((i) => i.path)
                        .sort(),
                ).toEqual([...refused].sort());
                if (profile !== "strict")
                    for (const selection of selections) {
                        const partial = await validate(target, { profile, ...selection });
                        expect(findings(partial.issues)).toEqual(findings(full.issues));
                        expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                        if (full.valid) expect(partial.valid).toBe(true);
                    }
            }
            expect(bytes).toEqual(snapshot);
            if (input === "packed") expect(readFileSync(target)).toEqual(snapshot);
            else for (const [name, text] of Object.entries(parts)) expect(readFileSync(path.join(target, name), "utf8")).toBe(text);
        });
    });
});

describe("Word-refused raw ZIP names integration", { timeout: 120_000 }, () => {
    it("preserves caller-owned byte API inputs while reporting both name defects", async () => {
        const bytes = await buildDocx({
            ...MINIMAL_PARTS,
            "word/a%21b.xml": "<x/>",
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/a%80b.xml" ContentType="application/xml"/></Types>',
            ),
        });
        const before = Buffer.from(bytes);
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            const result = await new Validate({ profile }).run(bytes);
            expect(findings(result.issues)).toHaveLength(1);
            expect(result.issues.filter((issue) => issue.code === "ct-override-partname-word-refused")).toHaveLength(1);
            expect(bytes).toEqual(before);
        }
    });

    it.each([
        { name: "caf\x82.xml", nameBytes: Buffer.from("caf\x82.xml", "latin1"), utf8: false },
        { name: "café.xml", utf8: true },
        { name: "café.xml", utf8: false },
    ])("reports raw bytes $name (UTF-8 flag $utf8)", async (entry) => {
        await withTempDir(async (dir) => {
            const bytes = buildRawZip([
                ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                { ...entry, content: "<x/>" },
            ]);
            const before = Buffer.from(bytes);
            const target = path.join(dir, "raw.docx");
            writeFileSync(target, bytes);
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const result = await validate(target, { profile });
                expect(findings(result.issues)).toHaveLength(1);
                expect(findings(result.issues)[0]!.severity).toBe(profile === "lenient" ? "warning" : "error");
            }
            expect(bytes).toEqual(before);
            expect(readFileSync(target)).toEqual(before);
        });
    });
    it("diagnoses stored names even when the package view refuses another entry (#420)", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "refused.docx");
            writeFileSync(
                target,
                buildRawZip([
                    ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                    { name: "word/a%21b.xml", content: "<x/>" },
                    { name: "word/a/../b.xml", content: "<x/>" },
                ]),
            );
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const full = await validate(target, { profile });
                expect(findings(full.issues)).toHaveLength(1);
                expect(full.issues.some((i) => i.code === "package-extract-failed")).toBe(true);
                if (profile !== "strict")
                    for (const selection of [{ checks: [ID] }, { checks: ["package"] }, { quick: true }])
                        expect(findings((await validate(target, { profile, ...selection })).issues)).toEqual(findings(full.issues));
            }
        });
    });
});

describe("Word-refused stored names: message and entry scope (#473, #474)", { timeout: 120_000 }, () => {
    it("scopes the recovery claim to a DOCX package for a PPTX finding", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "scope.pptx");
            writeFileSync(target, await buildDocx({ ...MINIMAL_PPTX_PARTS, "ppt/a%21b.xml": "<x/>" }));
            const found = findings((await validate(target, { profile: "strict" })).issues);
            expect(found).toHaveLength(1);
            expect(found[0]!.message).toContain("Word 16.115 offers recovery for this spelling in a DOCX package");
            expect(found[0]!.message).not.toMatch(/offers recovery\.$/);
        });
    });

    it("keeps naming the Word 16.115 recovery observation for a DOCX finding", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "scope.docx");
            writeFileSync(target, await buildDocx({ ...MINIMAL_PARTS, "word/a%21b.xml": "<x/>" }));
            const found = findings((await validate(target, { profile: "strict" })).issues);
            expect(found).toHaveLength(1);
            expect(found[0]!.message).toMatch(/Word 16\.115 offers recovery/);
        });
    });

    it.each([
        { label: "trailing slash and MS-DOS directory attribute", name: "word/caf%C3%A9/", directory: true },
        { label: "trailing slash only", name: "word/caf%C3%A9/", directory: false },
        { label: "MS-DOS directory attribute only", name: "word/caf%C3%A9", directory: true },
    ])("does not judge the folder entry ($label); the file entry is the control (#474)", async (variant) => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "folder-entry.docx");
            writeFileSync(
                target,
                buildRawZip([
                    ...Object.entries(MINIMAL_PARTS).map(([name, content]) => ({ name, content })),
                    { name: variant.name, content: "", directory: variant.directory },
                    { name: "word/xtra/a%21b.xml", content: "<x/>" },
                ]),
            );
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const found = findings((await validate(target, { profile })).issues);
                expect(found, `${variant.label} ${profile}`).toHaveLength(1);
                expect(found[0]!.path, `${variant.label} ${profile}`).toBe("word/xtra/a%21b.xml");
            }
        });
    });
});
