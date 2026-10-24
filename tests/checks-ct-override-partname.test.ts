/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `ct-override-partname-invalid` check (#414): an `Override` whose
 * `PartName` is not a valid OPC part name names no part. Word 16.115 offers
 * to recover such a package (probe sessions 3c, 4, 5 and 7), even when nothing
 * is stored under that name and the content type is well formed, while a
 * well-formed `PartName` for a missing part opens clean (#401).
 *
 * The check reports one finding per offending `Override`, at the stored
 * content-types stream. It is an error under strict and Word-blocking under
 * word-valid for DOCX/DOCM; lenient, and the unobserved PPTX/XLSX, warn. The
 * one exception is a folder segment ending with `.`, which Word opens clean
 * (session 7, #450): word-valid warns on it, and strict still rejects it.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const ID = "ct-override-partname-invalid";

const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);

/** A malformed `PartName` and the rule it breaks. */
interface Specimen {
    readonly label: string;
    readonly partName: string;
    readonly rule: string;
    /** libxml2 refuses the PartName as an `xs:anyURI`, so lenient's schema walk fails the package too. */
    readonly anyUriRefused?: boolean;
}

const SPECIMENS: readonly Specimen[] = [
    { label: "dot segment", partName: "/word/../word/nothere.xml", rule: "must not have a '.' or '..' segment" },
    { label: "query", partName: "/word/nothere.xml?x=1", rule: "must not contain '?', '#' or '\\'" },
    { label: "fragment", partName: "/word/nothere.xml#f", rule: "must not contain '?', '#' or '\\'" },
    { label: "backslash", partName: "/word\\nothere.xml", rule: "must not contain '?', '#' or '\\'" },
    { label: "no leading slash", partName: "word/nothere.xml", rule: "must start with '/'" },
    { label: "empty segment", partName: "/word//nothere.xml", rule: "must not have an empty segment" },
    { label: "trailing slash", partName: "/word/nothere.xml/", rule: "must not end with '/'" },
    { label: "escaped slash", partName: "/word/no%2Fthere.xml", rule: "must not percent-encode '/' or '\\' (OPC [M1.7])" },
    {
        label: "escaped unreserved",
        partName: "/word/n%6Fthere.xml",
        rule: "must not percent-encode an unreserved character (OPC [M1.8])",
    },
    // Session 7 (#450): a final segment ending with `.`, and characters
    // outside the RFC 3986 `pchar` set, make Word offer recovery too.
    { label: "trailing dot", partName: "/word/nothere.xml.", rule: "must not end with '.'" },
    { label: "space", partName: "/word/no there.xml", rule: "must contain only RFC 3986 pchar characters" },
    { label: "bracket", partName: "/word/no[there].xml", rule: "must contain only RFC 3986 pchar characters", anyUriRefused: true },
    { label: "non-ASCII", partName: "/word/café.xml", rule: "must contain only RFC 3986 pchar characters" },
    // Session 12 (#453): a segment made only of dots makes Word offer
    // recovery, even as a folder segment, unlike `/word./`.
    { label: "three-dot folder segment", partName: "/word/.../nothere.xml", rule: "must not have a dot-only segment (OPC [M1.10])" },
    { label: "four-dot folder segment", partName: "/word/..../nothere.xml", rule: "must not have a dot-only segment (OPC [M1.10])" },
    { label: "three-dot first segment", partName: "/.../nothere.xml", rule: "must not have a dot-only segment (OPC [M1.10])" },
];

/**
 * PartNames whose only defect is a folder segment ending with `.`. ECMA-376
 * Part 2 [M1.9] forbids it, but Word 16.115 opens each of these clean
 * (session 7, #450).
 */
const FOLDER_DOT_PART_NAMES = ["/word./nothere.xml", "/word/sub./nothere.xml", "/x./nothere.xml"] as const;
const FOLDER_DOT_RULE = "must not have a folder segment ending with '.'";

/** The content types the specimens are typed with: well formed, and the Word-rejected `invalid-*`. */
const CONTENT_TYPES = ["application/xml", "invalid-x"] as const;

const withOverride = (partName: string, contentType: string): Readonly<Record<string, string>> => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`,
    ),
});

const expected = (partName: string, rule: string, severity: "error" | "warning") => ({
    code: ID,
    severity,
    path: "[Content_Types].xml",
    message: `Override PartName ${JSON.stringify(partName)} is not a valid part name: it ${rule}.`,
});

const diagnose = async (ctx: ValidationContext): Promise<readonly ValidationIssue[]> => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check, "the check is registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};

describe("ct-override-partname-invalid", () => {
    it("registers in the content-types group for every package format", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["content-types"], formats: ["docx", "pptx", "xlsx"] });
        // It reads a part's DOM, so the XML syntax gate must precede it.
        expect(check?.requires ?? []).toContain("xml-syntax");
    });

    it.each(SPECIMENS)("reports the $label PartName once, naming the rule it breaks", async ({ partName, rule }) => {
        const ctx = ValidationContext.fromEntries(withOverride(partName, "application/xml"));
        expect(await diagnose(ctx)).toEqual([expected(partName, rule, "error")]);
    });

    it("reports every offending Override, in document order", async () => {
        const ctx = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/../a.xml" ContentType="application/xml"/>` +
                    `<Override PartName="/word/b.xml" ContentType="application/xml"/>` +
                    `<Override PartName="/word/c.xml?x=1" ContentType="application/xml"/>` +
                    `</Types>`,
            ),
        });
        expect(await diagnose(ctx)).toEqual([
            expected("/word/../a.xml", "must not have a '.' or '..' segment", "error"),
            expected("/word/c.xml?x=1", "must not contain '?', '#' or '\\'", "error"),
        ]);
    });

    it("accepts a well-formed PartName for a part that does not exist", async () => {
        const ctx = ValidationContext.fromEntries(withOverride("/word/nothere.xml", "application/xml"));
        expect(await diagnose(ctx)).toEqual([]);
    });

    it("reports nothing from a content-types stream that does not parse", async () => {
        const ctx = ValidationContext.fromEntries({ "[Content_Types].xml": "<Types><Override" });
        expect(await diagnose(ctx)).toEqual([]);
    });

    it("reports at the stored stream path, whatever its spelling", async () => {
        const ctx = ValidationContext.fromEntries({
            "[CONTENT_TYPES].XML": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/../a.xml" ContentType="application/xml"/></Types>`,
            ),
        });
        expect(await diagnose(ctx)).toEqual([
            { ...expected("/word/../a.xml", "must not have a '.' or '..' segment", "error"), path: "[CONTENT_TYPES].XML" },
        ]);
    });

    it("is Word-blocking for DOCX and DOCM", () => {
        expect(isWordBlockingIssue({ code: ID, severity: "error", message: "", path: "[Content_Types].xml" }, new Set())).toBe(true);
    });

    it.each(FOLDER_DOT_PART_NAMES)("reports the folder segment ending with '.' in %s", async (partName) => {
        const ctx = ValidationContext.fromEntries(withOverride(partName, "application/xml"));
        expect(await diagnose(ctx)).toEqual([expected(partName, FOLDER_DOT_RULE, "error")]);
    });

    it("reports a Word-blocking rule in preference to a folder segment ending with '.'", async () => {
        const ctx = ValidationContext.fromEntries(withOverride("/word./nothere.xml?x=1", "application/xml"));
        expect(await diagnose(ctx)).toEqual([expected("/word./nothere.xml?x=1", "must not contain '?', '#' or '\\'", "error")]);
    });

    it("is not Word-blocking for a folder segment ending with '.', which Word opens clean (#450)", async () => {
        const [folderDot] = await diagnose(ValidationContext.fromEntries(withOverride("/word./nothere.xml", "application/xml")));
        const [trailingDot] = await diagnose(ValidationContext.fromEntries(withOverride("/word/nothere.xml.", "application/xml")));
        expect(folderDot).toBeDefined();
        expect(trailingDot).toBeDefined();
        expect(isWordBlockingIssue(folderDot!, new Set())).toBe(false);
        expect(isWordBlockingIssue(trailingDot!, new Set())).toBe(true);
    });

    it("judges a finding by the rule it breaks, not by the PartName it quotes", async () => {
        // The quoted PartName spells the tolerated rule, but the space breaks pchar.
        const partName = `/word/a: it ${FOLDER_DOT_RULE}.xml`;
        const issues = await diagnose(ValidationContext.fromEntries(withOverride(partName, "application/xml")));
        expect(issues).toEqual([expected(partName, "must contain only RFC 3986 pchar characters", "error")]);
        expect(isWordBlockingIssue(issues[0]!, new Set())).toBe(true);
    });
});

describe("ct-override-partname-invalid in the unobserved formats", { timeout: 120_000 }, () => {
    it("warns for PPTX under word-valid and lenient, and errors under strict", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.pptx");
            writeFileSync(
                file,
                await buildPptx({
                    ...MINIMAL_PPTX_PARTS,
                    "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
                        "</Types>",
                        `<Override PartName="/ppt/../ppt/nothere.xml" ContentType="application/xml"/></Types>`,
                    ),
                }),
            );
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(findings(result.issues), profile).toEqual([
                    expected("/ppt/../ppt/nothere.xml", "must not have a '.' or '..' segment", profile === "strict" ? "error" : "warning"),
                ]);
            }
        });
    });

    it("warns for XLSX in a partial run, as no full XLSX validation exists", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.xlsx");
            writeFileSync(file, await buildDocx(withOverride("/word/../word/nothere.xml", "application/xml")));
            for (const profile of ["word-valid", "lenient"] as const) {
                const result = await validate(file, { profile, checks: [ID] });
                expect(findings(result.issues), profile).toEqual([
                    expected("/word/../word/nothere.xml", "must not have a '.' or '..' segment", "warning"),
                ]);
            }
        });
    });
});

describe.each(["docx", "docm"] as const)("ct-override-partname-invalid in a %s package", { timeout: 120_000 }, (extension) => {
    const specimen = async (dir: string, label: string, parts: Readonly<Record<string, string>>): Promise<string> => {
        const file = `${dir}/${label.replace(/[^A-Za-z0-9.-]/g, "_")}.${extension}`;
        writeFileSync(file, await buildDocx(parts));
        return file;
    };

    it.each(SPECIMENS)("reports the $label PartName under every profile", async ({ partName, rule, anyUriRefused }) => {
        for (const contentType of CONTENT_TYPES) {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `specimen-${contentType}`, withOverride(partName, contentType));
                // The `invalid-x` type is refused by the schema, so only the
                // well-formed type leaves the PartName as the sole defect,
                // unless the schema refuses the PartName itself.
                const lenientValid = contentType === "application/xml" && anyUriRefused !== true;
                for (const profile of ["word-valid", "lenient", "strict"] as const) {
                    const result = await validate(file, { profile });
                    const severity = profile === "lenient" ? "warning" : "error";
                    expect(findings(result.issues), `${profile} ${contentType}`).toEqual([expected(partName, rule, severity)]);
                    expect(result.metadata?.executedChecks, profile).toContain(ID);
                    // Word offers recovery for every malformed PartName
                    // probed, so word-valid fails the package (#414).
                    expect(result.valid, `${profile} ${contentType}`).toBe(profile === "lenient" ? lenientValid : false);
                }
            });
        }
    });

    it("leaves the plain PartName and the minimal package valid", async () => {
        await withTempDir(async (dir) => {
            const control = await specimen(dir, "control", withOverride("/word/nothere.xml", "application/xml"));
            const minimal = await specimen(dir, "minimal", MINIMAL_PARTS);
            for (const file of [control, minimal]) {
                for (const profile of ["word-valid", "lenient", "strict"] as const) {
                    const result = await validate(file, { profile });
                    expect(findings(result.issues), `${file} ${profile}`).toEqual([]);
                    expect(result.valid, `${file} ${profile}`).toBe(true);
                }
            }
        });
    });

    it("agrees between the full, quick and named runs", async () => {
        for (const contentType of CONTENT_TYPES) {
            await withTempDir(async (dir) => {
                const file = await specimen(dir, `agreement-${contentType}`, withOverride("/word/../word/nothere.xml", contentType));
                for (const profile of ["word-valid", "lenient"] as const) {
                    const full = await validate(file, { profile });
                    const severity = profile === "lenient" ? "warning" : "error";
                    const expectedFinding = expected("/word/../word/nothere.xml", "must not have a '.' or '..' segment", severity);
                    for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${profile} ${JSON.stringify(selection)}`;
                        expect(findings(partial.issues), tag).toEqual([expectedFinding]);
                        // #381: a valid full run means a valid partial run, and
                        // the partial errors are a subset of the full ones.
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                        expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }
                }
            });
        }
    });

    it.each(FOLDER_DOT_PART_NAMES)("passes word-valid and fails strict for %s, which Word opens clean (#450)", async (partName) => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "folder-dot", withOverride(partName, "application/xml"));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(findings(result.issues), profile).toEqual([
                    expected(partName, FOLDER_DOT_RULE, profile === "strict" ? "error" : "warning"),
                ]);
                // Strict holds to [M1.9]; Word 16.115 opens the package clean.
                expect(result.valid, profile).toBe(profile !== "strict");
            }
            const full = await validate(file, { profile: "word-valid" });
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(findings(partial.issues), tag).toEqual([expected(partName, FOLDER_DOT_RULE, "warning")]);
                expect(partial.valid, tag).toBe(true);
                expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
            }
        });
    });

    it("still fails word-valid when a folder segment ending with '.' comes with a Word-blocking rule", async () => {
        await withTempDir(async (dir) => {
            const partName = "/word./nothere.xml?x=1";
            const file = await specimen(dir, "folder-dot-query", withOverride(partName, "application/xml"));
            const result = await validate(file, { profile: "word-valid" });
            expect(findings(result.issues)).toEqual([expected(partName, "must not contain '?', '#' or '\\'", "error")]);
            expect(result.valid).toBe(false);
        });
    });

    it("does not resolve a malformed PartName onto a stored part", async () => {
        // "/word/../word/document.xml" resolves to the stored main document
        // part under dot-segment resolution, but a malformed part name names
        // no part: the stored part keeps its own Override (#414).
        await withTempDir(async (dir) => {
            const file = path.join(dir, "resolution.docx");
            writeFileSync(file, await buildDocx(withOverride("/word/../word/document.xml", "application/xml")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.filter((issue) => issue.code === "ct-uncovered-part" || issue.code === "word-main-document-part")).toEqual(
                [],
            );
            expect(findings(result.issues)).toEqual([
                expected("/word/../word/document.xml", "must not have a '.' or '..' segment", "error"),
            ]);
        });
    });

    it("refuses a partial run under strict, as every other profile does", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "strict", withOverride("/word/../word/nothere.xml", "application/xml"));
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toThrow(/strict/i);
            }
        });
    });
});
