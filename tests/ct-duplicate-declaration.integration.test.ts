/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { ModeConflictError } from "../src/lib/mode-policy";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const ID = "ct-duplicate-declaration";
const STREAM = "[Content_Types].xml";
const XML_DEFAULT = '<Default Extension="xml" ContentType="application/xml"/>';
const STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CASES = [
    { name: "default-identical", extra: XML_DEFAULT, first: "xml", repeated: "xml" },
    { name: "default-case", extra: XML_DEFAULT.replace('Extension="xml"', 'Extension="XML"'), first: "xml", repeated: "XML" },
    {
        name: "override-identical",
        extra: `<Override PartName="/word/styles.xml" ContentType="${STYLES_TYPE}"/>`,
        first: "/word/styles.xml",
        repeated: "/word/styles.xml",
    },
    {
        name: "override-conflicting",
        extra: '<Override PartName="/word/styles.xml" ContentType="application/xml"/>',
        first: "/word/styles.xml",
        repeated: "/word/styles.xml",
    },
    {
        name: "override-case",
        extra: `<Override PartName="/WORD/styles.xml" ContentType="${STYLES_TYPE}"/>`,
        first: "/word/styles.xml",
        repeated: "/WORD/styles.xml",
    },
];
const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const errors = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
const SELECTIONS = [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }];

const writePackage = async (dir: string, label: string, parts: Readonly<Record<string, string>>, input: "packed" | "directory") => {
    const target = path.join(dir, label);
    if (input === "packed") writeFileSync(target, await buildDocx(parts));
    else
        for (const [name, text] of Object.entries(parts)) {
            mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
            writeFileSync(path.join(target, name), text);
        }
    return target;
};

describe.each(["packed", "directory"] as const)("duplicate declarations: %s (#431)", { timeout: 120_000 }, (input) => {
    for (const suffix of ["docx", "docm"] as const)
        for (const profile of ["word-valid", "strict", "lenient"] as const) {
            it.each(CASES)(
                `${suffix} ${profile}: $name and unchanged control, full/partial parity and input preservation`,
                async ({ extra, first, repeated }) => {
                    await withTempDir(async (dir) => {
                        const parts = { ...MINIMAL_PARTS, [STREAM]: MINIMAL_PARTS[STREAM]!.replace("</Types>", `${extra}</Types>`) };
                        const target = await writePackage(dir, `specimen.${suffix}`, parts, input);
                        const control = await writePackage(dir, `control.${suffix}`, MINIMAL_PARTS, input);
                        const before = input === "packed" ? readFileSync(target) : undefined;
                        const full = await validate(target, { profile });
                        // The RED assertion precedes named selection (which does not exist yet).
                        expect(findings(full.issues)).toEqual([
                            expect.objectContaining({ code: ID, path: STREAM, severity: profile === "lenient" ? "warning" : "error" }),
                        ]);
                        expect(findings(full.issues)[0]!.message).toContain(JSON.stringify(first));
                        expect(findings(full.issues)[0]!.message).toContain(JSON.stringify(repeated));
                        expect(full.valid).toBe(profile === "lenient");
                        expect(full.metadata?.executedChecks).toContain(ID);
                        const clean = await validate(control, { profile });
                        expect(clean.valid).toBe(true);
                        expect(findings(clean.issues)).toEqual([]);
                        for (const selection of SELECTIONS) {
                            if (profile === "strict") {
                                // Strict has always rejected partial selectors; preserve that contract.
                                await expect(validate(target, { profile, ...selection })).rejects.toBeInstanceOf(ModeConflictError);
                                continue;
                            }
                            const partial = await validate(target, { profile, ...selection });
                            const cleanPartial = await validate(control, { profile, ...selection });
                            expect(findings(partial.issues)).toEqual(findings(full.issues));
                            expect(partial.valid).toBe(full.valid);
                            if (full.valid) expect(partial.valid).toBe(true);
                            expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                            if (clean.valid) expect(cleanPartial.valid).toBe(true);
                            expect(errors(clean.issues)).toEqual(expect.arrayContaining(errors(cleanPartial.issues)));
                        }
                        if (before) expect(readFileSync(target)).toEqual(before);
                        else
                            for (const [name, text] of Object.entries(parts))
                                expect(readFileSync(path.join(target, name), "utf8")).toBe(text);
                    });
                },
            );
        }

    it("retains the stored stream path and suppresses duplicates when the stream cannot parse", async () => {
        await withTempDir(async (dir) => {
            const stream = "[CONTENT_TYPES].XML";
            const { [STREAM]: types, ...rest } = MINIMAL_PARTS;
            const duplicate = types!.replace("</Types>", `${XML_DEFAULT}</Types>`);
            const target = await writePackage(dir, "case.docx", { ...rest, [stream]: duplicate }, input);
            const full = await validate(target);
            expect(findings(full.issues)).toEqual([expect.objectContaining({ code: ID, path: stream, severity: "error" })]);
            const broken = await writePackage(dir, "broken.docx", { ...rest, [stream]: duplicate.replace("</Types>", "") }, input);
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(broken, { profile });
                expect(result.valid).toBe(false);
                expect(findings(result.issues)).toEqual([]);
                expect(result.issues).toEqual(
                    expect.arrayContaining([expect.objectContaining({ code: "xml-syntax", path: stream, severity: "error" })]),
                );
                if (profile !== "strict")
                    for (const selection of SELECTIONS) {
                        const partial = await validate(broken, { profile, ...selection });
                        expect(partial.valid).toBe(false);
                        expect(findings(partial.issues)).toEqual([]);
                        expect(errors(result.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                    }
            }
        });
    });

    it("keeps unobserved PPTX/XLSX advisory outside strict", async () => {
        await withTempDir(async (dir) => {
            for (const suffix of ["pptx", "xlsx"] as const) {
                const parts = {
                    ...MINIMAL_PPTX_PARTS,
                    [STREAM]: MINIMAL_PPTX_PARTS[STREAM]!.replace("</Types>", `${XML_DEFAULT}</Types>`),
                };
                const target = await writePackage(dir, `specimen.${suffix}`, parts, input);
                for (const profile of suffix === "xlsx"
                    ? (["word-valid", "lenient"] as const)
                    : (["word-valid", "lenient", "strict"] as const)) {
                    // Full XLSX validation is unsupported; registered partial checks are supported.
                    const result = await validate(target, { profile, ...(suffix === "xlsx" ? { checks: ["content-types"] } : {}) });
                    expect(findings(result.issues)).toEqual([
                        expect.objectContaining({
                            code: ID,
                            path: STREAM,
                            severity: profile === "strict" && suffix === "pptx" ? "error" : "warning",
                        }),
                    ]);
                    if (suffix === "pptx") {
                        expect(result.valid).toBe(profile !== "strict");
                        expect(result.metadata?.executedChecks).toContain(ID);
                    }
                    if (profile !== "strict")
                        for (const selection of SELECTIONS) {
                            const partial = await validate(target, { profile, ...selection });
                            expect(findings(partial.issues)).toEqual(findings(result.issues));
                            expect(partial.valid).toBe(true);
                            expect(errors(result.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                        }
                }
            }
        });
    });

    it("reports PPTX duplicates in a full run while another part is malformed, except under lenient (#456)", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                [STREAM]: MINIMAL_PPTX_PARTS[STREAM]!.replace("</Types>", `${XML_DEFAULT}</Types>`),
                "junk.xml": "<broken",
            };
            const target = await writePackage(dir, "malformed-neighbour.pptx", parts, input);
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const result = await validate(target, { profile });
                expect(result.valid, profile).toBe(false);
                // Lenient stops at the malformed part, as its partial runs do
                // (#391); word-valid and strict partial runs judge the stream.
                expect(findings(result.issues), profile).toEqual(
                    profile === "lenient"
                        ? []
                        : [expect.objectContaining({ code: ID, path: STREAM, severity: profile === "strict" ? "error" : "warning" })],
                );
                if (profile !== "strict")
                    for (const selection of SELECTIONS) {
                        const partial = await validate(target, { profile, ...selection });
                        const tag = `${profile} ${JSON.stringify(selection)}`;
                        expect(findings(partial.issues), tag).toEqual(findings(result.issues));
                        expect(errors(result.issues), tag).toEqual(expect.arrayContaining(errors(partial.issues)));
                    }
            }
        });
    });
});
