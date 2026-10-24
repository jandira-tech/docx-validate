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

import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { parseContentTypes, resolveContentType } from "../src/lib/content-types";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { parseXml } from "../src/lib/xml-helpers";
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
    // `/word/café.xml` is no longer a specimen: a literal ucschar is a valid
    // Part IRI under ECMA-376 Part 2 §9.1.1.1 (#455); the describe below
    // covers it and the names that stay invalid.
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

    it("says an Override has no PartName rather than quoting an empty one (#458)", async () => {
        // Word 16.115 offers recovery for both (session 13); only the wording differs.
        const missing = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override ContentType="application/xml"/></Types>`,
            ),
        });
        const issues = await diagnose(missing);
        expect(issues).toEqual([
            {
                code: ID,
                severity: "error",
                path: "[Content_Types].xml",
                message: "Override has no PartName attribute, so it names no part.",
            },
        ]);
        expect(isWordBlockingIssue(issues[0]!, new Set())).toBe(true);
        const empty = ValidationContext.fromEntries(withOverride("", "application/xml"));
        expect(await diagnose(empty)).toEqual([expected("", "must start with '/'", "error")]);
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

    it("leaves resolution unchanged: a malformed PartName can still resolve a stored part (#457)", async () => {
        // The finding is reported, but content types are still keyed by the
        // canonical spelling, which supplies the leading slash.
        const parts = {
            ...MINIMAL_PARTS,
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                'PartName="/word/document.xml"',
                'PartName="word/document.xml"',
            ),
        };
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expected("word/document.xml", "must start with '/'", "error"),
        ]);
        expect(resolveContentType(parseContentTypes(parseXml(parts["[Content_Types].xml"])), "word/document.xml")).toEqual({
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            source: "override",
        });
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

    it("reports for PPTX in a full run while another part is malformed, except under lenient (#456)", async () => {
        await withTempDir(async (dir) => {
            const partName = "/ppt/../ppt/nothere.xml";
            const parts = {
                ...MINIMAL_PPTX_PARTS,
                "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
                    "</Types>",
                    `<Override PartName="${partName}" ContentType="application/xml"/></Types>`,
                ),
                "junk.xml": "<broken",
            };
            const packed = path.join(dir, "malformed-neighbour.pptx");
            writeFileSync(packed, await buildPptx(parts));
            const unpacked = path.join(dir, "malformed-neighbour-dir.pptx");
            for (const [name, text] of Object.entries(parts)) {
                mkdirSync(path.dirname(path.join(unpacked, name)), { recursive: true });
                writeFileSync(path.join(unpacked, name), text);
            }
            for (const file of [packed, unpacked]) {
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const result = await validate(file, { profile });
                    const tag = `${path.basename(file)} ${profile}`;
                    expect(result.valid, tag).toBe(false);
                    // Lenient stops at the malformed part, as its partial runs
                    // do (#391); word-valid and strict judge the stream.
                    expect(findings(result.issues), tag).toEqual(
                        profile === "lenient"
                            ? []
                            : [expected(partName, "must not have a '.' or '..' segment", profile === "strict" ? "error" : "warning")],
                    );
                    if (profile !== "strict")
                        for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                            const partial = await validate(file, { profile, ...selection });
                            expect(findings(partial.issues), `${tag} ${JSON.stringify(selection)}`).toEqual(findings(result.issues));
                        }
                }
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

    it("fails word-valid for an Override without a PartName, in full and partial runs (#458)", async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "no-partname", {
                ...MINIMAL_PARTS,
                "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                    "</Types>",
                    `<Override ContentType="application/xml"/></Types>`,
                ),
            });
            const finding = {
                code: ID,
                severity: "error",
                path: "[Content_Types].xml",
                message: "Override has no PartName attribute, so it names no part.",
            };
            const full = await validate(file, { profile: "word-valid" });
            expect(findings(full.issues)).toEqual([finding]);
            expect(full.valid).toBe(false);
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(findings(partial.issues), tag).toEqual([finding]);
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

    it("does not resolve a PartName with a dot segment onto a stored part", async () => {
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

describe("ct-override-partname-invalid under ECMA-376 Part 2 §9.1.1.1 (4th edition, #455)", () => {
    /** Names the 4th edition accepts: each is a valid Part URI or a valid Part IRI. */
    const PART_IRI_OR_URI_NAMES = ["/word/café.xml", "/a/%D1%86.xml", "/word/caf%C3%A9.xml", "/word/caf%c3%a9.xml"] as const;

    it.each(PART_IRI_OR_URI_NAMES)("accepts %s, a valid Part IRI or Part URI", async (partName) => {
        const ctx = ValidationContext.fromEntries(withOverride(partName, "application/xml"));
        expect(await diagnose(ctx)).toEqual([]);
    });

    it.each([
        // A literal ucschar leaves only the Part IRI form, and [M1.8] reads
        // "iunreserved" there: %C3%A9 encodes é.
        ["/word/caf%C3%A9-é.xml", "must not percent-encode an unreserved character (OPC [M1.8])"],
        // A literal non-ASCII character that is not ucschar is in neither grammar.
        ["/word/x\u{0080}y.xml", "must contain only RFC 3986 pchar characters"],
        ["/word/x\u{E000}y.xml", "must contain only RFC 3986 pchar characters"],
        ["/word/x\u{FFFE}y.xml", "must contain only RFC 3986 pchar characters"],
    ])("still reports %s", async (partName, rule) => {
        const ctx = ValidationContext.fromEntries(withOverride(partName, "application/xml"));
        expect(await diagnose(ctx)).toEqual([expected(partName, rule, "error")]);
    });

    it("keeps the escaped-iunreserved finding Word-blocking, as every other [M1.8] finding is", async () => {
        const [issue] = await diagnose(ValidationContext.fromEntries(withOverride("/word/caf%C3%A9-é.xml", "application/xml")));
        expect(issue).toBeDefined();
        expect(isWordBlockingIssue(issue!, new Set())).toBe(true);
    });
});

describe("a literal-ucschar Override PartName through validate() (#455)", { timeout: 120_000 }, () => {
    const CAFE = "/word/café.xml";
    const MIXED = "/word/caf%C3%A9-é.xml";
    const MIXED_RULE = "must not percent-encode an unreserved character (OPC [M1.8])";

    const withPptxOverride = (partName: string): Readonly<Record<string, string>> => ({
        ...MINIMAL_PPTX_PARTS,
        "[Content_Types].xml": MINIMAL_PPTX_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="${partName}" ContentType="application/xml"/></Types>`,
        ),
    });

    it("no longer reports /word/café.xml for PPTX under any profile, and strict stays valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "part-iri.pptx");
            writeFileSync(file, await buildPptx(withPptxOverride(CAFE)));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                const tag = `${profile}`;
                expect(findings(result.issues), tag).toEqual([]);
                // Word 16.115 offers recovery for this spelling (probe sessions
                // 7–12), so the Word-specific rule reports it as a warning (#451).
                const refused = result.issues.filter((issue) => issue.code === "ct-override-partname-word-refused");
                expect(refused, tag).toHaveLength(1);
                expect(refused[0]).toMatchObject({ severity: "warning", path: "[Content_Types].xml" });
                expect(refused[0]!.message).toContain("literal non-ASCII characters");
                expect(
                    result.issues.filter((issue) => issue.severity === "error"),
                    tag,
                ).toEqual([]);
                expect(result.valid, tag).toBe(true);
            }
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    expect(findings(partial.issues), tag).toEqual([]);
                    expect(
                        partial.issues.filter((issue) => issue.code === "ct-override-partname-word-refused"),
                        tag,
                    ).toHaveLength(selection.checks?.includes(ID) ? 0 : 1);
                    // #381: a warning-only full run keeps every partial run valid.
                    expect(partial.valid, tag).toBe(true);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            }
        });
    });

    it.each(["docx", "docm"] as const)(
        "still refuses /word/café.xml for a %s under word-valid and strict, through the Word rule",
        async (extension) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `part-iri.${extension}`);
                writeFileSync(file, await buildDocx(withOverride(CAFE, "application/xml")));
                for (const profile of ["word-valid", "lenient", "strict"] as const) {
                    const result = await validate(file, { profile });
                    const tag = `${profile}`;
                    expect(findings(result.issues), tag).toEqual([]);
                    const refused = result.issues.filter((issue) => issue.code === "ct-override-partname-word-refused");
                    expect(refused, tag).toHaveLength(1);
                    expect(refused[0]!.severity, tag).toBe(profile === "lenient" ? "warning" : "error");
                    expect(refused[0]!.message).toContain("literal non-ASCII characters");
                    expect(result.valid, tag).toBe(profile === "lenient");
                    if (profile !== "strict")
                        for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                            const partial = await validate(file, { profile, ...selection });
                            const partialTag = `${profile} ${JSON.stringify(selection)}`;
                            expect(findings(partial.issues), partialTag).toEqual([]);
                            expect(errorPairs(result.issues), partialTag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        }
                }
            });
        },
    );

    it.each([
        ["docx", buildDocx, () => withOverride(MIXED, "application/xml")],
        ["pptx", buildPptx, () => withPptxOverride(MIXED)],
    ] as const)("%s strict errors on the mixed spelling /word/caf%C3%A9-é.xml", async (extension, build, parts) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `mixed.${extension}`);
            writeFileSync(file, await build(parts()));
            const result = await validate(file, { profile: "strict" });
            expect(findings(result.issues)).toEqual([expected(MIXED, MIXED_RULE, "error")]);
            expect(result.valid).toBe(false);
        });
    });

    it("keeps full and partial runs agreeing on the mixed spelling for a DOCX (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "mixed.docx");
            writeFileSync(file, await buildDocx(withOverride(MIXED, "application/xml")));
            for (const profile of ["word-valid", "lenient"] as const) {
                const full = await validate(file, { profile });
                const severity = profile === "lenient" ? "warning" : "error";
                expect(findings(full.issues)).toEqual([expected(MIXED, MIXED_RULE, severity)]);
                for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                    const partial = await validate(file, { profile, ...selection });
                    const tag = `${profile} ${JSON.stringify(selection)}`;
                    expect(findings(partial.issues), tag).toEqual([expected(MIXED, MIXED_RULE, severity)]);
                    expect(errorPairs(full.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            }
        });
    });
});
