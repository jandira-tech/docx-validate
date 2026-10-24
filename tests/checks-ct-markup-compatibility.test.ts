/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `ct-markup-compatibility` check (#459): Microsoft Word 16.115 refuses
 * any Markup Compatibility usage in the `[Content_Types].xml` stream — an
 * `mc:Ignorable` attribute, an ignored foreign element, or an
 * `mc:AlternateContent` in either orientation all get a repair offer (probe
 * session 14a, recorded in docs/plans/evidence/word-probes-2026-10-24.md). A
 * bare `xmlns:mc` declaration that nothing uses opens clean.
 *
 * The check reads the raw stored stream, reports one finding per stream
 * listing the markup it found, and is Word-blocking for DOCX/DOCM; strict
 * reports it as an error (its XSD walk already refuses the stream), and
 * lenient and the unobserved PPTX/XLSX warn.
 *
 * The specimens are the session-14a ones: `MINIMAL_PARTS` with the
 * content-types stream changed as named. `BAD` is an Override Word refuses on
 * its own (OPC [M1.8], session 11) and `GOOD` one it accepts (session 7); the
 * Choice reads `Requires="x"`, so a processor ignores it and takes the
 * Fallback.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const ID = "ct-markup-compatibility";
const STREAM = "[Content_Types].xml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MC_DECL = `xmlns:mc="${MC}" xmlns:x="urn:x:unknown"`;

/** An Override Word refuses on its own (OPC [M1.8], session 11) and one it accepts (session 7). */
const BAD = `<Override PartName="/word/a%41b.xml" ContentType="application/xml"/>`;
const GOOD = `<Override PartName="/word/fallback.xml" ContentType="application/xml"/>`;
const alternate = (choice: string, fallback: string) =>
    `<mc:AlternateContent><mc:Choice Requires="x">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
/** An AlternateContent whose Choice carries no usable Requires (#494): none, empty, or blank. */
const noRequires = (requires: string, choice: string, fallback: string) =>
    `<mc:AlternateContent><mc:Choice${requires}>${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

/** The content-types stream with extra namespace declarations/attributes on the root and extra children. */
const withContentTypes = (
    parts: Readonly<Record<string, string>>,
    rootExtra: string,
    children: string,
): Readonly<Record<string, string>> => {
    const source = parts[STREAM]!;
    const open = source.match(/<Types [^>]*>/)![0];
    return { ...parts, [STREAM]: source.replace(open, open.replace(/>$/, ` ${rootExtra}>`)).replace("</Types>", `${children}</Types>`) };
};

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** The session-14a specimens: name → parts. */
const SPECIMENS: Readonly<Record<string, Readonly<Record<string, string>>>> = {
    "459-mc-namespace-declared-only": withContentTypes(MINIMAL_PARTS, MC_DECL, ""),
    "459-mc-ignorable-only": withContentTypes(MINIMAL_PARTS, `${MC_DECL} mc:Ignorable="x"`, ""),
    "459-mc-ignorable-ignored-element": withContentTypes(MINIMAL_PARTS, `${MC_DECL} mc:Ignorable="x"`, `<x:Foo/>`),
    "459-mc-choice-bad-fallback-good": withContentTypes(MINIMAL_PARTS, MC_DECL, alternate(BAD, GOOD)),
    "459-mc-choice-good-fallback-bad": withContentTypes(MINIMAL_PARTS, MC_DECL, alternate(GOOD, BAD)),
    "494-mc-choice-no-requires-bad-fallback-good": withContentTypes(MINIMAL_PARTS, MC_DECL, noRequires("", BAD, GOOD)),
    "494-mc-choice-no-requires-good-fallback-bad": withContentTypes(MINIMAL_PARTS, MC_DECL, noRequires("", GOOD, BAD)),
    "494-mc-choice-empty-requires-bad-fallback-good": withContentTypes(MINIMAL_PARTS, MC_DECL, noRequires(' Requires=""', BAD, GOOD)),
    "493-mc-ignorable-own-namespace": withContentTypes(
        MINIMAL_PARTS,
        `xmlns:mc="${MC}" xmlns:t="${CONTENT_TYPES_NAMESPACE}" mc:Ignorable="t"`,
        `<t:Default Extension="png" ContentType="image/png"/>`,
    ),
};

const findings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === ID);
const partnameFindings = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code === "ct-override-partname-invalid");
const errorPairs = (issues: readonly ValidationIssue[]) =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);

const found = (...items: string[]) => ({
    code: ID,
    severity: "error",
    path: STREAM,
    message: `Markup Compatibility markup in the content-types stream: ${items.join("; ")}`,
});

const diagnose = async (ctx: ValidationContext): Promise<readonly ValidationIssue[]> => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check, "the check is registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};

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

describe("ct-markup-compatibility", () => {
    it("registers in the content-types group for every package format", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["content-types"], formats: ["docx", "pptx", "xlsx"] });
        // It reads a part's DOM, so the XML syntax gate must precede it.
        expect(check?.requires ?? []).toContain("xml-syntax");
    });

    it("reports nothing from an absent or unparseable stream", async () => {
        expect(await diagnose(ValidationContext.fromEntries({}))).toEqual([]);
        expect(await diagnose(ValidationContext.fromEntries({ [STREAM]: "<Types><Override" }))).toEqual([]);
    });

    it("reports nothing for a bare xmlns:mc declaration nothing uses", async () => {
        const ctx = ValidationContext.fromEntries(SPECIMENS["459-mc-namespace-declared-only"]!);
        expect(await diagnose(ctx)).toEqual([]);
    });

    it("reports mc:Ignorable alone once, at the stored stream", async () => {
        const ctx = ValidationContext.fromEntries(SPECIMENS["459-mc-ignorable-only"]!);
        expect(await diagnose(ctx)).toEqual([found("mc:Ignorable attribute on <Types>")]);
    });

    it("reports an element in a namespace mc:Ignorable lists, in the same finding", async () => {
        const ctx = ValidationContext.fromEntries(SPECIMENS["459-mc-ignorable-ignored-element"]!);
        expect(await diagnose(ctx)).toEqual([
            found("mc:Ignorable attribute on <Types>", 'element <x:Foo> in the ignorable namespace "urn:x:unknown"'),
        ]);
    });

    it("reports every Markup Compatibility element of an AlternateContent", async () => {
        for (const name of ["459-mc-choice-bad-fallback-good", "459-mc-choice-good-fallback-bad"] as const) {
            const ctx = ValidationContext.fromEntries(SPECIMENS[name]!);
            expect(await diagnose(ctx), name).toEqual([
                found("<mc:AlternateContent> element", "<mc:Choice> element", "<mc:Fallback> element"),
            ]);
        }
    });

    it("lists every Markup Compatibility element once per occurrence, in document order (#495)", async () => {
        // An outer AlternateContent whose Choice holds an inner one: two
        // Choices and two Fallbacks, each listed once per occurrence.
        const nested = withContentTypes(
            MINIMAL_PARTS,
            MC_DECL,
            `<mc:AlternateContent><mc:Choice Requires="x">${alternate(BAD, GOOD)}</mc:Choice><mc:Fallback>${GOOD}</mc:Fallback></mc:AlternateContent>`,
        );
        expect(await diagnose(ValidationContext.fromEntries(nested))).toEqual([
            found(
                "<mc:AlternateContent> element",
                "<mc:Choice> element",
                "<mc:AlternateContent> element",
                "<mc:Choice> element",
                "<mc:Fallback> element",
                "<mc:Fallback> element",
            ),
        ]);
    });

    it("lists an MC attribute once per element that carries it (#495)", async () => {
        const ctx = ValidationContext.fromEntries(
            withContentTypes(
                MINIMAL_PARTS,
                `${MC_DECL} mc:Ignorable="x"`,
                `<Override PartName="/word/x.xml" ContentType="application/xml" mc:Ignorable="x"/>`,
            ),
        );
        expect(await diagnose(ctx)).toEqual([found("mc:Ignorable attribute on <Types>", "mc:Ignorable attribute on <Override>")]);
    });

    it("reports mc:MustUnderstand and mc:ProcessContent attributes", async () => {
        const ctx = ValidationContext.fromEntries(
            withContentTypes(MINIMAL_PARTS, `${MC_DECL} mc:MustUnderstand="x" mc:ProcessContent="x:Foo"`, ""),
        );
        expect(await diagnose(ctx)).toEqual([found("mc:MustUnderstand attribute on <Types>", "mc:ProcessContent attribute on <Types>")]);
    });

    it("reports an attribute in a namespace mc:Ignorable lists", async () => {
        const ctx = ValidationContext.fromEntries(
            withContentTypes(
                MINIMAL_PARTS,
                `${MC_DECL} mc:Ignorable="x"`,
                `<Override PartName="/word/x.xml" ContentType="application/xml" x:note="v"/>`,
            ),
        );
        expect(await diagnose(ctx)).toEqual([
            found("mc:Ignorable attribute on <Types>", 'x:note attribute on <Override> in the ignorable namespace "urn:x:unknown"'),
        ]);
    });

    it("reports at the stored stream path, whatever its spelling", async () => {
        const ctx = ValidationContext.fromEntries({
            "[CONTENT_TYPES].XML": withContentTypes(MINIMAL_PARTS, `${MC_DECL} mc:Ignorable="x"`, "")[STREAM]!,
        });
        expect(await diagnose(ctx)).toEqual([{ ...found("mc:Ignorable attribute on <Types>"), path: "[CONTENT_TYPES].XML" }]);
    });

    it("is Word-blocking for DOCX and DOCM", () => {
        expect(isWordBlockingIssue({ code: ID, severity: "error", message: "", path: STREAM }, new Set())).toBe(true);
    });
});

describe("ct-markup-compatibility: a Choice without Requires is never selected (#494)", { timeout: 120_000 }, () => {
    it.each(["494-mc-choice-no-requires-bad-fallback-good", "494-mc-choice-empty-requires-bad-fallback-good"] as const)(
        "%s: the Fallback's declarations are judged, not the unselectable Choice's",
        async (name) => {
            await withTempDir(async (dir) => {
                const target = path.join(dir, "specimen.docx");
                writeFileSync(target, await buildDocx(SPECIMENS[name]!));
                const wordValid = await validate(target, { profile: "word-valid" });
                for (const profile of ["word-valid", "strict"] as const) {
                    const result = profile === "word-valid" ? wordValid : await validate(target, { profile });
                    // The BAD Override sits in the malformed Choice, which a
                    // processor must not select (mc.xsd: Requires use="required");
                    // the Fallback's GOOD Override draws nothing. The run stays
                    // invalid through ct-markup-compatibility alone.
                    expect(partnameFindings(result.issues), profile).toEqual([]);
                    expect(findings(result.issues), profile).toEqual([
                        expect.objectContaining({ code: ID, path: STREAM, severity: "error" }),
                    ]);
                    expect(result.valid, profile).toBe(false);
                }
                for (const selection of [{ checks: ["ct-override-partname-invalid"] }, { checks: ["content-types"] }]) {
                    const partial = await validate(target, { profile: "word-valid", ...selection });
                    const tag = JSON.stringify(selection);
                    expect(partnameFindings(partial.issues), tag).toEqual([]);
                    // #381: partial errors are a subset of the full run's.
                    expect(errorPairs(wordValid.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            });
        },
    );

    it("494-mc-choice-no-requires-good-fallback-bad: the Fallback is judged once the Choice is unselectable", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "specimen.docx");
            writeFileSync(target, await buildDocx(SPECIMENS["494-mc-choice-no-requires-good-fallback-bad"]!));
            const result = await validate(target, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(partnameFindings(result.issues)).toEqual([
                expect.objectContaining({ code: "ct-override-partname-invalid", path: STREAM }),
            ]);
            const partial = await validate(target, { profile: "word-valid", checks: ["ct-override-partname-invalid"] });
            expect(partnameFindings(partial.issues)).toEqual(partnameFindings(result.issues));
        });
    });
});

describe("ct-markup-compatibility: an Ignorable naming the stream's own namespace blinds no pass (#493)", { timeout: 120_000 }, () => {
    const coverageFindings = (issues: readonly ValidationIssue[]) =>
        issues.filter((issue) => ["ct-uncovered-part", "ct-undeclared-part", "ct-undeclared-ext"].includes(issue.code ?? ""));

    it("reports no coverage findings in full and partial runs", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "specimen.docx");
            writeFileSync(target, await buildDocx(SPECIMENS["493-mc-ignorable-own-namespace"]!));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(target, { profile });
                // The processor applying OPC Part 2 §10 to this stream
                // understands the content-types vocabulary itself, so the
                // Ignorable naming it cannot un-declare every part's type.
                expect(
                    coverageFindings(result.issues).map((issue) => issue.code),
                    profile,
                ).toEqual([]);
                expect(findings(result.issues), profile).toEqual([expect.objectContaining({ code: ID, path: STREAM })]);
            }
            for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                const partial = await validate(target, { profile: "word-valid", ...selection });
                const tag = JSON.stringify(selection);
                expect(
                    coverageFindings(partial.issues).map((issue) => issue.code),
                    tag,
                ).toEqual([]);
                expect(findings(partial.issues), tag).toEqual([expect.objectContaining({ code: ID, path: STREAM })]);
            }
        });
    });
});

describe.each(["packed", "directory"] as const)("ct-markup-compatibility: %s (#459)", { timeout: 120_000 }, (input) => {
    for (const suffix of ["docx", "docm"] as const) {
        it(`${suffix}: the minimal control and the declared-only specimen pass every profile`, async () => {
            await withTempDir(async (dir) => {
                for (const [label, parts] of [
                    ["control", MINIMAL_PARTS],
                    ["declared-only", SPECIMENS["459-mc-namespace-declared-only"]!],
                ] as const) {
                    const target = await writePackage(dir, `${label}.${suffix}`, parts, input);
                    for (const profile of ["word-valid", "lenient", "strict"] as const) {
                        const result = await validate(target, { profile });
                        expect(findings(result.issues), profile).toEqual([]);
                        expect(result.valid, profile).toBe(true);
                    }
                }
            });
        });

        it.each(["459-mc-ignorable-only", "459-mc-ignorable-ignored-element"] as const)(
            `${suffix}: %s is invalid under word-valid with exactly one ${ID} error`,
            async (name) => {
                await withTempDir(async (dir) => {
                    const target = await writePackage(dir, `specimen.${suffix}`, SPECIMENS[name]!, input);
                    const before = input === "packed" ? readFileSync(target) : undefined;
                    const wordValid = await validate(target, { profile: "word-valid" });
                    // Word 16.115 offers recovery (session 14a); the run must
                    // fail through the new check alone.
                    expect(wordValid.valid).toBe(false);
                    expect(errorPairs(wordValid.issues)).toEqual([`${ID}@${STREAM}`]);
                    const lenient = await validate(target, { profile: "lenient" });
                    expect(findings(lenient.issues)).toEqual([expect.objectContaining({ code: ID, path: STREAM, severity: "warning" })]);
                    const strict = await validate(target, { profile: "strict" });
                    expect(strict.valid).toBe(false);
                    expect(errorPairs(strict.issues)).toEqual(expect.arrayContaining([`${ID}@${STREAM}`, `xsd-error@${STREAM}`]));
                    for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                        const partial = await validate(target, { profile: "word-valid", ...selection });
                        const tag = JSON.stringify(selection);
                        expect(partial.valid, tag).toBe(false);
                        expect(errorPairs(partial.issues), tag).toEqual([`${ID}@${STREAM}`]);
                        // #381: partial errors are a subset of the full run's.
                        expect(errorPairs(wordValid.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                        await expect(validate(target, { profile: "strict", ...selection }), tag).rejects.toBeInstanceOf(ModeConflictError);
                    }
                    if (before) expect(readFileSync(target)).toEqual(before);
                    else
                        for (const [part, text] of Object.entries(SPECIMENS[name]!))
                            expect(readFileSync(path.join(target, part), "utf8")).toBe(text);
                });
            },
        );

        it(`${suffix}: an AlternateContent keeps the verdict through ${ID} alone when the bad Override sits in the ignored Choice`, async () => {
            await withTempDir(async (dir) => {
                const target = await writePackage(dir, `specimen.${suffix}`, SPECIMENS["459-mc-choice-bad-fallback-good"]!, input);
                for (const profile of ["word-valid", "strict"] as const) {
                    const result = await validate(target, { profile });
                    // The bad Override sits in the dropped Choice ([M1.19]), so
                    // ct-override-partname-invalid no longer fires; the run stays
                    // invalid through ct-markup-compatibility.
                    expect(partnameFindings(result.issues), profile).toEqual([]);
                    expect(findings(result.issues), profile).toEqual([
                        expect.objectContaining({ code: ID, path: STREAM, severity: "error" }),
                    ]);
                    expect(result.valid, profile).toBe(false);
                }
            });
        });

        it(`${suffix}: an AlternateContent reports both findings when the selected Fallback holds the bad Override`, async () => {
            await withTempDir(async (dir) => {
                const target = await writePackage(dir, `specimen.${suffix}`, SPECIMENS["459-mc-choice-good-fallback-bad"]!, input);
                const result = await validate(target, { profile: "word-valid" });
                expect(result.valid).toBe(false);
                expect(errorPairs(result.issues)).toEqual(
                    expect.arrayContaining([`${ID}@${STREAM}`, `ct-override-partname-invalid@${STREAM}`]),
                );
                for (const selection of [{ checks: ["content-types"] }, { checks: [ID] }]) {
                    const partial = await validate(target, { profile: "word-valid", ...selection });
                    const tag = JSON.stringify(selection);
                    expect(errorPairs(partial.issues), tag).toContain(`${ID}@${STREAM}`);
                    expect(errorPairs(result.issues), tag).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                }
            });
        });
    }

    it("warns for PPTX under every profile, while its schema walk still rejects the stream", async () => {
        await withTempDir(async (dir) => {
            const parts = withContentTypes(MINIMAL_PPTX_PARTS, `${MC_DECL} mc:Ignorable="x"`, "");
            const target = path.join(dir, `specimen${input === "packed" ? "" : "-dir"}.pptx`);
            if (input === "packed") writeFileSync(target, await buildPptx(parts));
            else
                for (const [name, text] of Object.entries(parts)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(target, { profile });
                // Nothing is observed for PPTX: the new finding warns, and the
                // existing schema walk keeps the package invalid on its own.
                expect(findings(result.issues), profile).toEqual([
                    expect.objectContaining({ code: ID, path: STREAM, severity: "warning" }),
                ]);
                expect(result.valid, profile).toBe(false);
                expect(errorPairs(result.issues), profile).toContain(`xsd-error@${STREAM}`);
            }
        });
    });

    it("warns for XLSX in a partial run, as no full XLSX validation exists", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.xlsx");
            writeFileSync(file, await buildDocx(SPECIMENS["459-mc-ignorable-only"]!));
            for (const profile of ["word-valid", "lenient"] as const) {
                const result = await validate(file, { profile, checks: [ID] });
                expect(findings(result.issues), profile).toEqual([
                    expect.objectContaining({ code: ID, path: STREAM, severity: "warning" }),
                ]);
                expect(result.valid, profile).toBe(true);
            }
        });
    });
});
