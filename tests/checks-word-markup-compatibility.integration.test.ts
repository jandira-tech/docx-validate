/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { NS, type ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import {
    RELATIONSHIPS_MCE_CASES,
    STORY_RELS_MCE_CASES,
    documentRelsChoiceUndeclaredPrefix,
    documentRelsRootMustUnderstand,
    packageRelsRootMustUnderstand,
} from "./support/relationships-markup-compatibility-specimens";

const CODE = "word-markup-compatibility";
const PACKAGE = "_rels/.rels";
const SIDECAR = "word/_rels/document.xml.rels";
const mce = (issues: readonly ValidationIssue[]) => issues.filter((issue) => issue.code?.startsWith("mce-"));
const errors = (issues: readonly ValidationIssue[]) =>
    new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`));

/**
 * Session 34 (#636): Word honours markup-compatibility errors in
 * relationships parts, so word-valid and strict report them at the stored
 * part name, and quick and named runs agree with the full run (#381).
 */
describe("markup compatibility in relationships parts (#636, #381)", () => {
    const specimens = [
        {
            label: "mce-rels-root-mustunderstand",
            parts: packageRelsRootMustUnderstand,
            // s34c: recovery offer, then "Word experienced an error trying to
            // open the file." after it was declined.
            code: "mce-must-understand",
            part: PACKAGE,
            wordValidBlocks: true,
            // Lenient walks the OPC schema over the stored root and runs no
            // Word rule, so its findings stay as they were (#636 changes
            // nothing there; this pins them).
            lenient: ["error:xsd-error@_rels/.rels", "error:xsd-summary@_rels/.rels", "info:latent-styles-missing@word/styles.xml"],
        },
        {
            label: "mce-rels-choice-undeclared-prefix",
            parts: documentRelsChoiceUndeclaredPrefix,
            // s34d: recovery offer. The footer relationship disappears, so
            // word-header-footer-relationship already rejected the file.
            code: "mce-malformed",
            part: SIDECAR,
            wordValidBlocks: true,
            lenient: [
                "error:mce-malformed@word/_rels/document.xml.rels",
                "info:latent-styles-missing@word/styles.xml",
                "warning:paraid-missing-element@word/footer1.xml",
            ],
        },
        {
            label: "mce-rels-sidecar-root-mustunderstand",
            parts: documentRelsRootMustUnderstand,
            // Session 36 s36b observed recovery on the main sidecar too (#644).
            code: "mce-must-understand",
            part: SIDECAR,
            wordValidBlocks: true,
            lenient: [
                "error:mce-must-understand@word/_rels/document.xml.rels",
                "info:latent-styles-missing@word/styles.xml",
                "warning:paraid-missing-element@word/footer1.xml",
            ],
        },
    ] as const;

    for (const suffix of ["docx", "docm"] as const) {
        it.each(specimens)(`${suffix}: $label`, async ({ parts, code, part, wordValidBlocks, lenient }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `specimen.${suffix}`);
                const bytes = await buildDocx(parts());
                await writeFile(file, bytes);

                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity: wordValidBlocks ? "error" : "warning", path: part }),
                ]);
                expect(wordValid.valid).toBe(!wordValidBlocks);

                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity: "error", path: part }),
                ]);
                expect(strict.valid).toBe(false);

                // Session 34 s34d also loses the footer relationship, so the
                // existing finding stays beside the new one. The sidecar-root
                // shape keeps its relationship, so it has none.
                if (code === "mce-malformed") {
                    expect(wordValid.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([
                        expect.objectContaining({ severity: "error", path: "word/document.xml" }),
                    ]);
                } else {
                    expect(wordValid.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
                }

                for (const selection of [{ quick: true }, { checks: [CODE] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect.soft(partial.valid, JSON.stringify(selection)).toBe(wordValid.valid);
                    expect.soft(mce(partial.issues), JSON.stringify(selection)).toEqual(mce(wordValid.issues));
                    expect
                        .soft(
                            [...errors(partial.issues)].filter((pair) => !errors(wordValid.issues).has(pair)),
                            JSON.stringify(selection),
                        )
                        .toEqual([]);
                }

                // Lenient runs no Word rule, so its findings stay as they were.
                const lenientRun = await validate(file, { profile: "lenient" });
                expect(lenientRun.issues.map((issue) => `${issue.severity}:${issue.code}@${issue.path ?? ""}`).sort()).toEqual([
                    ...lenient,
                ]);
                expect(await readFile(file)).toEqual(bytes);
            });
        });
    }

    it.each(RELATIONSHIPS_MCE_CASES)(
        "session 33's $label keeps no mce-* finding and its word-valid verdict (#632)",
        async ({ parts, clean }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts));
                for (const profile of ["word-valid", "strict"] as const) {
                    const result: { valid: boolean; issues: readonly ValidationIssue[] } = await validate(file, { profile });
                    expect.soft(mce(result.issues), profile).toEqual([]);
                    if (profile === "word-valid") expect(result.valid).toBe(clean);
                }
            });
        },
    );
});

// Captured from unchanged 840df310 production sources before the #644 fix.
// Explicit values preserve message text, positions, ordering and validity.
const lenientIssues: readonly ValidationIssue[] = [
    {
        severity: "info",
        message:
            "word/styles.xml has no <w:latentStyles> block. This is unusual for Word-generated files and indicates the repairer may have regenerated styles from scratch.",
        path: "word/styles.xml",
        code: "latent-styles-missing",
    },
    {
        severity: "error",
        code: "mce-malformed",
        message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
        path: "word/_rels/document.xml.rels",
    },
    {
        severity: "error",
        message: "_rels/.rels: 1 new error(s)",
        path: "_rels/.rels",
        code: "xsd-summary",
    },
    {
        severity: "error",
        message:
            "  - Element '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent': This element is not expected. Expected is ( {http://schemas.openxmlformats.org/package/2006/relationships}Relationship ).",
        path: "_rels/.rels",
        code: "xsd-error",
    },
    {
        severity: "warning",
        message:
            "1 <w:p> element(s) lack a w14:paraId. Word stamps one on every paragraph and table row as the anchor for tracked-changes / comment-range infrastructure.",
        path: "word/glossary/document.xml",
        code: "paraid-missing-element",
    },
    {
        severity: "warning",
        message:
            "1 <w:p> element(s) lack a w14:paraId. Word stamps one on every paragraph and table row as the anchor for tracked-changes / comment-range infrastructure.",
        path: "word/glossary/footer9.xml",
        code: "paraid-missing-element",
    },
    {
        severity: "error",
        code: "mce-malformed",
        message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
        path: "word/glossary/_rels/document.xml.rels",
    },
    {
        severity: "error",
        message: "Unreferenced file: word/document.xml",
        path: "word/document.xml",
        code: "rels-unreferenced",
    },
    {
        severity: "error",
        message: "File with <document> root not declared in [Content_Types].xml",
        path: "word/document.xml",
        code: "ct-undeclared-part",
    },
    {
        severity: "error",
        code: "mce-must-understand",
        message: "Required namespace is not understood: foo",
        path: "word/_rels/document.xml.rels",
    },
    {
        severity: "error",
        code: "mce-malformed",
        message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
        path: "word/_rels/styles.xml.rels",
    },
    {
        severity: "warning",
        message:
            "1 <w:p> element(s) lack a w14:paraId. Word stamps one on every paragraph and table row as the anchor for tracked-changes / comment-range infrastructure.",
        path: "word/footer1.xml",
        code: "paraid-missing-element",
    },
    {
        severity: "error",
        code: "mce-malformed",
        message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
        path: "word/_rels/footer1.xml.rels",
    },
    {
        severity: "error",
        code: "mce-malformed",
        message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
        path: "word/_rels/doc2.xml.rels",
    },
    {
        severity: "error",
        code: "mce-must-understand",
        message: "Required namespace is not understood: foo",
        path: "word/_rels/doc2.xml.rels",
    },
    {
        severity: "error",
        code: "mce-must-understand",
        message: "Required namespace is not understood: foo",
        path: "word/glossary/_rels/document.xml.rels",
    },
    {
        severity: "error",
        message:
            "  - Element '{http://schemas.openxmlformats.org/package/2006/relationships}Relationships', attribute '{http://schemas.openxmlformats.org/markup-compatibility/2006}MustUnderstand': The attribute '{http://schemas.openxmlformats.org/markup-compatibility/200...",
        path: "_rels/.rels",
        code: "xsd-error",
    },
    {
        severity: "error",
        message:
            "  - Element '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', attribute '{http://schemas.openxmlformats.org/markup-compatibility/2006}MustUnderstand': The attribute '{http://schemas.openxmlformats.org/markup-compatibility/2006...",
        path: "_rels/.rels",
        code: "xsd-error",
    },
];
const lenientBaseline: Readonly<Record<string, { valid: boolean; indices: readonly number[] }>> = {
    "docrels-mc-malformed-sibling": {
        valid: false,
        indices: [0, 1],
    },
    "pkgrels-mc-malformed-sibling": {
        valid: false,
        indices: [0, 2, 3],
    },
    "glossary-rels-mc-malformed-sibling": {
        valid: false,
        indices: [4, 5, 0, 6],
    },
    "stray-docrels-mc-malformed-sibling": {
        valid: false,
        indices: [7, 8, 0, 1],
    },
    "docrels-root-mustunderstand": {
        valid: false,
        indices: [0, 9],
    },
    "styles-rels-mc-malformed-sibling": {
        valid: false,
        indices: [0, 10],
    },
    "footer-rels-mc-malformed-sibling": {
        valid: false,
        indices: [11, 0, 12],
    },
    "renamed-main-control": {
        valid: true,
        indices: [0],
    },
    "renamed-main-rels-mc-malformed-sibling": {
        valid: false,
        indices: [0, 13],
    },
    "renamed-main-rels-root-mustunderstand": {
        valid: false,
        indices: [0, 14],
    },
    "stray-docrels-root-mustunderstand": {
        valid: false,
        indices: [7, 8, 0, 9],
    },
    "glossary-rels-root-mustunderstand": {
        valid: false,
        indices: [4, 5, 0, 15],
    },
    packageRelsRootMustUnderstand: {
        valid: false,
        indices: [0, 2, 16],
    },
    documentRelsChoiceUndeclaredPrefix: {
        valid: false,
        indices: [11, 0, 1],
    },
    "pkgrels-relationship-mustunderstand": {
        valid: false,
        indices: [0, 2, 17],
    },
    "docrels-mc-fallback-only": {
        valid: false,
        indices: [0, 1],
    },
    "docrels-mc-foreign-branch": {
        valid: false,
        indices: [0, 1],
    },
};

describe("story-root relationships MCE classification (#644, session 36)", () => {
    for (const suffix of ["docx", "docm"]) {
        it.each(STORY_RELS_MCE_CASES)(
            `${suffix}: $label keeps the observed verdict across full, quick and named runs`,
            async ({ parts, code, part, blocks, label }) => {
                await withTempDir(async (dir) => {
                    const file = path.join(dir, `specimen.${suffix}`);
                    const bytes = await buildDocx(parts());
                    await writeFile(file, bytes);
                    let fullMce: ValidationIssue[] = [];
                    for (const selection of [{}, { quick: true }, { checks: [CODE] }]) {
                        const result = await validate(file, { profile: "word-valid", ...selection });
                        expect.soft(result.valid, JSON.stringify(selection)).toBe(!blocks);
                        const findings = mce(result.issues);
                        expect
                            .soft(findings, JSON.stringify(selection))
                            .toEqual(
                                code === undefined
                                    ? []
                                    : [expect.objectContaining({ code, path: part, severity: blocks ? "error" : "warning" })],
                            );
                        if (!("quick" in selection) && !("checks" in selection)) fullMce = findings;
                        else expect.soft(findings).toEqual(fullMce);
                    }
                    const strict = await validate(file, { profile: "strict" });
                    expect.soft(strict.valid).toBe(code === undefined);
                    expect
                        .soft(mce(strict.issues))
                        .toEqual(code === undefined ? [] : [expect.objectContaining({ code, path: part, severity: "error" })]);
                    const lenient = await validate(file, { profile: "lenient" });
                    const baseline = lenientBaseline[label]!;
                    expect.soft(JSON.stringify({ valid: lenient.valid, issues: lenient.issues })).toBe(
                        JSON.stringify({
                            valid: baseline.valid,
                            issues: baseline.indices.map((index) => lenientIssues[index]),
                        }),
                    );
                    expect(await readFile(file)).toEqual(bytes);
                });
            },
        );
    }
});

/**
 * #639: the judged set is compared as OPC compares part names, so an
 * OPC-equivalent spelling with an escaped unreserved character is judged as
 * its unescaped twin is, in full, quick and named runs (#381). Word was not
 * shown these escaped shapes: the package stays invalid through its stored
 * name's `opc-part-name-escaped-unreserved` rejection (#408), so no verdict
 * changes — only the findings do.
 */
describe("markup compatibility judging of OPC-equivalent word/*.xml part names (#639, #381)", () => {
    const HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    // `%78` is an escaped `x` (M1.8 forbids the escape; OPC still equates it).
    const MUST_UNDERSTAND_FOO = `${HEAD}<w:document xmlns:w="${NS.W}" xmlns:mc="${NS.MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"><w:body/></w:document>`;
    const IGNORABLE_UNDECLARED = `${HEAD}<w:document xmlns:w="${NS.W}" xmlns:mc="${NS.MC}" mc:Ignorable="undecl"><w:body/></w:document>`;
    const withPart = (part: string, xml: string): Record<string, string> => ({ ...MINIMAL_PARTS, [part]: xml });

    const specimens = [
        {
            label: "escaped-extension-must-understand",
            part: "word/foo.%78ml",
            parts: withPart("word/foo.%78ml", MUST_UNDERSTAND_FOO),
            code: "mce-must-understand",
            // No relationship names the part, so word-valid warns in every
            // spelling: Word session 84 (s84a, two runs, all readings agree)
            // opened the unescaped specimen clean (#1116). The escaped
            // spelling keeps its stored-name rejection.
            severity: "warning",
            fullValid: false,
            namedValid: true,
            quickValid: false,
            lenient: [
                "error:ct-undeclared-part@word/foo.%78ml",
                "error:mce-must-understand@word/foo.%78ml",
                "error:rels-unreferenced@word/foo.%78ml",
                "info:latent-styles-missing@word/styles.xml",
                "warning:opc-part-name-escaped-unreserved@word/foo.%78ml",
            ],
        },
        {
            label: "escaped-extension-ignorable-undeclared",
            part: "word/foo.%78ml",
            parts: withPart("word/foo.%78ml", IGNORABLE_UNDECLARED),
            code: "ignorable-undeclared",
            // Not Word-blocking, so word-valid keeps the warning the full
            // run's validateNamespaces already reports at this part.
            severity: "warning",
            fullValid: false,
            namedValid: true,
            quickValid: false,
            lenient: [
                "error:ct-undeclared-part@word/foo.%78ml",
                "error:ignorable-undeclared@word/foo.%78ml",
                "error:rels-unreferenced@word/foo.%78ml",
                "info:latent-styles-missing@word/styles.xml",
                "warning:opc-part-name-escaped-unreserved@word/foo.%78ml",
            ],
        },
        {
            label: "upper-case-must-understand",
            part: "word/FOO.XML",
            parts: withPart("word/FOO.XML", MUST_UNDERSTAND_FOO),
            code: "mce-must-understand",
            severity: "warning",
            fullValid: true,
            namedValid: true,
            quickValid: true,
            lenient: [
                "error:ct-undeclared-part@word/FOO.XML",
                "error:mce-must-understand@word/FOO.XML",
                "error:rels-unreferenced@word/FOO.XML",
                "info:latent-styles-missing@word/styles.xml",
            ],
        },
        {
            label: "upper-case-ignorable-undeclared",
            part: "word/FOO.XML",
            parts: withPart("word/FOO.XML", IGNORABLE_UNDECLARED),
            code: "ignorable-undeclared",
            severity: "warning",
            fullValid: true,
            namedValid: true,
            quickValid: true,
            lenient: [
                "error:ct-undeclared-part@word/FOO.XML",
                "error:ignorable-undeclared@word/FOO.XML",
                "error:rels-unreferenced@word/FOO.XML",
                "info:latent-styles-missing@word/styles.xml",
            ],
        },
        {
            label: "unescaped-must-understand",
            part: "word/foo.xml",
            parts: withPart("word/foo.xml", MUST_UNDERSTAND_FOO),
            code: "mce-must-understand",
            severity: "warning",
            fullValid: true,
            namedValid: true,
            quickValid: true,
            lenient: [
                "error:ct-undeclared-part@word/foo.xml",
                "error:mce-must-understand@word/foo.xml",
                "error:rels-unreferenced@word/foo.xml",
                "info:latent-styles-missing@word/styles.xml",
            ],
        },
        {
            label: "unescaped-ignorable-undeclared",
            part: "word/foo.xml",
            parts: withPart("word/foo.xml", IGNORABLE_UNDECLARED),
            code: "ignorable-undeclared",
            severity: "warning",
            fullValid: true,
            namedValid: true,
            quickValid: true,
            lenient: [
                "error:ct-undeclared-part@word/foo.xml",
                "error:ignorable-undeclared@word/foo.xml",
                "error:rels-unreferenced@word/foo.xml",
                "info:latent-styles-missing@word/styles.xml",
            ],
        },
    ] as const;

    for (const suffix of ["docx", "docm"] as const) {
        it.each(specimens)(`${suffix}: $label`, async ({ parts, part, code, severity, fullValid, namedValid, quickValid, lenient }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `specimen.${suffix}`);
                const bytes = await buildDocx(parts);
                await writeFile(file, bytes);

                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity, path: part }),
                ]);
                expect(wordValid.valid).toBe(fullValid);
                // The escaped spelling keeps its stored-name rejection, so
                // the verdict never rests on the new finding (#408, #639).
                if (part.includes("%")) {
                    expect(wordValid.issues.filter((issue) => issue.code === "opc-part-name-escaped-unreserved")).toEqual([
                        expect.objectContaining({ severity: "error", path: part }),
                    ]);
                }

                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.filter((issue) => issue.code === code)).toEqual([
                    expect.objectContaining({ code, severity: "error", path: part }),
                ]);
                expect(strict.valid).toBe(false);

                const selections = [
                    { quick: true, expectedValid: quickValid },
                    { checks: [CODE], expectedValid: namedValid },
                ] as const;
                for (const { expectedValid, ...selection } of selections) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    const label = JSON.stringify(selection);
                    expect.soft(partial.valid, label).toBe(expectedValid);
                    // The named run reports the full run's finding through the
                    // shared predicate, and the root-Ignorable pairing (#499)
                    // keeps the same identity and count.
                    expect
                        .soft(
                            partial.issues.filter((issue) => issue.code === code),
                            label,
                        )
                        .toEqual(wordValid.issues.filter((issue) => issue.code === code));
                    expect.soft(mce(partial.issues), label).toEqual(mce(wordValid.issues));
                    expect
                        .soft(
                            [...errors(partial.issues)].filter((pair) => !errors(wordValid.issues).has(pair)),
                            label,
                        )
                        .toEqual([]);
                }

                // Lenient runs no Word rule, so its findings stay as they were.
                const lenientRun = await validate(file, { profile: "lenient" });
                expect(lenientRun.issues.map((issue) => `${issue.severity}:${issue.code}@${issue.path ?? ""}`).sort()).toEqual([
                    ...lenient,
                ]);
                expect(await readFile(file)).toEqual(bytes);
            });
        });
    }
});

/**
 * #851: Word 16.115 offered recovery for character data directly inside an
 * `mc:AlternateContent` — before the Choice, between the Choice and the
 * Fallback (session 65, s65c, two runs, all four readings agreeing), after
 * the Fallback, after a Choice with no Fallback, and between two Choices
 * (session 67, s67b, likewise) — in a used `word/comments.xml` typed
 * comments whose stored root is the AlternateContent, and in a run-level
 * body AlternateContent (session 67, s67b). The same roots open clean with
 * the text inside the Choice or the Fallback (session 65, s65b), with
 * whitespace only between the branches, and without the text (session 67,
 * s67b), so the rule is non-whitespace character data anywhere directly
 * inside the element: whitespace is note 8's `[ \t\n\r]`, which
 * pretty-printed parts carry. `mce/mc.xsd` declares none of
 * AlternateContent, Choice and Fallback mixed, so the schema view forbids
 * all three positions alike and cannot draw Word's line. AlternateContent
 * in parts other than the main document and comments, and in relationships
 * parts, is unprobed.
 */
describe("character data directly inside AlternateContent (#851, sessions 65 and 67)", () => {
    const COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    const COMMENT = `<w:comment w:id="0" w:author="a"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:comment>`;
    const CHOICE = `<mc:Choice Requires="w14"><w:comments xmlns:w="${NS.W}">${COMMENT}</w:comments></mc:Choice>`;
    const UNKNOWN_CHOICE = `<mc:Choice Requires="u"><w:comments xmlns:w="${NS.W}">${COMMENT}</w:comments></mc:Choice>`;
    const FALLBACK = "<mc:Fallback/>";
    const COMMENT_REF = `<w:r><w:commentReference w:id="0"/></w:r>`;
    /** A used `word/comments.xml` typed comments whose stored root is an AlternateContent holding `inner`; the body references comment 0. */
    const commentsAlternate = (inner: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            `xmlns:w="${NS.W}"`,
            `xmlns:w="${NS.W}" xmlns:r="${NS.R}"`,
        ).replace("</w:r></w:p>", `</w:r>${COMMENT_REF}</w:p>`),
        "word/comments.xml":
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<mc:AlternateContent xmlns:mc="${NS.MC}" xmlns:w14="${NS.W14}" xmlns:u="urn:example:unknown">${inner}</mc:AlternateContent>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdC" Type="${NS.R}/comments" Target="comments.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"].replace(
            "</Types>",
            `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/></Types>`,
        ),
    });
    /** The main document's paragraph holding a run-level AlternateContent with `between` between its Choice run and its Fallback run. */
    const bodyAlternate = (between: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            "</w:r></w:p>",
            `</w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>Chosen</w:t></w:r></mc:Choice>${between}` +
                `<mc:Fallback><w:r><w:t>Fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
        ),
    });
    // Every position the sessions showed Word refusing, in the comments
    // part's stored root: sessions 65 (s65c) and 67 (s67b).
    const refused = [
        ["text before the Choice", `Stray${CHOICE}${FALLBACK}`],
        ["text between the Choice and the Fallback", `${CHOICE}Stray${FALLBACK}`],
        ["text after the Fallback", `${CHOICE}${FALLBACK}Stray`],
        ["text after a Choice with no Fallback", `${CHOICE}Stray`],
        ["text between two Choices", `${CHOICE}Stray${UNKNOWN_CHOICE}${FALLBACK}`],
    ] as const;
    // The calibrations the same sessions opened clean: the text inside a
    // branch (session 65, s65b), whitespace only between the branches, and
    // the same roots without any text (session 67, s67b).
    const calibrations = [
        ["text inside the Choice before the root", CHOICE.replace("<w:comments", "Stray<w:comments") + FALLBACK],
        ["text inside the Fallback", `${CHOICE}<mc:Fallback>Stray</mc:Fallback>`],
        ["whitespace only between the branches", `${CHOICE}\n  ${FALLBACK}`],
        ["a Choice with no Fallback and no text", CHOICE],
        ["two Choices and no text", `${CHOICE}${UNKNOWN_CHOICE}${FALLBACK}`],
    ] as const;

    it.each(refused)(
        "rejects a used comments part whose root AlternateContent holds %s",
        async (_shape, inner) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                const bytes = await buildDocx(commentsAlternate(inner));
                await writeFile(file, bytes);

                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([
                    expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/comments.xml" }),
                ]);
                expect(wordValid.valid).toBe(false);

                // Quick and named runs report the full run's finding (#381).
                for (const selection of [{ quick: true }, { checks: [CODE] }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect.soft(partial.valid, JSON.stringify(selection)).toBe(false);
                    expect.soft(mce(partial.issues), JSON.stringify(selection)).toEqual(mce(wordValid.issues));
                    expect
                        .soft(
                            [...errors(partial.issues)].filter((pair) => !errors(wordValid.issues).has(pair)),
                            JSON.stringify(selection),
                        )
                        .toEqual([]);
                }

                // Strict keeps its schema-view findings beside the new one: its
                // compatibility view does not understand w14, so the empty or
                // absent Fallback leaves the root unwrapping to no element.
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid).toBe(false);
                expect(strict.issues.filter((issue) => issue.code === "mce-malformed")).toContainEqual(
                    expect.objectContaining({ severity: "error", path: "word/comments.xml" }),
                );
                expect(await readFile(file)).toEqual(bytes);
            });
        },
        60_000,
    );

    // #1016: sessions 65 and 67 stored Transitional packages, so every
    // profile drops their citation beside a Strict part.
    it.each(["word-valid", "lenient", "strict"] as const)(
        "cites sessions 65 and 67 under %s only in a Transitional package (#1016)",
        async (profile) => {
            await withTempDir(async (dir) => {
                for (const [parts, message] of [
                    [
                        bodyAlternate("Stray"),
                        "AlternateContent holds character data directly inside it; Word 16.115 offered recovery for every position probed (sessions 65 and 67).",
                    ],
                    [
                        {
                            ...bodyAlternate("Stray"),
                            "docProps/app.xml": '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"/>',
                        },
                        "AlternateContent holds character data directly inside it.",
                    ],
                ] as const) {
                    const file = path.join(dir, "specimen.docx");
                    await writeFile(file, await buildDocx(parts));
                    const result = await validate(file, { profile });
                    expect(result.valid).toBe(false);
                    expect(
                        result.issues.filter((issue) => issue.code === "mce-malformed").map((issue) => [issue.path, issue.message]),
                    ).toEqual([["word/document.xml", message]]);
                }
            });
        },
        60_000,
    );

    // The citation is decided per package, not per finding: another
    // mce-malformed beside the character data keeps it, and keeps its own
    // message (#1030).
    it.each(["word-valid", "lenient"] as const)(
        "cites sessions 65 and 67 under %s beside another mce-malformed finding (#1030)",
        async (profile) => {
            await withTempDir(async (dir) => {
                const parts = {
                    ...bodyAlternate("Stray"),
                    "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace(
                        "</w:styles>",
                        `<mc:AlternateContent xmlns:mc="${NS.MC}"><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></w:styles>`,
                    ),
                };
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(parts));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.code === "mce-malformed").map((issue) => [issue.path, issue.message])).toEqual(
                    [
                        [
                            "word/document.xml",
                            "AlternateContent holds character data directly inside it; Word 16.115 offered recovery for every position probed (sessions 65 and 67).",
                        ],
                        [
                            "word/styles.xml",
                            "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
                        ],
                    ],
                );
            });
        },
        60_000,
    );

    it.each(calibrations)(
        "keeps the used comments part clean with %s",
        async (_shape, inner) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                await writeFile(file, await buildDocx(commentsAlternate(inner)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([]);
                expect(result.valid).toBe(true);
            });
        },
        60_000,
    );

    it("rejects a body run-level AlternateContent holding text between its Choice and its Fallback (session 67 s67b)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            const bytes = await buildDocx(bodyAlternate("Stray"));
            await writeFile(file, bytes);

            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([
                expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
            ]);
            expect(wordValid.valid).toBe(false);

            for (const selection of [{ quick: true }, { checks: [CODE] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                expect.soft(partial.valid, JSON.stringify(selection)).toBe(false);
                expect.soft(mce(partial.issues), JSON.stringify(selection)).toEqual(mce(wordValid.issues));
                expect
                    .soft(
                        [...errors(partial.issues)].filter((pair) => !errors(wordValid.issues).has(pair)),
                        JSON.stringify(selection),
                    )
                    .toEqual([]);
            }

            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([
                expect.objectContaining({ severity: "error", path: "word/document.xml" }),
            ]);
            expect(strict.valid).toBe(false);
            expect(await readFile(file)).toEqual(bytes);
        });
    }, 60_000);

    it("keeps the body calibration without the text clean (session 67 s67b)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(bodyAlternate("")));
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([]);
            expect(wordValid.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === "mce-malformed")).toEqual([]);
            expect(strict.valid).toBe(true);
        });
    }, 60_000);
});
