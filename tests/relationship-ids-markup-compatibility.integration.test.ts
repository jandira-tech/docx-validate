/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import {
    DOCUMENT_RELS,
    MCE_DECLARATIONS,
    PACKAGE_RELS,
    RELATIONSHIPS_MCE_CASES,
    fallback,
    wrapRelationship,
} from "./support/relationships-markup-compatibility-specimens";
import { STORY_FOOTER_XML, storyContentType, storyReferenceParts } from "./support/story-content-type-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "word/document.xml";
const footer = RELATIONSHIPS_MCE_CASES[0].parts;

/** Session 34 s34a's package duplicate and direct twin; the sidecar extension is unobserved in Word. */
const duplicate = (part: string, wrapped: boolean): Record<string, string> => {
    const added = `<Relationship Id="rId1" Type="${R}/styles" Target="${part === PACKAGE_RELS ? "word/styles.xml" : "styles.xml"}"/>`;
    const stored = MINIMAL_PARTS[part]!;
    return {
        ...MINIMAL_PARTS,
        [part]: (wrapped ? stored.replace("<Relationships ", `<Relationships ${MCE_DECLARATIONS} `) : stored).replace(
            "</Relationships>",
            `${wrapped ? fallback(added) : added}</Relationships>`,
        ),
    };
};
const DUPLICATES = [
    { label: "mce-fallback-split-duplicate-id", part: PACKAGE_RELS, parts: duplicate(PACKAGE_RELS, true) },
    { label: "direct-duplicate-id", part: PACKAGE_RELS, parts: duplicate(PACKAGE_RELS, false) },
    { label: "mce-fallback-split-sidecar-duplicate-id", part: DOCUMENT_RELS, parts: duplicate(DOCUMENT_RELS, true) },
    { label: "direct-sidecar-duplicate-id", part: DOCUMENT_RELS, parts: duplicate(DOCUMENT_RELS, false) },
];

// Use session 26's direct footer shape, then wrap only its relationship.
const directFooter = storyReferenceParts("footerReference", "footer", "footer1.xml", STORY_FOOTER_XML, storyContentType("footer"));
const unselected = wrapRelationship(
    directFooter,
    DOCUMENT_RELS,
    "rIdX",
    (rel) => `<mc:AlternateContent><mc:Choice Requires="foo">${rel}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
);
const mismatch = wrapRelationship(directFooter, DOCUMENT_RELS, "rIdX", (rel) => fallback(rel.replace(`${R}/footer`, `${R}/header`)));
const malformed = { ...footer, [DOCUMENT_RELS]: "<Relationships" };
const SPECIMENS = [
    ...DUPLICATES,
    ...RELATIONSHIPS_MCE_CASES,
    { label: "unselected-choice-footer", parts: unselected },
    { label: "selected-fallback-type-mismatch", parts: mismatch },
    { label: "malformed-sidecar", parts: malformed },
];
/**
 * Lenient keeps its raw traversal: each specimen's lenient verdict and issues
 * at f166c0a1, before #635 (severity, code and path; the message too for the
 * relationship walk's own codes).
 */
const LENIENT_BEFORE: Record<string, { valid: boolean; issues: string[] }> = {
    "mce-fallback-split-duplicate-id": {
        valid: false,
        issues: ["info latent-styles-missing@word/styles.xml", "error xsd-summary@_rels/.rels", "error xsd-error@_rels/.rels"],
    },
    "direct-duplicate-id": {
        valid: false,
        issues: ["info latent-styles-missing@word/styles.xml", "error xsd-summary@_rels/.rels", "error xsd-error@_rels/.rels"],
    },
    "mce-fallback-split-sidecar-duplicate-id": {
        valid: false,
        issues: [
            "error rels-id-duplicate@word/_rels/document.xml.rels: Duplicate relationship ID 'rId1' (IDs must be unique)",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/_rels/document.xml.rels",
            "error xsd-error@word/_rels/document.xml.rels",
        ],
    },
    "direct-sidecar-duplicate-id": {
        valid: false,
        issues: [
            "error rels-id-duplicate@word/_rels/document.xml.rels: Duplicate relationship ID 'rId1' (IDs must be unique)",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/_rels/document.xml.rels",
            "error xsd-error@word/_rels/document.xml.rels",
        ],
    },
    "mce-rels-fallback-footer": {
        valid: true,
        issues: ["warning paraid-missing-element@word/footer1.xml", "info latent-styles-missing@word/styles.xml"],
    },
    "mce-rels-choice-w14-footer": {
        valid: true,
        issues: ["warning paraid-missing-element@word/footer1.xml", "info latent-styles-missing@word/styles.xml"],
    },
    "mce-rels-ignorable-foreign": {
        valid: true,
        issues: ["warning paraid-missing-element@word/footer1.xml", "info latent-styles-missing@word/styles.xml"],
    },
    "mce-package-rels-fallback-main": {
        valid: false,
        issues: ["info latent-styles-missing@word/styles.xml", "error xsd-summary@_rels/.rels", "error xsd-error@_rels/.rels"],
    },
    "mce-rels-fallback-foreign": {
        valid: false,
        issues: [
            "warning paraid-missing-element@word/footer1.xml",
            "info latent-styles-missing@word/styles.xml",
            "error xsd-summary@word/_rels/document.xml.rels",
            "error xsd-error@word/_rels/document.xml.rels",
        ],
    },
    "unselected-choice-footer": {
        valid: true,
        issues: ["warning paraid-missing-element@word/footer1.xml", "info latent-styles-missing@word/styles.xml"],
    },
    "selected-fallback-type-mismatch": {
        valid: false,
        issues: [
            "error rels-id-mismatch@word/document.xml: <footerReference> references 'rIdX' which points to 'header' but should point to a 'footer' relationship",
            "warning paraid-missing-element@word/footer1.xml",
            "info latent-styles-missing@word/styles.xml",
        ],
    },
    "malformed-sidecar": {
        valid: false,
        issues: ["error xml-syntax@word/_rels/document.xml.rels"],
    },
};
const lenientKey = (issue: ValidationIssue): string =>
    `${issue.severity} ${issue.code}@${issue.path ?? ""}` + (issue.code?.startsWith("rels-") ? `: ${issue.message}` : "");
const matching = (issues: readonly ValidationIssue[], ...codes: string[]): ValidationIssue[] =>
    issues.filter((issue) => codes.includes(issue.code ?? ""));
const errorKeys = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`);
const writeParts = async (dir: string, parts: Readonly<Record<string, string>>): Promise<void> => {
    for (const [part, xml] of Object.entries(parts)) {
        const file = path.join(dir, part);
        await mkdir(path.dirname(file), { recursive: true });
        await writeFile(file, xml);
    }
};

describe("relationship walks after markup compatibility (#635)", { timeout: 120_000 }, () => {
    it.each(DUPLICATES)("$label: reports the duplicate at the relationships part under both profiles", async ({ parts, part }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const full = await validate(file, { profile });
                expect.soft(full.valid, profile).toBe(false);
                expect
                    .soft(matching(full.issues, "rels-id-duplicate"), profile)
                    .toEqual([expect.objectContaining({ code: "rels-id-duplicate", severity: "error", path: part })]);
            }
        });
    });

    it.each(RELATIONSHIPS_MCE_CASES.filter(({ clean }) => clean))(
        "$label: selected targets are neither missing nor unreferenced",
        async ({ parts }) => {
            await withTempDir(async (dir) => {
                const target = path.join(dir, "specimen.docx");
                await writeParts(target, parts);
                for (const profile of ["word-valid", "strict"] as const) {
                    const full = await validate(target, { profile });
                    expect.soft(matching(full.issues, "rels-id-missing", "rels-unreferenced"), profile).toEqual([]);
                    expect.soft(matching(full.issues, "word-header-footer-relationship"), profile).toEqual([]);
                    if (profile === "word-valid") expect.soft(full.valid).toBe(true);
                }
            });
        },
    );

    it.each(["word-valid", "strict"] as const)(
        "%s: unselected ids stay missing and the identity hand-off drops only the copy",
        async (profile) => {
            await withTempDir(async (dir) => {
                const target = path.join(dir, "specimen.docx");
                await writeParts(target, unselected);
                const validator = new DOCXSchemaValidator({ unpackedDir: target, profile });
                const ids = await validator.validateAllRelationshipIds();
                expect(matching(ids.issues, "rels-id-missing")).toEqual([
                    expect.objectContaining({ path: MAIN, severity: "error", message: expect.stringContaining("'rIdX'") }),
                ]);
                const references = await validator.validateFileReferences();
                expect(matching(references.issues, "rels-unreferenced")).toEqual([
                    expect.objectContaining({ path: "word/footer1.xml", severity: "error" }),
                ]);
                const full = await validate(target, { profile });
                expect(full.valid).toBe(false);
                expect(matching(full.issues, "word-header-footer-relationship")).toEqual([
                    expect.objectContaining({ path: MAIN, severity: "error" }),
                ]);
                expect(matching(full.issues, "rels-id-missing")).toEqual([]);
            });
        },
    );

    it.each(["word-valid", "strict"] as const)("%s: selected relationship types reach the legacy mismatch check", async (profile) => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "specimen.docx");
            await writeParts(target, mismatch);
            const ids = await new DOCXSchemaValidator({ unpackedDir: target, profile }).validateAllRelationshipIds();
            expect(matching(ids.issues, "rels-id-missing")).toEqual([]);
            expect(matching(ids.issues, "rels-id-mismatch")).toEqual([
                expect.objectContaining({ path: MAIN, severity: "error", message: expect.stringContaining("'header'") }),
            ]);
            const full = await validate(target, { profile });
            expect(full.valid).toBe(false);
            expect(matching(full.issues, "word-header-footer-relationship")).toHaveLength(1);
            expect(matching(full.issues, "rels-id-missing", "rels-id-mismatch")).toEqual([]);
        });
    });

    it.each(["word-valid", "strict"] as const)("%s: malformed sidecars retain rels-id-parse", async (profile) => {
        await withTempDir(async (dir) => {
            await writeParts(dir, malformed);
            const ids = await new DOCXSchemaValidator({ unpackedDir: dir, profile }).validateAllRelationshipIds();
            expect(matching(ids.issues, "rels-id-parse")).toEqual([expect.objectContaining({ path: MAIN, severity: "error" })]);
        });
    });

    it.each(SPECIMENS)("$label: full, quick and named findings agree on their shared checks", async ({ parts }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts));
            const full = await validate(file, { profile: "word-valid" });
            for (const selection of [{ quick: true }, { checks: ["relationships"] }]) {
                const partial = await validate(file, { profile: "word-valid", ...selection });
                for (const code of [
                    "opc-relationships-structure",
                    "opc-relationships-markup-compatibility",
                    "word-header-footer-relationship",
                ]) {
                    expect.soft(matching(partial.issues, code), code).toEqual(matching(full.issues, code));
                }
                expect.soft(errorKeys(partial.issues).filter((key) => !errorKeys(full.issues).includes(key))).toEqual([]);
                if (full.valid) expect.soft(partial.valid).toBe(true);
                // These legacy walks remain outside the named-check registry, as #615 specifies.
                expect
                    .soft(matching(partial.issues, "rels-id-duplicate", "rels-id-missing", "rels-id-mismatch", "rels-unreferenced"))
                    .toEqual([]);
            }
        });
    });

    it("leaves every lenient result as it was before the change", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            for (const { label, parts } of SPECIMENS) {
                await writeFile(file, await buildDocx(parts));
                const result = await validate(file, { profile: "lenient" });
                expect.soft({ valid: result.valid, issues: result.issues.map(lenientKey) }, label).toEqual(LENIENT_BEFORE[label]);
            }
        });
    });
});
