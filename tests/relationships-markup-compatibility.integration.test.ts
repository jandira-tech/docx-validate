/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Reachability through relationships markup compatibility selects (#634).
 *
 * Word 16.115 (session 34, s34a/s34b in
 * docs/plans/evidence/word-probes-2026-10-24.md) refuses or offers to recover
 * a defect in a part reached through a relationship inside an `mc:Fallback` —
 * or an understood `mc:Choice` — exactly as it does through a direct
 * relationship: a truncated main document reached through `_rels/.rels`'s
 * Fallback officeDocument relationship gets an open error, a zero-byte
 * `word/_rels/document.xml.rels` beside that Fallback main document gets a
 * recovery offer, a truncated `docProps/core.xml` reached through a Fallback
 * core-properties relationship gets a recovery offer, and session 26's
 * footer typed as a header (#600) with its relationship in the Fallback of
 * `word/_rels/document.xml.rels` gets a recovery offer. Each wrapped specimen
 * runs beside its direct-relationship twin, which Word refuses identically.
 *
 * At ef5d8107 word-valid passed every wrapped shape: the reachability
 * readers (`wordReadParts`, `wordStoryRootParts`, `propertiesParts` and the
 * typed-target walks behind them) read the stored relationships part, where
 * the selected relationship is not a direct `Relationship` child. The tests
 * below pin the processed view under word-valid and strict, selection parity
 * with the full run (#381), byte-identical lenient output, and session 33's
 * clean shapes.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import {
    DOCUMENT_RELS,
    fallback,
    MCE_DECLARATIONS,
    PACKAGE_RELS,
    RELATIONSHIPS_MCE_CASES,
    wrapRelationship,
} from "./support/relationships-markup-compatibility-specimens";
import { STORY_FOOTER_XML, storyContentType, storyReferenceParts } from "./support/story-content-type-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml";

type Parts = Record<string, string>;

/** Session 34's truncated main document (s34a). */
const MALFORMED_MAIN = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>`;
/** Session 34's truncated core-properties part (s34b). */
const MALFORMED_CORE = `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">broken`;
const CORE_REL = `<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`;
const OFFICE_REL = `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`;
const relsWith = (body: string, declarations = ""): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"${declarations ? ` ${declarations}` : ""}>${body}</Relationships>`;
const withCore = (rels: string): Parts => ({
    ...MINIMAL_PARTS,
    [PACKAGE_RELS]: rels,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/docProps/core.xml" ContentType="${CT_CORE}"/></Types>`,
    ),
    "docProps/core.xml": MALFORMED_CORE,
});
const footer = (ct: "footer" | "header"): Parts =>
    storyReferenceParts("footerReference", "footer", "footer1.xml", STORY_FOOTER_XML, storyContentType(ct));
const mainInFallback = (): Record<string, string> => wrapRelationship({ ...MINIMAL_PARTS }, PACKAGE_RELS, "rId1", fallback);

/**
 * The lenient findings each wrapped shape reported at ef5d8107, pinned so the
 * fix cannot touch lenient output: lenient reads the stored relationships
 * part, sees no selected relationship, and reports what it always reported.
 */
const LENIENT_ISSUES: Readonly<Record<string, readonly Record<string, unknown>[]>> = {
    "mce-fallback-main-malformed-doc": [
        { severity: "error", code: "xml-syntax", path: "word/document.xml", message: "Premature end of data in tag p line 1" },
    ],
    "mce-fallback-main-zerobyte-rels": [
        { severity: "error", code: "xml-syntax", path: "word/_rels/document.xml.rels", message: "Document is empty" },
    ],
    "mce-fallback-malformed-coreprops": [
        { severity: "error", code: "xml-syntax", path: "docProps/core.xml", message: "Premature end of data in tag title line 1" },
    ],
    "mce-fallback-footer-header-ct": [
        {
            severity: "warning",
            code: "paraid-missing-element",
            path: "word/footer1.xml",
            message:
                "1 <w:p> element(s) lack a w14:paraId. Word stamps one on every paragraph and table row as the anchor for tracked-changes / comment-range infrastructure.",
        },
        {
            severity: "info",
            code: "latent-styles-missing",
            path: "word/styles.xml",
            message:
                "word/styles.xml has no <w:latentStyles> block. This is unusual for Word-generated files and indicates the repairer may have regenerated styles from scratch.",
        },
    ],
};

const SHAPES = [
    {
        label: "mce-fallback-main-malformed-doc",
        code: "xml-syntax",
        part: "word/document.xml",
        named: "xml-syntax",
        wrapped: (): Parts => ({ ...mainInFallback(), "word/document.xml": MALFORMED_MAIN }),
        direct: (): Parts => ({ ...MINIMAL_PARTS, "word/document.xml": MALFORMED_MAIN }),
    },
    {
        label: "mce-fallback-main-zerobyte-rels",
        code: "xml-syntax",
        part: "word/_rels/document.xml.rels",
        named: "xml-syntax",
        wrapped: (): Parts => ({ ...mainInFallback(), [DOCUMENT_RELS]: "" }),
        direct: (): Parts => ({ ...MINIMAL_PARTS, [DOCUMENT_RELS]: "" }),
    },
    {
        label: "mce-fallback-malformed-coreprops",
        code: "xml-syntax",
        part: "docProps/core.xml",
        named: "xml-syntax",
        wrapped: (): Parts => withCore(relsWith(OFFICE_REL + fallback(CORE_REL), MCE_DECLARATIONS)),
        direct: (): Parts => withCore(relsWith(OFFICE_REL + CORE_REL)),
    },
    {
        label: "mce-fallback-footer-header-ct",
        code: "word-part-content-type-mismatch",
        part: "word/footer1.xml",
        named: "word-part-content-type-mismatch",
        wrapped: (): Parts => wrapRelationship(footer("header"), DOCUMENT_RELS, "rIdX", fallback),
        direct: (): Parts => footer("header"),
    },
] as const;

const errors = (issues: readonly ValidationIssue[]): Set<string> =>
    new Set(issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`));
const pinned = (issues: readonly ValidationIssue[]) =>
    issues.map((issue) => ({ severity: issue.severity, code: issue.code, path: issue.path, message: issue.message }));

describe("reachability through relationships markup compatibility (#634)", () => {
    it.each(SHAPES)("word-valid blocks $label like its direct twin (session 34)", { timeout: 120_000 }, async (shape) => {
        await withTempDir(async (dir) => {
            const wrappedFile = path.join(dir, `${shape.label}.docx`);
            const directFile = path.join(dir, `direct-${shape.label.replace(/^mce-fallback-/, "")}.docx`);
            const wrappedBytes = await buildDocx(shape.wrapped());
            const directBytes = await buildDocx(shape.direct());
            writeFileSync(wrappedFile, wrappedBytes);
            writeFileSync(directFile, directBytes);

            for (const [file, label] of [
                [wrappedFile, shape.label],
                [directFile, `${shape.label} twin`],
            ] as const) {
                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.valid, `${label} word-valid`).toBe(false);
                expect(
                    wordValid.issues.filter((issue) => issue.code === shape.code && issue.path === shape.part),
                    `${label} word-valid finding`,
                ).toEqual([expect.objectContaining({ severity: "error" })]);
                const strict = await validate(file, { profile: "strict" });
                expect(strict.valid, `${label} strict`).toBe(false);
            }

            // Quick and named runs agree with the full run for the shared
            // blocking finding (#381): same error, no errors the full run
            // does not have.
            const full = await validate(wrappedFile, { profile: "word-valid" });
            for (const selection of [{ quick: true }, { checks: [shape.named] }] as const) {
                const partial = await validate(wrappedFile, { profile: "word-valid", ...selection });
                const tag = `${shape.label} ${JSON.stringify(selection)}`;
                expect(
                    partial.issues.filter((issue) => issue.code === shape.code && issue.path === shape.part),
                    tag,
                ).toEqual([expect.objectContaining({ severity: "error" })]);
                for (const error of errors(partial.issues)) {
                    expect(errors(full.issues), tag).toContain(error);
                }
            }

            // Lenient keeps its stored-view findings, byte for byte.
            const lenient = await validate(wrappedFile, { profile: "lenient" });
            expect(pinned(lenient.issues), `${shape.label} lenient`).toEqual(LENIENT_ISSUES[shape.label]);

            expect(readFileSync(wrappedFile)).toEqual(wrappedBytes);
            expect(readFileSync(directFile)).toEqual(directBytes);
        });
    });

    it.each(RELATIONSHIPS_MCE_CASES.filter((specimen) => specimen.clean))(
        "word-valid stays clean on session 33's $label",
        { timeout: 120_000 },
        async ({ parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
                expect(result.valid).toBe(true);
            });
        },
    );

    it("keeps an unpacked directory agreeing with the packed file", { timeout: 120_000 }, async () => {
        await withTempDir(async (dir) => {
            const shape = SHAPES[0]!;
            const directory = path.join(dir, "unpacked.docx");
            for (const [part, content] of Object.entries(shape.wrapped())) {
                mkdirSync(path.dirname(path.join(directory, part)), { recursive: true });
                writeFileSync(path.join(directory, part), content);
            }
            const result = await validate(directory, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.code === shape.code && issue.path === shape.part)).toEqual([
                expect.objectContaining({ severity: "error" }),
            ]);
        });
    });
});
