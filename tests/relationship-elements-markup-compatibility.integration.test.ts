/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The legacy relationship walks process markup compatibility for DOCX/DOCM
 * only (#643): every observation behind #632–#636 is Word 16.115 opening
 * `.docx` files (probe sessions 33–35 in
 * docs/plans/evidence/word-probes-2026-10-24.md). PowerPoint was never shown
 * a relationships part with markup compatibility, so a PPTX keeps the stored
 * relationships view under every profile — f166c0a1's verdicts — while the
 * DOCX specimens stay processed.
 *
 * A DOCX run processes each relationships part once (#642): the walks share
 * the run context's cached `effectiveXml` view rather than re-running markup
 * compatibility per call.
 */

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";

import { isRelationshipsPartName } from "../src/lib/content-types";
import * as markupCompatibility from "../src/lib/markup-compatibility";
import { RELATIONSHIPS_UNDERSTOOD_NAMESPACES } from "../src/lib/markup-compatibility";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";
import { MCE_DECLARATIONS, fallback, wrapRelationship } from "./support/relationships-markup-compatibility-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PPTX_RELS = "ppt/_rels/presentation.xml.rels";

/** The PPTX specimen of #643: the slide relationship behind an unselected Choice's Fallback. */
const PPTX_SPECIMEN: Record<string, string> = {
    ...MINIMAL_PPTX_PARTS,
    [PPTX_RELS]:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" ` +
        `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:foo="urn:foo">` +
        `<Relationship Id="rId1" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<mc:AlternateContent><mc:Choice Requires="foo"/><mc:Fallback>` +
        `<Relationship Id="rId2" Type="${R}/slide" Target="slides/slide1.xml"/>` +
        `</mc:Fallback></mc:AlternateContent></Relationships>`,
};

const pinned = (issues: readonly ValidationIssue[]) =>
    issues.map((issue) => ({ severity: issue.severity, code: issue.code, path: issue.path, message: issue.message }));
const matching = (issues: readonly ValidationIssue[], ...codes: string[]): ValidationIssue[] =>
    issues.filter((issue) => codes.includes(issue.code ?? ""));

afterEach(() => {
    vi.restoreAllMocks();
});

describe("PPTX relationships parts keep the stored view under every profile (#643)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "strict"] as const)(
        "%s: a slide relationship behind mc:Fallback stays unread — f166c0a1's verdict",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.pptx");
                await writeFile(file, await buildPptx(PPTX_SPECIMEN));
                const result = await validate(file, { profile });
                expect(result.valid).toBe(false);
                expect(matching(result.issues, "rels-id-missing")).toEqual([
                    expect.objectContaining({ severity: "error", path: "ppt/presentation.xml" }),
                ]);
                expect(matching(result.issues, "rels-unreferenced")).toEqual([
                    expect.objectContaining({ severity: "error", path: "ppt/slides/slide1.xml" }),
                ]);
            });
        },
    );

    it("lenient keeps its byte-identical output: the descendant scan reads the Fallback relationship", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.pptx");
            await writeFile(file, await buildPptx(PPTX_SPECIMEN));
            const result = await validate(file, { profile: "lenient" });
            expect(result.valid).toBe(true);
            expect(pinned(result.issues)).toEqual([]);
        });
    });

    it("DOCX keeps the processed view: session 34 s34a's split duplicate still blocks", async () => {
        await withTempDir(async (dir) => {
            const added = `<Relationship Id="rId1" Type="${R}/styles" Target="word/styles.xml"/>`;
            const parts = {
                ...MINIMAL_PARTS,
                "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace("<Relationships ", `<Relationships ${MCE_DECLARATIONS} `).replace(
                    "</Relationships>",
                    `${fallback(added)}</Relationships>`,
                ),
            };
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.valid, profile).toBe(false);
                expect(matching(result.issues, "rels-id-duplicate"), profile).toEqual([
                    expect.objectContaining({ severity: "error", path: "_rels/.rels" }),
                ]);
            }
        });
    });

    it("DOCX keeps the processed view: session 33 s33d's Fallback footer stays clean", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(wrapRelationship(MINIMAL_PARTS, "_rels/.rels", "rId1", fallback)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.valid).toBe(true);
        });
    });
});

describe("one markup-compatibility pass per relationships part (#642)", { timeout: 120_000 }, () => {
    // A third relationships part beside MINIMAL_PARTS' two: an external-only
    // sidecar for a customXml part that is not in the package, so it adds no
    // finding and the package stays clean.
    const parts: Record<string, string> = {
        ...MINIMAL_PARTS,
        "customXml/_rels/item1.xml.rels":
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
            `</Relationships>`,
    };
    const relsParts = Object.keys(parts).filter(isRelationshipsPartName);

    it.each(["word-valid", "strict"] as const)("%s: a full run processes each relationships part once", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            await writeFile(file, await buildDocx(parts));
            const spy = vi.spyOn(markupCompatibility, "applyMarkupCompatibility");
            const result = await validate(file, { profile });
            expect(result.valid, JSON.stringify(result.issues)).toBe(true);
            const relationshipsPasses = spy.mock.calls.filter(([, opts]) => opts.understood === RELATIONSHIPS_UNDERSTOOD_NAMESPACES);
            expect(relationshipsPasses).toHaveLength(relsParts.length);
        });
    });
});

/**
 * #654: outside a `validate()` run a check method called directly reads the
 * package as it is now, as `context()` does. A stored `c:dummy.bin` makes the
 * shared context refuse the directory, so the walks take the processed
 * clone; in a readable directory they read `context().effectiveXml`, which
 * outside a run is a fresh context per call (#659). `repairIgnorable()` then
 * declares `w14`, which moves markup compatibility from the Fallback (no
 * `Target`) to the complete Choice.
 */
describe("the processed relationships view outside a run (#654)", { timeout: 120_000 }, () => {
    const SIDECAR = "word/_rels/document.xml.rels";
    const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const readable: Record<string, string> = {
        ...MINIMAL_PARTS,
        [SIDECAR]:
            `<Relationships xmlns="${PR}" xmlns:mc="${MC}" mc:Ignorable="w14"><mc:AlternateContent>` +
            `<mc:Choice Requires="w14"><Relationship Id="rIdC" Type="${R}/styles" Target="styles.xml"/></mc:Choice>` +
            `<mc:Fallback><Relationship Id="rIdF" Type="${R}/styles"/></mc:Fallback></mc:AlternateContent></Relationships>`,
    };

    const refused: Record<string, string> = { ...readable, "c:dummy.bin": "a stored name the shared context refuses" };

    it.each([
        ["word-valid", "a refused", refused],
        ["strict", "a refused", refused],
        ["word-valid", "a readable", readable],
        ["strict", "a readable", readable],
    ] as const)("%s: in %s directory, a direct check after a repair reads the repaired branch", async (profile, _kind, parts) => {
        await withTempDir(async (dir) => {
            for (const [part, xml] of Object.entries(parts)) {
                await mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await writeFile(path.join(dir, part), xml);
            }
            const validator = new DOCXSchemaValidator({ unpackedDir: dir, profile });
            const before = await validator.validateRelationshipElements();
            expect(before.issues.map((issue) => `${issue.code}@${issue.path}`)).toEqual([
                "rels-empty-element@word/_rels/document.xml.rels",
            ]);
            expect(await validator.repairIgnorable()).toBe(1);
            const after = await validator.validateRelationshipElements();
            expect(after).toEqual(await new DOCXSchemaValidator({ unpackedDir: dir, profile }).validateRelationshipElements());
            expect(after.valid).toBe(true);
            expect(after.issues).toEqual([]);
        });
    });
});

/**
 * #657: the XSD walk's main document comes from `_rels/.rels` read through
 * Word's markup compatibility view, so a repair can change it. Its
 * `officeDocument` relationship is in an `mc:Choice Requires="w14"` (to
 * `word/doc2.xml`) with a Fallback to `word/document.xml`, whose root is
 * ignorable, and `w14` is undeclared until `repairIgnorable()` declares it.
 * Main, that root is `mce-main-root-ignored`; not main, `mce-root-ignored`.
 */
describe("the main document after a repair (#657)", { timeout: 120_000 }, () => {
    const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const parts: Record<string, string> = {
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/doc2.xml" ContentType="${MAIN}"/></Types>`,
        ),
        "_rels/.rels":
            `<Relationships xmlns="${PR}" xmlns:mc="${MC}" mc:Ignorable="w14"><mc:AlternateContent>` +
            `<mc:Choice Requires="w14"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/doc2.xml"/></mc:Choice>` +
            `<mc:Fallback><Relationship Id="rId2" Type="${R}/officeDocument" Target="word/document.xml"/></mc:Fallback>` +
            `</mc:AlternateContent></Relationships>`,
        "word/document.xml": `<x:document xmlns:x="urn:x" xmlns:mc="${MC}" mc:Ignorable="x"/>`,
        "word/doc2.xml": MINIMAL_PARTS["word/document.xml"]!,
    };
    const rootFindings = (result: { issues: readonly ValidationIssue[] }): string[] =>
        result.issues
            .filter((issue) => issue.code === "mce-main-root-ignored" || issue.code === "mce-root-ignored")
            .map((issue) => `${issue.severity}:${issue.code}@${issue.path}`);

    it.each(["validateAgainstXsd", "validate"] as const)("strict: %s after a repair reads the repaired main document", async (method) => {
        await withTempDir(async (dir) => {
            for (const [part, xml] of Object.entries(parts)) {
                await mkdir(path.dirname(path.join(dir, part)), { recursive: true });
                await writeFile(path.join(dir, part), xml);
            }
            const validator = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
            expect(rootFindings(await validator.validateAgainstXsd())).toEqual(["error:mce-main-root-ignored@word/document.xml"]);
            expect(await validator.repairIgnorable()).toBe(1);
            const after = rootFindings(await validator[method]());
            expect(after).toEqual(rootFindings(await new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" })[method]()));
            expect(after).toEqual(["info:mce-root-ignored@word/document.xml"]);
        });
    });
});
