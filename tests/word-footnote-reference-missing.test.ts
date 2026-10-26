/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";
import { writeFile } from "node:fs/promises";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";
import { withTempDir } from "../src/lib/run-cli";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-footnote-reference-missing.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-footnote-reference-present.docx");

describe("dangling endnote references (#784, session 58 s58a)", { timeout: 30000 }, () => {
    it.each(["5", "1"])("judges endnote 1 against a part holding endnote %s", async (held) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "endnotes.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                        "</w:r></w:p>",
                        '</w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>',
                    ),
                    "word/endnotes.xml": `<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="${held}"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:endnote></w:endnotes>`,
                    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                        "</Relationships>",
                        '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>',
                    ),
                    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                        "</Types>",
                        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>',
                    ),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(held === "1");
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual(
                held === "1"
                    ? []
                    : [
                          expect.objectContaining({
                              code: "word-endnote-reference-missing",
                              path: "word/document.xml",
                              message: expect.stringContaining("endnote reference uses id 1"),
                          }),
                      ],
            );
            const selected = await validate(file, { profile: "word-valid", checks: ["word-endnote-reference-missing"] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });
});

describe("notes parts stored under another name (#793, session 60 s60a)", { timeout: 30000 }, () => {
    /** The s60a builder's `notesAt`: a used notes part holding `held`, stored as `word/<name>` and related as rIdN; the body references note 1. */
    const notesAt = (kind: "footnote" | "endnote", name: string, held: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
            ),
            [`word/${name}`]:
                `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${name}"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    it.each([
        ["footnote", "notes-foot.xml", "word-footnote-reference-missing"],
        ["endnote", "notes-end.xml", "word-endnote-reference-missing"],
    ] as const)("accepts the %ss part its relationship names, stored as word/%s", async (kind, name, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "renamed.docx");
            await writeFile(file, await notesAt(kind, name, "1"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it.each([
        ["footnote", "footnotes.xml"],
        ["endnote", "endnotes.xml"],
    ] as const)("accepts the standard-name calibration (%s)", async (kind, name) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "standard-name.docx");
            await writeFile(file, await notesAt(kind, name, "1"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it.each([
        ["footnote", "notes-foot.xml", "word-footnote-reference-missing", "footnote reference uses id 1"],
        ["endnote", "notes-end.xml", "word-endnote-reference-missing", "endnote reference uses id 1"],
    ] as const)("still rejects the %ss part stored as word/%s when it holds another id", async (kind, name, code, message) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "renamed-other-id.docx");
            await writeFile(file, await notesAt(kind, name, "5"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({ code, path: "word/document.xml", message: expect.stringContaining(message) }),
            ]);
        });
    });
});

describe("standard-name notes parts no relationship names (#797, session 61 s61a/s61b)", { timeout: 30000 }, () => {
    /** The s61 builder's `notesXml`: a notes part holding the note `held`. */
    const notesXml = (kind: "footnote" | "endnote", held: string): string =>
        `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`;

    /** The s61 builder's `mainWith`: `parts` with a `w:<kind>Reference w:id="1"` run appended in the main document's only paragraph. */
    const mainWith = (kind: "footnote" | "endnote", parts: Record<string, string>): Promise<Buffer> =>
        buildDocx({
            ...parts,
            "word/document.xml": parts["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
            ),
        });

    /** The s61a builder's `unrelatedStandard`: `word/<name>` holding note 1, typed by its Override but named by no relationship; the body references note 1. */
    const unrelatedStandard = (kind: "footnote" | "endnote", name: string): Promise<Buffer> =>
        mainWith(kind, {
            ...MINIMAL_PARTS,
            [`word/${name}`]: notesXml(kind, "1"),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    /** The s61a builder's `relatedOtherIdPlusUnrelatedStandard`: the `<kind>s` relationship names `word/<renamed>` holding note 5, and an unrelated `word/<kind>s.xml` holds note 1; the body references note 1. */
    const relatedOtherIdPlusUnrelatedStandard = (kind: "footnote" | "endnote", renamed: string): Promise<Buffer> =>
        mainWith(kind, {
            ...MINIMAL_PARTS,
            [`word/${renamed}`]: notesXml(kind, "5"),
            [`word/${kind}s.xml`]: notesXml(kind, "1"),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${renamed}"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${renamed}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/>` +
                    `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    /** The s61b builder's `unrelatedStandardNoReference`: the unrelated `word/<kind>s.xml` of s61a, with nothing in the body referencing it. */
    const unrelatedStandardNoReference = (kind: "footnote" | "endnote"): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            [`word/${kind}s.xml`]: notesXml(kind, "1"),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    it.each([
        ["footnote", "footnotes.xml", "word-footnote-reference-missing", "footnote reference uses id 1"],
        ["endnote", "endnotes.xml", "word-endnote-reference-missing", "endnote reference uses id 1"],
    ] as const)("rejects the %s reference the unrelated word/%s alone holds", async (kind, name, code, message) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unrelated-standard.docx");
            await writeFile(file, await unrelatedStandard(kind, name));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([
                expect.objectContaining({ code, path: "word/document.xml", message: expect.stringContaining(message) }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });

    it.each([
        ["footnote", "notes-foot.xml", "word-footnote-reference-missing", "footnote reference uses id 1"],
        ["endnote", "notes-end.xml", "word-endnote-reference-missing", "endnote reference uses id 1"],
    ] as const)("rejects the %s reference the unrelated part holds beside word/%s holding note 5", async (kind, renamed, code, message) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "related-other-id-plus-unrelated-standard.docx");
            await writeFile(file, await relatedOtherIdPlusUnrelatedStandard(kind, renamed));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([
                expect.objectContaining({ code, path: "word/document.xml", message: expect.stringContaining(message) }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });

    it.each([["footnote"], ["endnote"]] as const)(
        "accepts the unrelated word/%ss.xml nothing references (s61b calibration)",
        async (kind) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "unrelated-standard-no-reference.docx");
                await writeFile(file, await unrelatedStandardNoReference(kind));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            });
        },
    );
});

describe("a main document stored under another name (#800, session 62 s62a)", { timeout: 30000 }, () => {
    /** The #800 builder's `notesXml`: a notes part holding the note `held`, stored under its standard name. */
    const notesXml = (kind: "footnote" | "endnote", held: string): string =>
        `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`;

    /**
     * The #800 builder's `referenceIn`: `parts` with a `w:<kind>Reference
     * w:id="1"` run appended in the main document's only paragraph, at
     * `main` — `word/document.xml` for the standard-name row, `word/main.xml`
     * for the renamed one.
     */
    const referenceIn = (kind: "footnote" | "endnote", main: string, parts: Record<string, string>): Promise<Buffer> =>
        buildDocx({
            ...parts,
            [main]: parts[main]!.replace("</w:r></w:p>", `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`),
        });

    /**
     * The #800 builder's `renamedMain`: the main document stored as
     * `word/main.xml` — the part, its sidecar, the `_rels/.rels`
     * officeDocument Target and the content-type Override PartName all
     * renamed, content identical — whose sidecar relates `word/<kind>s.xml`
     * holding `held`; the body references note 1.
     */
    const renamedMain = (kind: "footnote" | "endnote", held: string): Promise<Buffer> => {
        const { "word/document.xml": main, "word/_rels/document.xml.rels": sidecar, ...rest } = MINIMAL_PARTS;
        return referenceIn(kind, "word/main.xml", {
            ...rest,
            "word/main.xml": main,
            "word/_rels/main.xml.rels": sidecar.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
            ),
            [`word/${kind}s.xml`]: notesXml(kind, held),
            "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace('Target="word/document.xml"', 'Target="word/main.xml"'),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                'PartName="/word/document.xml"',
                'PartName="/word/main.xml"',
            ).replace(
                "</Types>",
                `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });
    };

    /**
     * The #800 builder's `standardMain`: the minimal package unchanged in
     * its names, whose sidecar relates `word/endnotes.xml` holding `held`;
     * the body references endnote 1.
     */
    const standardMain = (held: string): Promise<Buffer> =>
        referenceIn("endnote", "word/document.xml", {
            ...MINIMAL_PARTS,
            "word/endnotes.xml": notesXml("endnote", held),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>',
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>',
            ),
        });

    it.each([
        ["footnote", "word-footnote-reference-missing", "footnote reference uses id 1"],
        ["endnote", "word-endnote-reference-missing", "endnote reference uses id 1"],
    ] as const)("rejects the %s reference in the main document stored as word/main.xml", async (kind, code, message) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "renamed-main.docx");
            await writeFile(file, await renamedMain(kind, "5"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([expect.objectContaining({ code, path: "word/main.xml", message: expect.stringContaining(message) })]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });

    it("accepts the renamed main whose related part holds the referenced endnote", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "renamed-main-calibration.docx");
            await writeFile(file, await renamedMain("endnote", "1"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it("still rejects the standard-name row at word/document.xml", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "standard-main.docx");
            await writeFile(file, await standardMain("5"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: "word-endnote-reference-missing",
                    path: "word/document.xml",
                    message: expect.stringContaining("endnote reference uses id 1"),
                }),
            ]);
        });
    });
});

describe("a story-named part no relationship reaches (#800, session 62 s62d)", { timeout: 30000 }, () => {
    /**
     * The session 62 specimen `unrelated-header-dangling-footnote-reference`,
     * which Word opened clean: `word/header9.xml`, typed header by its
     * Override and named by no relationship, holds a reference to footnote 7;
     * the package has no footnotes part.
     */
    const unrelatedHeader = (): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/header9.xml":
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
                '<w:p><w:r><w:t>Header</w:t></w:r><w:r><w:footnoteReference w:id="7"/></w:r></w:p></w:hdr>',
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/header9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
            ),
        });

    it("reads no references from an unrelated word/header9.xml", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unrelated-header.docx");
            await writeFile(file, await unrelatedHeader());
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: ["word-footnote-reference-missing"] });
            expect(selected.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === "word-footnote-reference-missing")).toEqual([]);
        });
    });
});

describe("word-valid dangling footnote reference", { timeout: 30000 }, () => {
    it("rejects a footnote reference whose id is not a footnote id", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-footnote-reference-missing")).toBe(true);
    });

    it("does not reject the document once that reference is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-footnote-reference-missing")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-footnote-reference-missing")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-footnote-reference-missing")).toBe(true);
    });
});
