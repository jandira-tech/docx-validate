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

describe("notes defined only through a reached glossary (#828, session 63 s63c)", { timeout: 30000 }, () => {
    /** The s63c builder's `glossaryWithFootnotes`; the endnotes twin is unprobed. */
    const glossaryWithNotes = (kind: "footnote" | "endnote"): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/glossary/document.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>' +
            '<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>' +
            "<w:docPartBody><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>",
        "word/glossary/_rels/document.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
        [`word/glossary/${kind}s.xml`]:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:${kind} w:id="1"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            '<Relationship Id="rIdG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>',
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>' +
                `<Override PartName="/word/glossary/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
        ),
    });

    it.each([
        ["footnote", "full"],
        ["footnote", "selected"],
        ["endnote", "full"],
        ["endnote", "selected"],
    ] as const)("rejects the body %s reference defined only by the glossary in a %s run", async (kind, mode) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const parts = glossaryWithNotes(kind);
            const file = path.join(dir, "glossary-only-notes.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...parts,
                    "word/document.xml": parts["word/document.xml"]!.replace(
                        "</w:r></w:p>",
                        `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
                    ),
                }),
            );
            const result = await validate(file, { profile: "word-valid", ...(mode === "selected" ? { checks: [code] } : {}) });
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([
                expect.objectContaining({
                    code,
                    path: "word/document.xml",
                    message: expect.stringContaining(`${kind} reference uses id 1`),
                }),
            ]);
            expect(errors[0]?.message).toContain(
                `the notes parts the main document relates define no w:${kind} with that id; only a reached glossary's notes part does.`,
            );
            expect(errors[0]?.message).toContain("Microsoft Word offers to recover the file.");
            expect(errors[0]?.message).toContain(
                `Remove the w:${kind}Reference, or add the note to a notes part the main document relates.`,
            );
            expect(result.valid).toBe(false);
        });
    });

    it.each([
        ["footnote", "glossary"],
        ["endnote", "glossary"],
        ["footnote", "main"],
        ["endnote", "main"],
    ] as const)("keeps the glossary's own %s reference satisfied by the %s notes (unprobed)", async (kind, owner) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const parts = glossaryWithNotes(kind);
            if (owner === "main") {
                parts[`word/${kind}s.xml`] = parts[`word/glossary/${kind}s.xml`]!;
                delete parts[`word/glossary/${kind}s.xml`];
                delete parts["word/glossary/_rels/document.xml.rels"];
                parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
                );
                parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(`/word/glossary/${kind}s.xml`, `/word/${kind}s.xml`);
            }
            const file = path.join(dir, "glossary-reference.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...parts,
                    "word/glossary/document.xml": parts["word/glossary/document.xml"]!.replace(
                        "</w:r></w:p>",
                        `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
                    ),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });
});

describe("notes parts typed application/octet-stream (#829, sessions 63 s63c and 64 s64b)", { timeout: 30000 }, () => {
    /**
     * The s64 builder's `footnotesAt`: note 1 stored at `stored`, typed
     * `type` by its Override and named only by the main document's
     * `<kind>s` relationship; the body references note 1.
     */
    const notesAt = (kind: "footnote" | "endnote", stored: string, type: string, target: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
            ),
            [stored]:
                `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${target}"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/${stored}" ContentType="${type}"/></Types>`,
            ),
        });

    it.each([
        ["footnote", "word/notes.xml", "notes.xml"],
        ["footnote", "word/footnotes.xml", "footnotes.xml"],
        ["footnote", "word/notes.bin", "notes.bin"],
        ["endnote", "word/endnotes.xml", "endnotes.xml"],
    ] as const)("rejects the body %s reference whose note sits in %s typed application/octet-stream", async (kind, stored, target) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "octet-stream-notes.docx");
            await writeFile(file, await notesAt(kind, stored, "application/octet-stream", target));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([
                expect.objectContaining({
                    code,
                    path: "word/document.xml",
                    message: expect.stringContaining(`${kind} reference uses id 1`),
                }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(false);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });

    it.each([
        ["footnote", "word/footnotes.xml", "footnotes.xml"],
        ["endnote", "word/endnotes.xml", "endnotes.xml"],
    ] as const)("accepts the %s calibration: word/%s typed its own type still defines the note", async (kind, stored, target) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "octet-stream-calibration.docx");
            await writeFile(
                file,
                await notesAt(kind, stored, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml`, target),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });
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

describe("note references in a read header (#806, session 62 s62d/s62e)", { timeout: 30000 }, () => {
    /** The s62 builder's `withNotes`: the related `word/<kind>s.xml` holds note `held`. */
    const withNotes = (kind: "footnote" | "endnote", held: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        [`word/${kind}s.xml`]:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
        ),
    });

    /** The s62 builder's `withReadHeader`: a section-referenced header holding note reference 1. */
    const withReadHeader = (kind: "footnote" | "endnote", held: string): Record<string, string> => {
        const parts = withNotes(kind, held);
        return {
            ...parts,
            "word/document.xml": parts["word/document.xml"]!.replace(
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:u="urn:example:unknown"',
            ).replace("<w:sectPr>", `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`),
            "word/header1.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:p><w:r><w:t>Header</w:t></w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p></w:hdr>`,
            "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": parts["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
            ),
        };
    };

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read header's dangling %s reference in a full run", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-header.docx");
            await writeFile(file, await buildDocx(withReadHeader(kind, "5")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read header's dangling %s reference in a selected run", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-header-selected.docx");
            await writeFile(file, await buildDocx(withReadHeader(kind, "5")));
            const result = await validate(file, { profile: "word-valid", checks: [code] });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("reports no missing %s reference from the read header under strict", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-header-strict.docx");
            await writeFile(file, await buildDocx(withReadHeader(kind, "5")));
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read header's matching-%s calibration", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-header-calibration.docx");
            await writeFile(file, await buildDocx(withReadHeader(kind, "1")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it("still rejects the dangling footnote reference in the body", async () => {
        await withTempDir(async (dir) => {
            const parts = withNotes("footnote", "5");
            const file = path.join(dir, "body.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...parts,
                    "word/document.xml": parts["word/document.xml"]!.replace(
                        "</w:r></w:p>",
                        '</w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
                    ),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({ code: "word-footnote-reference-missing", path: "word/document.xml" }),
            ]);
        });
    });
});

describe("note references in a read footer (#826, session 63 s63d)", { timeout: 30000 }, () => {
    /** The s63 builder's `withNotes`: the related `word/<kind>s.xml` holds note `held`. */
    const withNotes = (kind: "footnote" | "endnote", held: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        [`word/${kind}s.xml`]:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
        ),
    });

    /** The s63 builder's `withReadFooter`: a section-referenced `word/footer1.xml` holding note reference 1. */
    const withReadFooter = (kind: "footnote" | "endnote", held: string): Record<string, string> => {
        const parts = withNotes(kind, held);
        return {
            ...parts,
            "word/document.xml": parts["word/document.xml"]!.replace(
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
            ).replace("<w:sectPr>", `<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/>`),
            "word/footer1.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:p><w:r><w:t>Footer</w:t></w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p></w:ftr>`,
            "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>',
            ),
            "[Content_Types].xml": parts["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
            ),
        };
    };

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read footer's dangling %s reference in a full run", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-footer.docx");
            await writeFile(file, await buildDocx(withReadFooter(kind, "5")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read footer's dangling %s reference in a selected run", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-footer-selected.docx");
            await writeFile(file, await buildDocx(withReadFooter(kind, "5")));
            const result = await validate(file, { profile: "word-valid", checks: [code] });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("reports no missing %s reference from the read footer under strict", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-footer-strict.docx");
            await writeFile(file, await buildDocx(withReadFooter(kind, "5")));
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each([
        ["endnote", "word-endnote-reference-missing"],
        ["footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the read footer's matching-%s calibration", async (kind, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "read-footer-calibration.docx");
            await writeFile(file, await buildDocx(withReadFooter(kind, "1")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });
});

describe("note ids spelled 01 or +1 against a note spelled 1 (#801, session 62 s62b)", { timeout: 30000 }, () => {
    /**
     * The s62b builder's `referenced`: `word/<kind>s.xml` holding the note
     * `held`, related by the main document and typed by its Override, and a
     * `w:<kind>Reference w:id="<spelled>"` run appended in the main document's
     * only paragraph.
     */
    const referenced = (kind: "footnote" | "endnote", held: string, spelled: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="${spelled}"/></w:r></w:p>`,
            ),
            [`word/${kind}s.xml`]:
                `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:${kind} w:id="${held}"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    it.each([
        ["footnote", "01", "word-footnote-reference-missing"],
        ["endnote", "01", "word-endnote-reference-missing"],
        ["footnote", "+1", "word-footnote-reference-missing"],
        ["endnote", "+1", "word-endnote-reference-missing"],
    ] as const)("accepts the %s reference spelled %s against note 1", async (kind, spelled, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "spelled.docx");
            await writeFile(file, await referenced(kind, "1", spelled));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it.each([
        ["footnote", "01", "word-footnote-reference-missing"],
        ["endnote", "+1", "word-endnote-reference-missing"],
    ] as const)("accepts the %s calibration spelling note and reference %s alike", async (kind, spelled, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "spelled-calibration.docx");
            await writeFile(file, await referenced(kind, spelled, spelled));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    it.each([
        ["footnote", "word-footnote-reference-missing", "footnote reference uses id 01"],
        ["endnote", "word-endnote-reference-missing", "endnote reference uses id 01"],
    ] as const)("still rejects the %s reference spelled 01 against note 5", async (kind, code, message) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "spelled-other-id.docx");
            await writeFile(file, await referenced(kind, "5", "01"));
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
});

describe("a package view the read model cannot use (#816)", { timeout: 30000 }, () => {
    /**
     * The #816 specimen: the minimal package whose `word/document.xml`
     * holds a `w:<kind>Reference w:id="1"` run, with no notes part, so the
     * reference dangles wherever it is read from.
     */
    const danglingInBody = (kind: "footnote" | "endnote"): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            "</w:r></w:p>",
            `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
        ),
    });

    /**
     * The four package-relationship defects of the issue's rows, each
     * applied to `parts`: the package relationships part missing, truncated
     * past its start tag so it does not parse, holding no `officeDocument`
     * relationship, and naming a main document the package does not store.
     * Each defect leaves `word/document.xml` well formed, and its own
     * finding — `word-main-document-part`, or `xml-syntax` on `_rels/.rels`
     * for the truncated row — still blocks the full run, while the read
     * model and the typed walks return nothing for the notes.
     */
    const packageRelsDefects: Readonly<Record<string, (parts: Record<string, string>) => Record<string, string>>> = {
        missing: (parts) => {
            const next = { ...parts };
            delete next["_rels/.rels"];
            return next;
        },
        truncated: (parts) => ({ ...parts, "_rels/.rels": '<?xml version="1.0"?><Relationships' }),
        "holding no officeDocument relationship": (parts) => ({
            ...parts,
            "_rels/.rels":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        }),
        "naming a missing main document": (parts) => ({
            ...parts,
            "_rels/.rels": parts["_rels/.rels"]!.replace('Target="word/document.xml"', 'Target="word/missing.xml"'),
        }),
    };

    it.each([
        ["footnote", "missing"],
        ["footnote", "truncated"],
        ["footnote", "holding no officeDocument relationship"],
        ["footnote", "naming a missing main document"],
        ["endnote", "missing"],
        ["endnote", "truncated"],
        ["endnote", "holding no officeDocument relationship"],
        ["endnote", "naming a missing main document"],
    ] as const)("reports the dangling %s reference with _rels/.rels %s", async (kind, defect) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const message = `${kind} reference uses id 1`;
            const file = path.join(dir, "package-defect.docx");
            await writeFile(file, await buildDocx(packageRelsDefects[defect]!(danglingInBody(kind))));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(
                result.issues.some(
                    (issue) =>
                        issue.severity === "error" && issue.code === (defect === "truncated" ? "xml-syntax" : "word-main-document-part"),
                ),
            ).toBe(true);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([
                expect.objectContaining({ code, path: "word/document.xml", message: expect.stringContaining(message) }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(false);
            expect(selected.issues.filter((issue) => issue.code === code)).toEqual([
                expect.objectContaining({ code, path: "word/document.xml", message: expect.stringContaining(message) }),
            ]);
        });
    });

    it.each(["footnote", "endnote"] as const)("keeps a header's dangling %s reference unjudged, as #806 does", async (kind) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "package-defect-header.docx");
            const parts = packageRelsDefects.missing!({
                ...MINIMAL_PARTS,
                "word/header1.xml":
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                    `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:${kind}Reference w:id="1"/></w:r></w:p></w:hdr>`,
            });
            await writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-main-document-part")).toBe(true);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });

    it.each(["footnote", "endnote"] as const)("keeps a footer's dangling %s reference unjudged, as #826 does", async (kind) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "package-defect-footer.docx");
            const parts = packageRelsDefects.missing!({
                ...MINIMAL_PARTS,
                "word/footer1.xml":
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                    `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:${kind}Reference w:id="1"/></w:r></w:p></w:ftr>`,
            });
            await writeFile(file, await buildDocx(parts));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-main-document-part")).toBe(true);
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.issues.filter((issue) => issue.code === code)).toEqual([]);
        });
    });
});

describe("a notes relationship target escaping the package (#830, sessions 63 s63c and 64 s64a)", { timeout: 30000 }, () => {
    /**
     * The s64 builder's `notesAt`: note 1 stored at `stored`, typed by its
     * Override and named only by the `<kind>s` relationship whose target is
     * `target`; the body references note 1.
     */
    const notesAt = (kind: "footnote" | "endnote", stored: string, target: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
            ),
            [stored]:
                `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${target}"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/${stored}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    // Session 64 (s64a, two runs, all readings agree) opened each of these
    // clean: Word resolves the target as RFC 3986's remove_dot_segments
    // does, stopping the `..` segments at the package root, so the note at
    // `/footnotes.xml` (or `/endnotes.xml`) is defined. The two-level
    // footnotes row repeats session 63 (s63c), which word-valid rejected as
    // a false positive; the three-level row and the endnotes twin are
    // session 64's confirmations.
    it.each([
        ["footnote", "../../footnotes.xml"],
        ["footnote", "../../../footnotes.xml"],
        ["endnote", "../../endnotes.xml"],
    ] as const)("accepts the %s reference the escaping target %s resolves at the package root", async (kind, target) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "escaping.docx");
            await writeFile(file, await notesAt(kind, `${kind}s.xml`, target));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    // Unprobed: since #947 a package-absolute target's dot segments resolve
    // as a relative target's do, so its climb stops at the package root the
    // same way. Word has not been probed with a package-absolute target.
    it.each([
        ["footnote", "/../footnotes.xml"],
        ["endnote", "/../endnotes.xml"],
    ] as const)("accepts the %s reference the package-absolute target %s resolves at the package root (#947)", async (kind, target) => {
        await withTempDir(async (dir) => {
            const code = `word-${kind}-reference-missing`;
            const file = path.join(dir, "absolute-escaping.docx");
            await writeFile(file, await notesAt(kind, `${kind}s.xml`, target));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.filter((issue) => issue.code === code)).toEqual([]);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            expect(result.valid).toBe(true);
        });
    });

    it("still rejects the escaping relationship whose note sits only in word/footnotes.xml", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "escaping-in-word-folder.docx");
            await writeFile(file, await notesAt("footnote", "word/footnotes.xml", "../../footnotes.xml"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            const errors = result.issues.filter((issue) => issue.severity === "error");
            expect(errors).toEqual([
                expect.objectContaining({
                    code: "word-footnote-reference-missing",
                    path: "word/document.xml",
                    message: expect.stringContaining("footnote reference uses id 1"),
                }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: ["word-footnote-reference-missing"] });
            expect(selected.valid).toBe(result.valid);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(errors);
        });
    });

    it("accepts the in-package ../footnotes.xml calibration", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "one-level.docx");
            await writeFile(file, await notesAt("footnote", "footnotes.xml", "../footnotes.xml"));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });
});

describe("note references inside a notes or comments part (#852, sessions 65 s65a and 66 s66a)", { timeout: 30000 }, () => {
    /** The s65 builder's `noteRef`: a run holding a `w:<kind>Reference w:id="<id>"`. */
    const noteRef = (kind: "footnote" | "endnote", id: string): string => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;

    /**
     * The s65 builder's `notesWithNested`: the related in-package
     * `word/<kind>s.xml` holds note 1, whose paragraph also holds `nested`
     * — a same- or cross-kind note reference — and the body references
     * note 1.
     */
    const nestedNotes = (kind: "footnote" | "endnote", nested: string): Promise<Buffer> =>
        buildDocx({
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
                "</w:r></w:p>",
                `</w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`,
            ),
            [`word/${kind}s.xml`]:
                `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
                `<w:${kind} w:id="1"><w:p><w:r><w:t>Text</w:t></w:r>${nested}</w:p></w:${kind}></w:${kind}s>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`,
            ),
        });

    /**
     * The s66 builder's `glossaryWithFootnotes`: a reached glossary whose
     * own sidecar relates `word/glossary/footnotes.xml` holding footnote 1.
     */
    const glossaryWithFootnotes = (): Record<string, string> => ({
        ...MINIMAL_PARTS,
        "word/glossary/document.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>' +
            '<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>' +
            "<w:docPartBody><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>",
        "word/glossary/_rels/document.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
        "word/glossary/footnotes.xml":
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:footnote w:id="1"><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:footnote></w:footnotes>',
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            '<Relationship Id="rIdG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>',
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>' +
                '<Override PartName="/word/glossary/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
        ),
    });

    /**
     * The s66 builder's `commentsReach`: `word/comments.xml` holding
     * comment 0 with `inComment` in its paragraph, read by the main
     * document and/or the reached glossary, with (`notes`) or without the
     * glossary's footnotes part holding footnote 1.
     */
    const commentsReach = (main: boolean, glossary: boolean, inComment: string, notes: boolean): Promise<Buffer> => {
        const parts = glossaryWithFootnotes();
        const next: Record<string, string> = {
            ...parts,
            "word/comments.xml":
                '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
                `<w:comment w:id="0" w:author="a"><w:p><w:r><w:t>Text</w:t></w:r>${inComment}</w:p></w:comment></w:comments>`,
            "[Content_Types].xml": parts["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
            ),
        };
        if (!notes) {
            delete next["word/glossary/footnotes.xml"];
            next["word/glossary/_rels/document.xml.rels"] = next["word/glossary/_rels/document.xml.rels"]!.replace(
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
                "",
            );
            next["[Content_Types].xml"] = next["[Content_Types].xml"]!.replace(
                '<Override PartName="/word/glossary/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
                "",
            );
        }
        if (glossary) {
            next["word/glossary/document.xml"] = next["word/glossary/document.xml"]!.replace(
                "</w:r></w:p>",
                '</w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>',
            );
            next["word/glossary/_rels/document.xml.rels"] = next["word/glossary/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments.xml"/></Relationships>',
            );
        }
        if (!main) {
            return buildDocx(next);
        }
        return buildDocx({
            ...next,
            "word/document.xml": next["word/document.xml"]!.replace(
                "</w:r></w:p>",
                '</w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>',
            ),
            "word/_rels/document.xml.rels": next["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
            ),
        });
    };

    // Session 65 (s65a, two runs, all four readings agree) opened footnote 1
    // holding a footnote reference to 99 clean — as did its escaping twin —
    // and session 66 (s66a, same) endnote 1 holding a footnote reference to
    // 99, no footnotes part anywhere: with session 16 (s16a) having opened an
    // endnote reference inside a footnote clean, every note reference these
    // sessions stored inside a notes part opens clean, so none is judged.
    it.each([
        ["footnote", "footnote", "word-footnote-reference-missing"],
        ["endnote", "endnote", "word-endnote-reference-missing"],
        ["endnote", "footnote", "word-footnote-reference-missing"],
    ] as const)("accepts the %s 1 holding a dangling %s reference", async (kind, nested, code) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "nested.docx");
            await writeFile(file, await nestedNotes(kind, noteRef(nested, "99")));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            const selected = await validate(file, { profile: "word-valid", checks: [code] });
            expect(selected.valid).toBe(true);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        });
    });

    // Session 66 (s66a, two runs, all four readings agree) opened clean a
    // comment holding a footnote reference to 99, no part defining a
    // footnote, whether the main document or only the reached glossary
    // reads the comments part; session 65 (s65d, same) opened the main-read
    // twin clean whose footnote 1 only the glossary's notes part defines.
    // The shared-reach rows are refused for the reach, not the reference
    // (s66a), which is #850's rule, not this check's.
    it.each([
        [true, false, "99", false],
        [false, true, "99", false],
        [true, false, "1", true],
    ] as const)(
        "accepts the comment's footnote reference (read by main=%s, glossary=%s; id %s, glossary notes part %s)",
        async (main, glossary, id, notes) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "comments-reference.docx");
                await writeFile(file, await commentsReach(main, glossary, noteRef("footnote", id), notes));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
                const selected = await validate(file, { profile: "word-valid", checks: ["word-footnote-reference-missing"] });
                expect(selected.valid).toBe(true);
                expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual([]);
            });
        },
    );

    // The story roots stay judged: a reference in a reached glossary's own
    // document still blocks, an unprobed reading this layer keeps.
    it.each(["footnote", "endnote"] as const)("still rejects the dangling %s reference in the glossary's own document", async (kind) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "glossary-root.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...glossaryWithFootnotes(),
                    "word/glossary/document.xml": glossaryWithFootnotes()["word/glossary/document.xml"]!.replace(
                        "</w:r></w:p>",
                        `</w:r>${noteRef(kind, "5")}</w:p>`,
                    ),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([
                expect.objectContaining({
                    code: `word-${kind}-reference-missing`,
                    path: "word/glossary/document.xml",
                    message: expect.stringContaining(`${kind} reference uses id 5`),
                }),
            ]);
            const selected = await validate(file, { profile: "word-valid", checks: [`word-${kind}-reference-missing`] });
            expect(selected.valid).toBe(false);
            expect(selected.issues.filter((issue) => issue.severity === "error")).toEqual(
                result.issues.filter((issue) => issue.severity === "error"),
            );
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
