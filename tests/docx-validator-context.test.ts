/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { MINIMAL_PARTS } from "./support/minimal-docx";

function writeParts(dir: string, parts: Readonly<Record<string, string>>): void {
    for (const [name, content] of Object.entries(parts)) {
        mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
        writeFileSync(path.join(dir, name), content);
    }
}

afterEach(() => {
    vi.restoreAllMocks();
});

describe("DOCXSchemaValidator over the shared validation context", { timeout: 60000 }, () => {
    it("parses word/document.xml once for every Word check of a run", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, MINIMAL_PARTS);
            const xml = vi.spyOn(ValidationContext.prototype, "xml");
            const result = await new DOCXSchemaValidator({ unpackedDir: dir, profile: "word-valid" }).validate();
            expect(result.valid).toBe(true);
            const contexts = new Set(xml.mock.contexts as ValidationContext[]);
            expect(contexts.size).toBe(1);
            const [ctx] = [...contexts];
            const documentReads = xml.mock.calls.filter(([part]) => String(part).endsWith("word/document.xml")).length;
            expect(documentReads).toBeGreaterThan(1);
            expect(ctx?.parseCount("word/document.xml")).toBe(1);
        });
    });

    it("builds a fresh context for each run, so a rewritten part is seen by the next run", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, MINIMAL_PARTS);
            const validator = new DOCXSchemaValidator({ unpackedDir: dir, profile: "word-valid" });
            expect((await validator.validate()).valid).toBe(true);
            const broken = MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:sectPr>",
                "<w:tbl><w:tblPr/><w:tblGrid/><w:tr/></w:tbl><w:p/><w:sectPr>",
            );
            writeFileSync(path.join(dir, "word", "document.xml"), broken);
            const second = await validator.validate();
            expect(second.issues.some((i) => i.severity === "error" && i.code === "word-row-without-cell")).toBe(true);
        });
    });

    // The comment rules read the shared context (#384); called directly,
    // outside validate(), each call must still read the package as it is now.
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const COMMENTS = `<w:comments ${W}><w:comment w:id="0"><w:p/></w:comment></w:comments>`;
    const withMarkers = (markers: string): string =>
        MINIMAL_PARTS["word/document.xml"]!.replace("<w:r><w:t>Hello</w:t></w:r>", `${markers}<w:r><w:t>Hello</w:t></w:r>`);
    const PAIRED = '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>';
    // Related, so the part defines its comment (#896).
    const RELATED = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
    );

    it("reads the package afresh for each direct comment-marker call", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, {
                ...MINIMAL_PARTS,
                "word/_rels/document.xml.rels": RELATED,
                "word/comments.xml": COMMENTS,
                "word/document.xml": withMarkers('<w:commentRangeEnd w:id="0"/>'),
            });
            const validator = new DOCXSchemaValidator({ unpackedDir: dir });
            expect((await validator.validateCommentMarkers()).issues.map((i) => i.code)).toContain("comment-orphan-end");
            writeFileSync(path.join(dir, "word", "document.xml"), withMarkers(PAIRED));
            expect((await validator.validateCommentMarkers()).issues).toEqual([]);
        });
    });

    it("reads the package afresh for each direct comment-threading call", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, { ...MINIMAL_PARTS, "word/_rels/document.xml.rels": RELATED, "word/comments.xml": COMMENTS });
            const validator = new DOCXSchemaValidator({ unpackedDir: dir });
            expect((await validator.validateCommentThreading()).issues.map((i) => i.code)).toContain("comment-thread-count-mismatch");
            writeFileSync(path.join(dir, "word", "document.xml"), withMarkers(PAIRED));
            expect((await validator.validateCommentThreading()).issues).toEqual([]);
        });
    });
});
