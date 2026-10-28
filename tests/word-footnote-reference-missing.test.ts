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
