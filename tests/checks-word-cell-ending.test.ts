/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordCellEndsWithTableCheck } from "../src/scripts/office/checks/word-cell-ending";

describe("word-cell-ends-with-table", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<w:tc>text<w:p/><w:tbl/>tail</w:tc>") });
        expect(await wordCellEndsWithTableCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-cell-ends-with-table",
                path: "word/document.xml",
                message:
                    "A table cell ends with w:tbl and has no following w:p. Microsoft Word offers to recover the file. Add a w:p after that table.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:tc/><w:tc><w:p/></w:tc><w:tc><w:tbl/><w:p/></w:tc><w:tc><a:tbl/></w:tc>"),
        });
        expect(await wordCellEndsWithTableCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordCellEndsWithTableCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordCellEndsWithTableCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
