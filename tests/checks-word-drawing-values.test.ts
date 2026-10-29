/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import {
    wordDrawingScalarWhitespaceCheck,
    wordDrawingArcRadiusCheck,
    wordDrawingExtentNegativeCheck,
} from "../src/scripts/office/checks/word-drawing-values";

describe("word-drawing-scalar-whitespace", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<wp:posOffset> 1 </wp:posOffset>") });
        expect(await wordDrawingScalarWhitespaceCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-drawing-scalar-whitespace",
                path: "word/document.xml",
                message:
                    "<wp:posOffset> contains leading/trailing whitespace around scalar value ' 1 '; Microsoft Word refuses this shape in drawing anchors.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                "<wp:posOffset>1</wp:posOffset><wp:align>left</wp:align><wp14:pctWidth>2</wp14:pctWidth><wp14:pctHeight>3</wp14:pctHeight>",
            ),
        });
        expect(await wordDrawingScalarWhitespaceCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDrawingScalarWhitespaceCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDrawingScalarWhitespaceCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-drawing-arc-radius", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<a:arcTo wR="1.5" hR="2"/>') });
        expect(await wordDrawingArcRadiusCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-drawing-arc-radius",
                path: "word/document.xml",
                message: '<a:arcTo> wR="1.5" is not an integer. Microsoft Word cannot open the file. Use an integer radius.',
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<a:arcTo wR="1" hR="0"/><a:arcTo/>') });
        expect(await wordDrawingArcRadiusCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDrawingArcRadiusCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDrawingArcRadiusCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-drawing-extent-negative", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<wp:extent cx="-1" cy="2"/>') });
        expect(await wordDrawingExtentNegativeCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-drawing-extent-negative",
                path: "word/document.xml",
                message: '<wp:extent> cx="-1" is negative. Microsoft Word cannot open the file. Use a positive extent.',
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<wp:extent cx="0" cy="2"/><wp:extent/>') });
        expect(await wordDrawingExtentNegativeCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDrawingExtentNegativeCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDrawingExtentNegativeCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("checks both radii/extents, ignores scalar child elements, and preserves repr escapes/truncation", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/header1.xml": document(`<a:arcTo wR="-1" hR="x"/><wp:extent cx="1" cy="-2"/><wp:align><w:t>x</w:t></wp:align>`),
    });
    expect((await wordDrawingArcRadiusCheck.run(ctx)).map((i) => i.message)).toEqual([
        '<a:arcTo> wR="-1" is not an integer. Microsoft Word cannot open the file. Use an integer radius.',
        '<a:arcTo> hR="x" is not an integer. Microsoft Word cannot open the file. Use an integer radius.',
    ]);
    expect(await wordDrawingExtentNegativeCheck.run(ctx)).toHaveLength(1);
    expect(await wordDrawingScalarWhitespaceCheck.run(ctx)).toEqual([]);
    const text = " " + "a".repeat(90) + " ";
    expect(
        (
            await wordDrawingScalarWhitespaceCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": document(`<wp:align>${text}</wp:align>`) }),
            )
        )[0]?.message,
    ).toContain("' " + "a".repeat(78) + "...");
    const escaped = ValidationContext.fromEntries({ "word/document.xml": document("<wp:align> \\ ' \n\r\t</wp:align>") });
    expect(await wordDrawingScalarWhitespaceCheck.run(escaped)).toHaveLength(1);
});
