/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, WS } from "./support/word-rule-test-xml";
import {
    wordEmptyDrawingCheck,
    wordAnchorMissingAttributeCheck,
    wordPositionOffsetAndAlignCheck,
} from "../src/scripts/office/checks/word-drawing-anchors";

describe("word-empty-drawing", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<w:p><w:r><w:drawing/></w:r></w:p>") });
        expect(await wordEmptyDrawingCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-empty-drawing",
                path: "word/document.xml",
                message:
                    "An empty <w:drawing> has no child element. Microsoft Word does not open the file cleanly. Remove the drawing, or give it a wp:anchor or wp:inline child.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:p><w:r><w:drawing><wp:inline/></w:drawing></w:r><w:drawing/></w:p>"),
        });
        expect(await wordEmptyDrawingCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordEmptyDrawingCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordEmptyDrawingCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-anchor-missing-attribute", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<wp:anchor locked="0" layoutInCell="1"/>') });
        expect(await wordAnchorMissingAttributeCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-anchor-missing-attribute",
                path: "word/document.xml",
                message: "A wp:anchor is missing allowOverlap. Microsoft Word cannot open the file. Add that attribute.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document('<wp:anchor locked="0" layoutInCell="1" allowOverlap="1"/>'),
        });
        expect(await wordAnchorMissingAttributeCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordAnchorMissingAttributeCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordAnchorMissingAttributeCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-position-offset-and-align", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<wp:positionH>text<wp:posOffset>1</wp:posOffset><wp:align>left</wp:align></wp:positionH>"),
        });
        expect(await wordPositionOffsetAndAlignCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-position-offset-and-align",
                path: "word/document.xml",
                message:
                    "A wp:positionH contains both wp:posOffset and wp:align. Microsoft Word cannot open the file. Keep only one of those children.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                "<wp:positionH><wp:posOffset>1</wp:posOffset><w:align/></wp:positionH><wp:positionV><wp:align>top</wp:align></wp:positionV><wp:positionH/>",
            ),
        });
        expect(await wordPositionOffsetAndAlignCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordPositionOffsetAndAlignCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordPositionOffsetAndAlignCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("checks each missing anchor attribute and the vertical-position conflict", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(
            `<wp:anchor/><wp:positionV><wp:align>top</wp:align><wp:posOffset>1</wp:posOffset></wp:positionV><w:drawing>text</w:drawing>`,
        ),
    });
    expect((await wordAnchorMissingAttributeCheck.run(ctx)).map((i) => i.message)).toEqual(
        ["locked", "layoutInCell", "allowOverlap"].map(
            (a) => `A wp:anchor is missing ${a}. Microsoft Word cannot open the file. Add that attribute.`,
        ),
    );
    expect((await wordPositionOffsetAndAlignCheck.run(ctx))[0]?.message).toContain("wp:positionV contains both");
    expect(await wordEmptyDrawingCheck.run(ctx)).toHaveLength(1);
    expect(
        await wordEmptyDrawingCheck.run(ValidationContext.fromEntries({ "word/document.xml": document(`<w:r><w:drawing/></w:r>`, WS) })),
    ).toHaveLength(1);
});
