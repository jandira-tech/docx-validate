/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { wordDrawingExtentMissingCheck as check } from "../src/scripts/office/checks/word-drawing";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const drawing = (children: string) => `<drawing xmlns:d="${WP}" xmlns:a="${A}">${children}</drawing>`;

describe("missing drawing extent", () => {
    it("rejects drawing containers without a direct extent, as in header-with-rid-no-sidecar.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/header1.xml": drawing("<d:inline><a:graphic/></d:inline>"),
            "word/header2.xml": drawing("<d:anchor><a:graphic/></d:anchor>"),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/header1.xml",
                message: expect.stringMatching(/extent.*inline.*before.*graphic/),
            }),
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/header2.xml",
                message: expect.stringMatching(/extent.*anchor.*before.*graphic/),
            }),
        ]);
    });

    it("accepts the complete header picture scaffold control and leaves the XML untouched", async () => {
        const xml = drawing('<d:inline><d:extent cx="1" cy="1"/><d:docPr id="1"/><a:graphic/></d:inline>');
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": xml });
        expect(await check.run(ctx)).toEqual([]);
        expect(ctx.text("word/header1.xml")).toBe(xml);
    });

    it("leaves other namespaces, unrelated elements, and non-story parts alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": drawing("<a:inline/><a:anchor/><d:docPr/>"),
            "customXml/item1.xml": drawing("<d:inline/>"),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("skips unparseable story parts", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": drawing("<d:inline>") });
        expect(ctx.xml("word/document.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
