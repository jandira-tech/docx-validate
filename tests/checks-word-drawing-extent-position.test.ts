/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { wordDrawingExtentMisplacedCheck as check } from "../src/scripts/office/checks/word-drawing-extent-position";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const drawing = (children: string) => `<drawing xmlns:d="${WP}" xmlns:a="${A}">${children}</drawing>`;

describe("misplaced drawing extent", () => {
    it("rejects extents after graphic, before anchor positioning, duplicated, or under the wrong parent", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/header1.xml": drawing("<d:inline><a:graphic/><d:extent/></d:inline>"),
            "word/header2.xml": drawing("<d:anchor><d:simplePos/><d:positionH/><d:positionV/><a:graphic/><d:extent/></d:anchor>"),
            "word/header3.xml": drawing("<d:anchor><d:simplePos/><d:positionH/><d:extent/><d:positionV/><a:graphic/></d:anchor>"),
            "word/header4.xml": drawing("<d:inline><d:extent/><d:extent/><a:graphic/></d:inline>"),
            "word/header5.xml": drawing("<a:graphic><d:extent/></a:graphic>"),
        });
        const issues = await check.run(ctx);
        expect(issues).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/header1.xml",
                message: expect.stringMatching(/extent.*inline.*before.*graphic/),
            }),
            expect.objectContaining({ code: check.id, path: "word/header2.xml", message: expect.stringContaining("anchor") }),
            expect.objectContaining({ code: check.id, path: "word/header3.xml" }),
            expect.objectContaining({ code: check.id, path: "word/header4.xml" }),
            expect.objectContaining({
                code: check.id,
                path: "word/header5.xml",
                message: expect.stringMatching(/extent.*a:graphic.*inline.*anchor/),
            }),
        ]);
    });

    it("accepts correctly positioned inline and anchor extents, leaving text and comments alone", async () => {
        const xml = drawing(
            "<d:inline> \n<!-- extent first --><d:extent/><d:effectExtent/><d:docPr/><d:cNvGraphicFramePr/><a:graphic/></d:inline>" +
                "<d:anchor><d:simplePos/><d:positionH/><d:positionV/> \n<d:extent/><d:effectExtent/><d:wrapNone/><d:docPr/><a:graphic/></d:anchor>",
        );
        const ctx = ValidationContext.fromEntries({ "word/document.xml": xml });
        expect(await check.run(ctx)).toEqual([]);
        expect(ctx.text("word/document.xml")).toBe(xml);
    });

    it("does not report missing extents or missing anchor positioning as misplaced extents", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": drawing("<d:inline/><d:anchor><d:extent/><a:graphic/></d:anchor>"),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("uses namespace URIs for extents, parents, and positioning siblings", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/header1.xml": drawing("<d:inline><a:extent/></d:inline><d:anchor><d:extent/><a:positionV/></d:anchor>"),
            "word/header2.xml": drawing("<a:inline><d:extent/></a:inline>"),
            "word/header3.xml": drawing("<d:anchor><a:positionH/><d:extent/></d:anchor>"),
            "customXml/item1.xml": drawing("<d:extent/>"),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({ code: check.id, path: "word/header2.xml" }),
            expect.objectContaining({ code: check.id, path: "word/header3.xml" }),
        ]);
    });

    it("reports a document-root extent as having the document parent", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": `<d:extent xmlns:d="${WP}"/>` });
        expect(await check.run(ctx)).toEqual([expect.objectContaining({ code: check.id, message: expect.stringContaining("#document") })]);
    });

    it("skips unparseable story parts", async () => {
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": drawing("<d:inline>") });
        expect(ctx.xml("word/header1.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
