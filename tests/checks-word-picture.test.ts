/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { wordPictureNvpicprMissingCheck as check, wordPictureSpprMissingCheck } from "../src/scripts/office/checks/word-picture";

const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const picture = (children: string) => `<p:pic xmlns:p="${PIC}" xmlns:a="${A}">${children}</p:pic>`;

describe("missing picture shape properties", () => {
    it("rejects pic:pic without its direct spPr, as in header-with-rid-no-sidecar.docx", async () => {
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": picture("<p:nvPicPr/><p:blipFill/>") });
        expect(await wordPictureSpprMissingCheck.run(ctx)).toEqual([
            expect.objectContaining({
                code: wordPictureSpprMissingCheck.id,
                severity: "error",
                path: "word/header1.xml",
                message: expect.stringMatching(/spPr.*pic:pic.*after.*blipFill/),
            }),
        ]);
    });

    it("accepts the complete picture control and leaves the XML untouched", async () => {
        const xml = picture("<p:nvPicPr/><p:blipFill/><p:spPr/>");
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": xml });
        expect(await wordPictureSpprMissingCheck.run(ctx)).toEqual([]);
        expect(ctx.text("word/header1.xml")).toBe(xml);
    });

    it("requires a direct spPr in the picture namespace", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": picture("text<p:nvPicPr/><a:spPr/><p:blipFill><p:spPr/></p:blipFill>"),
        });
        expect(await wordPictureSpprMissingCheck.run(ctx)).toEqual([expect.objectContaining({ code: wordPictureSpprMissingCheck.id })]);
    });

    it("ignores foreign pictures, non-story parts, and unparseable story parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": `<a:pic xmlns:a="${A}"/>`,
            "word/header1.xml": picture("<p:nvPicPr>"),
            "customXml/item1.xml": picture(""),
        });
        expect(ctx.xml("word/header1.xml")?.ok).toBe(false);
        expect(await wordPictureSpprMissingCheck.run(ctx)).toEqual([]);
    });
});

describe("missing picture non-visual properties", () => {
    it("rejects pic:pic without its direct nvPicPr, as in header-with-rid-no-sidecar.docx", async () => {
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": picture("<p:blipFill/><p:spPr/>") });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "word/header1.xml",
                message: expect.stringMatching(/nvPicPr.*pic:pic.*before.*blipFill/),
            }),
        ]);
    });

    it("accepts the complete header scaffold control without changing its XML", async () => {
        const xml = picture("<p:nvPicPr/><p:blipFill/><p:spPr/>");
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": xml });
        expect(await check.run(ctx)).toEqual([]);
        expect(ctx.text("word/header1.xml")).toBe(xml);
    });

    it("does not mistake a nested or foreign nvPicPr for a direct picture child", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": picture("text<a:nvPicPr/><p:blipFill><p:nvPicPr/></p:blipFill><p:spPr/>"),
        });
        expect(await check.run(ctx)).toEqual([expect.objectContaining({ code: check.id })]);
    });

    it("ignores foreign picture names and non-story parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": `<a:pic xmlns:a="${A}"/>`,
            "customXml/item1.xml": picture(""),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it("skips unparseable story parts", async () => {
        const ctx = ValidationContext.fromEntries({ "word/header1.xml": picture("<p:blipFill>") });
        expect(ctx.xml("word/header1.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
