/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, CT } from "./support/word-rule-test-xml";
import { wordContentTypeInvalidCheck } from "../src/scripts/office/checks/word-content-types";

describe("word-content-type-invalid", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "[Content_Types].xml": `<Types xmlns="${CT}"><Override PartName="/word/document.xml" ContentType="invalid-document"/></Types>`,
        });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-content-type-invalid",
                path: "[Content_Types].xml",
                message: "Microsoft Word refuses content type 'invalid-document' for /word/document.xml",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "[Content_Types].xml": `<Types xmlns="${CT}"><Override PartName="/word/document.xml" ContentType="application/xml"/><Default Extension="xml"/></Types>`,
        });
        expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordContentTypeInvalidCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordContentTypeInvalidCheck.run(
                ValidationContext.fromEntries({ "[Content_Types].xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("checks default extensions and leaves nonmatching/absent types alone", async () => {
    const ctx = ValidationContext.fromEntries({
        "[Content_Types].xml": `<Types xmlns="${CT}"><Default Extension="bin" ContentType="invalid-bin"/><Default/><Override ContentType="invalid-part"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
        {
            severity: "error",
            code: "word-content-type-invalid",
            path: "[Content_Types].xml",
            message: "Microsoft Word refuses content type 'invalid-part' for ",
        },
        {
            severity: "error",
            code: "word-content-type-invalid",
            path: "[Content_Types].xml",
            message: "Microsoft Word refuses content type 'invalid-bin' for bin",
        },
    ]);
});

it.each(["[CONTENT_TYPES].XML", "[content_types].xml"])("finds invalid types in stored stream %s (#396)", async (name) => {
    const ctx = ValidationContext.fromEntries({
        [name]: `<Types xmlns="${CT}"><Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([
        {
            severity: "error",
            code: "word-content-type-invalid",
            path: name,
            message: "Microsoft Word refuses content type 'invalid-type' for /word/x.xml",
        },
    ]);
});

it("does not read a percent-escaped name as the stream (#396)", async () => {
    const ctx = ValidationContext.fromEntries({
        "%5BContent_Types%5D.xml": `<Types xmlns="${CT}"><Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`,
    });
    expect(await wordContentTypeInvalidCheck.run(ctx)).toEqual([]);
});
