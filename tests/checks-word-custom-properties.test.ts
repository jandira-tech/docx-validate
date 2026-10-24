/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { wordCustomPropertyFiletimeEmptyCheck as check } from "../src/scripts/office/checks/word-custom-properties";

const CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const properties = (children: string) => `<Properties xmlns="${CP}" xmlns:v="${VT}">${children}</Properties>`;

describe("empty custom-property filetime", () => {
    it("rejects empty dateTime scalars directly inside custom properties", async () => {
        const ctx = ValidationContext.fromEntries({
            "docProps/custom.xml": properties("<property><v:filetime/></property><property><v:filetime> \t\r\n</v:filetime></property>"),
        });
        expect(await check.run(ctx)).toEqual([
            expect.objectContaining({
                code: check.id,
                severity: "error",
                path: "docProps/custom.xml",
                message: expect.stringMatching(/filetime.*property.*non-empty/),
            }),
            expect.objectContaining({ code: check.id, severity: "error", path: "docProps/custom.xml" }),
        ]);
    });

    it("preserves the empty lpwstr control Word opens in complex2.docx and populated filetimes", async () => {
        const xml = properties("<property><v:lpwstr/></property><property><v:filetime>2026-10-23T12:00:00Z</v:filetime></property>");
        const ctx = ValidationContext.fromEntries({ "docProps/custom.xml": xml });
        expect(await check.run(ctx)).toEqual([]);
        expect(ctx.text("docProps/custom.xml")).toBe(xml);
    });

    it("does not infer blocking defects for other scalar types or other parents and namespaces", async () => {
        const ctx = ValidationContext.fromEntries({
            "docProps/custom.xml": properties(
                "<property><v:date/><v:i4/><v:bool/><filetime/></property><other><v:filetime/></other>" +
                    "<v:property><v:filetime/></v:property><property><v:filetime>\u00a0</v:filetime></property>",
            ),
            "word/document.xml": properties("<property><v:filetime/></property>"),
        });
        expect(await check.run(ctx)).toEqual([]);
        expect(await check.run(ValidationContext.fromEntries({ "docProps/custom.xml": `<v:filetime xmlns:v="${VT}"/>` }))).toEqual([]);
    });

    it("skips missing and unparseable custom properties", async () => {
        expect(await check.run(ValidationContext.fromEntries({}))).toEqual([]);
        const ctx = ValidationContext.fromEntries({ "docProps/custom.xml": properties("<property>") });
        expect(ctx.xml("docProps/custom.xml")?.ok).toBe(false);
        expect(await check.run(ctx)).toEqual([]);
    });
});
