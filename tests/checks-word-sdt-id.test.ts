/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, W, WS } from "./support/word-rule-test-xml";
import { wordSdtIdEmptyCheck } from "../src/scripts/office/checks/word-sdt-id";

describe("word-sdt-id-empty", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<w:sdt><w:sdtPr><w:id w:val=""/></w:sdtPr></w:sdt>') });
        expect(await wordSdtIdEmptyCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-sdt-id-empty",
                path: "word/document.xml",
                message: 'A w:sdt has w:id w:val="". Microsoft Word cannot open the file. Set w:val to a number.',
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                '<w:sdtPr><w:id w:val="42"/><w:id/></w:sdtPr><w:p><w:id w:val=""/></w:p><a:sdtPr><w:id w:val=""/></a:sdtPr>',
            ),
        });
        expect(await wordSdtIdEmptyCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordSdtIdEmptyCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordSdtIdEmptyCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("ignores ids at a part root and handles Strict namespaces", async () => {
    expect(
        await wordSdtIdEmptyCheck.run(ValidationContext.fromEntries({ "word/document.xml": `<w:id xmlns:w="${W}" w:val=""/>` })),
    ).toEqual([]);
    expect(
        await wordSdtIdEmptyCheck.run(
            ValidationContext.fromEntries({ "word/document.xml": document(`<w:sdtPr><w:id w:val=""/></w:sdtPr>`, WS) }),
        ),
    ).toHaveLength(1);
});
