/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, W, WS } from "./support/word-rule-test-xml";
import { wordMathParseCheck, wordMathSpreBodyCheck } from "../src/scripts/office/checks/word-body-math";

describe("word-math-parse", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": "<w:document" });
        const parsed = ctx.xml("word/document.xml");
        expect(parsed?.ok).toBe(false);
        expect(await wordMathParseCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-math-parse",
                path: "word/document.xml",
                message: `Error parsing XML: ${parsed && !parsed.ok ? parsed.message : ""}`,
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("") });
        expect(await wordMathParseCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordMathParseCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordMathParseCheck.run(
                ValidationContext.fromEntries({ "word/other.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-math-spre-body", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<m:oMathPara><m:oMath><m:sPre/></m:oMath></m:oMathPara>"),
        });
        expect(await wordMathSpreBodyCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-math-spre-body",
                path: "word/document.xml",
                message:
                    "Microsoft Word refuses to open documents with a body-level <m:oMathPara> that contains <m:sPre>; wrap it in a paragraph or avoid m:sPre in display math.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                "text<w:p><m:oMathPara><m:sPre/></m:oMathPara></w:p><m:oMathPara><m:r/></m:oMathPara><a:oMathPara/>",
            ),
        });
        expect(await wordMathSpreBodyCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordMathSpreBodyCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordMathSpreBodyCheck.run(
                ValidationContext.fromEntries({ "word/other.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("leaves absent bodies, non-Word document parts and malformed math to their own checks", async () => {
    const packages: Record<string, string>[] = [
        { "customXml/document.xml": document(`<m:oMathPara><m:sPre/></m:oMathPara>`) },
        { "word/document.xml": `<w:document xmlns:w="${W}"/>` },
        { "word/document.xml": "<w:document" },
    ];
    for (const entries of packages) {
        expect(await wordMathSpreBodyCheck.run(ValidationContext.fromEntries(entries))).toEqual([]);
    }
    expect(
        await wordMathSpreBodyCheck.run(
            ValidationContext.fromEntries({ "word/document.xml": document(`<m:oMathPara><m:sPre/></m:oMathPara>`, WS) }),
        ),
    ).toHaveLength(1);
});
