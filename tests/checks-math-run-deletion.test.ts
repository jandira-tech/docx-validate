/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { serializeXml } from "../src/lib/xml-helpers";
import { mathRunDeletionCheck } from "../src/scripts/office/checks/math-run-deletion";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const xml = (body: string): string => `<w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:x="urn:foreign">${body}</w:document>`;
const parts = ["word/document.xml", "word/header1.xml", "word/footer2.xml", "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"];

describe("word-math-run-deletion", () => {
    it("reports each direct WML deletion in a math run in every story, independently of prefix spelling", async () => {
        const defect = `<m:oMath><m:r><w:del w:id="1" w:author="A"><m:t>x</m:t></w:del><w:del w:id="2" w:author="A"/></m:r></m:oMath>`;
        const ctx = ValidationContext.fromEntries({
            ...Object.fromEntries(parts.map((part) => [part, xml(defect)])),
            "word/header2.xml": `<a:oMath xmlns:a="${M}" xmlns:b="${W}"><a:r><b:del b:id="3" b:author="A"/></a:r></a:oMath>`,
        });
        const issues = await mathRunDeletionCheck.run(ctx);
        const expectedPaths = [...parts.flatMap((part) => [part, part]), "word/header2.xml"].sort();
        expect(issues.map((issue) => issue.path).sort()).toEqual(expectedPaths);
        for (const issue of issues) {
            expect(issue).toMatchObject({ severity: "warning", code: "word-math-run-deletion" });
            expect(issue.message).toContain("m:r");
            expect(issue.message).toContain("w:del");
            expect(issue.message).toContain("CT_R");
            expect(issue.message).toContain("m:ctrlPr");
            expect(issue.message).toContain("wrap");
        }
    });

    it("accepts legal control-property marks, whole-run and whole-zone wrappers, and ordinary run content", async () => {
        // CT_CtrlPr -> EG_RPrMath; CT_OMath/CT_OMathArg -> EG_OMathElements
        // -> EG_PContentMath -> EG_RunLevelElts -> CT_RunTrackChange -> m:r.
        const mark = `w:id="1" w:author="A"`;
        const run = `<m:r><m:rPr/><w:rPr/><m:t>x</m:t><w:delText>y</w:delText></m:r>`;
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": xml(`<w:body><w:p><w:pPr><w:rPr><w:ins ${mark}/><w:del ${mark}/></w:rPr></w:pPr>
                <m:oMath><w:ins ${mark}>${run}</w:ins><w:del ${mark}>${run}</w:del>
                  <m:sSup><m:sSupPr><m:ctrlPr><w:ins ${mark}><w:del ${mark}><w:rPr/></w:del></w:ins></m:ctrlPr></m:sSupPr>
                    <m:e><w:ins ${mark}>${run}</w:ins><m:ctrlPr><w:del ${mark}><w:rPr/></w:del></m:ctrlPr></m:e>
                    <m:sup><w:del ${mark}>${run}</w:del></m:sup></m:sSup>
                </m:oMath>
                <w:ins ${mark}><m:oMath>${run}</m:oMath></w:ins>
                <w:del ${mark}><m:oMathPara><m:oMath>${run}</m:oMath></m:oMathPara></w:del>
                </w:p></w:body>`),
        });
        expect(await mathRunDeletionCheck.run(ctx)).toEqual([]);
    });

    it("matches both namespaces and the immediate parent, ignoring similar foreign names and other revisions", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": xml(`<x:r><w:del/></x:r><m:r><x:del/><m:del/><w:ins/></m:r><m:ctrlPr><w:del/></m:ctrlPr>`),
            "word/header1.xml": `<w:del xmlns:w="${W}"/>`,
        });
        expect(await mathRunDeletionCheck.run(ctx)).toEqual([]);
    });

    it("skips malformed XML without throwing and never changes source bytes or the parsed revision content", async () => {
        const source = xml(`<m:r><w:del w:id="7" w:author="A"><m:t>keep me</m:t></w:del></m:r>`);
        const ctx = ValidationContext.fromEntries({ "word/document.xml": source, "word/header1.xml": "<broken>" });
        expect(ctx.xml("word/header1.xml")?.ok).toBe(false);
        const parsed = ctx.xml("word/document.xml");
        expect(parsed?.ok).toBe(true);
        if (!parsed?.ok) {
            throw new Error("Expected parsed control");
        }
        const before = serializeXml(parsed.document);
        expect(await mathRunDeletionCheck.run(ctx)).toEqual([
            expect.objectContaining({ path: "word/document.xml", code: "word-math-run-deletion", severity: "warning" }),
        ]);
        expect(serializeXml(parsed.document)).toBe(before);
        expect(ctx.text("word/document.xml")).toBe(source);
        expect(await mathRunDeletionCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
    });
});
