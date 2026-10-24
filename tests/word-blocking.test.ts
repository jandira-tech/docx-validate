/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { canonicalPartName } from "../src/lib/content-types";
import type { ValidationIssue } from "../src/lib/types";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";

const issue = (code: string, rest: Partial<ValidationIssue> = {}): ValidationIssue => ({ severity: "error", code, message: "", ...rest });

describe("isWordBlockingIssue", () => {
    it("blocks relationships structure in the package relationships part and under word/ (#433)", () => {
        // Word 16.115 offers recovery for a foreign element in _rels/.rels and in
        // word/_rels/document.xml.rels (session 4c); other sidecars are unobserved.
        for (const path of [
            "_rels/.rels",
            "_RELS/.RELS",
            "_rels/.%72els",
            "_rels\\.rels",
            "word/_rels/document.xml.rels",
            "WORD/_RELS/document.xml.RELS",
            "word\\_rels\\document.xml.rels",
            "word/_rels/.rels",
        ]) {
            expect(isWordBlockingIssue(issue("opc-relationships-structure", { path })), path).toBe(true);
        }
        for (const path of [undefined, "_rels/a.rels", "customXml/_rels/item1.xml.rels", "docProps/_rels/app.xml.rels"]) {
            expect(isWordBlockingIssue(issue("opc-relationships-structure", { path })), path).toBe(false);
        }
    });

    it.each([
        "comment-thread-commentid-paraid-orphan",
        "comment-thread-commentid-missing-paraid",
        "comment-thread-commentid-missing-durableid",
        "comment-thread-commentid-duplicate-paraid",
        "comment-thread-commentid-duplicate-durableid",
        "comment-thread-durableid-orphan",
        "comment-thread-durableid-missing",
        "comment-thread-durableid-duplicate",
        "id-durable-overflow",
        "word-math-parse",
        "word-math-spre-body",
        "word-content-type-invalid",
        "word-document-root",
        // Word 16.115 offers recovery for a package without [Content_Types].xml
        // and for one naming no main document part (#435, #380; session 3c in
        // docs/plans/evidence/word-probes-2026-10-24.md).
        "ct-missing",
        "word-main-document-part",
        "ct-uncovered-part",
        "word-custom-property-filetime-empty",
        "word-drawing-extent-missing",
        "word-drawing-extent-misplaced",
        "word-picture-nvpicpr-missing",
        "word-picture-sppr-missing",
        "word-run-misplaced",
        "word-run-properties-misplaced",
        "word-underline-misplaced",
        "word-paragraph-properties-misplaced",
        "word-section-properties-misplaced",
        "word-table-properties-misplaced",
        "word-unknown-element",
        "word-link-without-namespace",
        "word-drawing-scalar-whitespace",
        "word-drawing-arc-radius",
        "word-empty-drawing",
        "word-drawing-extent-negative",
        "word-anchor-missing-attribute",
        "word-position-offset-and-align",
        "word-sdt-id-empty",
        "word-attached-toolbar-utf8",
        "word-row-without-cell",
        "word-cell-ends-with-table",
        "word-textbox-deleted-text",
        "word-move-from-deleted-text",
        "word-choice-requires-undeclared",
        "word-text-contains-paragraph",
        "word-hyperlink-relationship",
        "word-duplicate-theme-relationship",
        "word-duplicate-font-table-relationship",
        "word-duplicate-web-settings-relationship",
        "word-duplicate-settings-relationship",
        "word-numbering-picture-bullet",
        "word-footnote-reference-missing",
        "documentbuilder-insert-unresolved",
        "docprops-boolean-invalid",
        "mce-malformed",
        "mce-must-understand",
        // Word 16.115 offers recovery for every duplicate spelling probed (#351).
        "opc-equivalent-part-names",
    ])("keeps %s an error wherever it is found", (code) => {
        expect(isWordBlockingIssue(issue(code))).toBe(true);
    });

    it.each(["run-content-misplaced", "comment-orphan-start", "rels-unreferenced", "paraid-missing-element", "xsd-error", ""])(
        "tolerates %j",
        (code) => {
            expect(isWordBlockingIssue(issue(code))).toBe(false);
        },
    );

    it.each([
        ["word/document.xml", true],
        // OPC part names compare ASCII case-insensitively, with escapes of
        // unreserved characters decoded: these all name a part under word/ (#392).
        ["WORD/document.xml", true],
        ["Word/Comments.XML", true],
        ["%77ord/document.xml", true],
        ["wordy/document.xml", false],
        ["DOCPROPS/APP.XML", false],
        ["[Content_Types].xml", true],
        // The content-types stream is a ZIP item, matched as System.IO.Packaging
        // matches it: ASCII case-insensitively, with no escape decoding (review Q).
        ["[CONTENT_TYPES].XML", true],
        ["[content_types].xml", true],
        ["%5BContent_Types%5D.xml", false],
        // Word 16.115 offers recovery for a zero-byte, truncated or
        // undeclared-prefix _rels/.rels (#388); matched by OPC part name.
        ["_rels/.rels", true],
        ["_RELS/.RELS", true],
        ["_rels/.%72els", true],
        ["_rels/x.rels", false],
        ["docProps/app.xml", false],
        ["docProps/custom.xml", false],
        ["metadata/core.xml", false],
        [undefined, false],
    ] as const)("judges xml-syntax, xml-undeclared-prefix and xml-encoding in %s as blocking: %s", (path, blocking) => {
        expect(isWordBlockingIssue(issue("xml-syntax", { path }))).toBe(blocking);
        expect(isWordBlockingIssue(issue("xml-undeclared-prefix", { path }))).toBe(blocking);
        // Word 16.115 refuses a word/document.xml declaring an unknown encoding (#397).
        expect(isWordBlockingIssue(issue("xml-encoding", { path }))).toBe(blocking);
    });

    describe.each(["xml-syntax", "xml-encoding", "xml-undeclared-prefix"])("%s on properties parts (#409)", (code) => {
        it.each([
            "docProps/core.xml",
            "DOCPROPS/CORE.XML",
            "docProps/%63ore.xml",
            "docProps/app.xml",
            "DOCPROPS/APP.XML",
            "docProps/%61pp.xml",
            "docProps/custom.xml",
            "DOCPROPS/CUSTOM.XML",
            "docProps/%63ustom.xml",
            "props/meta.xml",
            "PROPS/%6Deta.xml",
        ])("blocks %s only when named by a properties relationship", (path) => {
            const properties = new Set(
                ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml", "props/meta.xml"].map(canonicalPartName),
            );
            expect.soft(isWordBlockingIssue(issue(code, { path }), properties), "referenced").toBe(true);
            expect.soft(isWordBlockingIssue(issue(code, { path }), new Set()), "unreferenced").toBe(false);
        });
    });

    it.each([
        ["word/document.xml", true],
        ["word/header1.xml", false],
        ["word/FOOTER.xml", false],
        [undefined, true],
    ] as const)("judges a missing .rels sidecar for %s as blocking: %s", (path, blocking) => {
        expect(isWordBlockingIssue(issue("rels-missing-sidecar", { path }))).toBe(blocking);
    });

    it.each([
        ["Broken reference: ../customXml/item1.xml", true],
        ["Broken reference in word/_rels/document.xml.rels", true],
        ["Broken reference in _rels/.rels", true],
        ["Broken reference: media/image1.png", false],
    ])("judges rels-broken %j as blocking: %s", (message, blocking) => {
        expect(isWordBlockingIssue(issue("rels-broken", { message }))).toBe(blocking);
    });

    it.each([
        ["Relationship is missing required attribute Target", true],
        ["Relationship has an empty Id", false],
    ])("judges rels-empty-element %j as blocking: %s", (message, blocking) => {
        expect(isWordBlockingIssue(issue("rels-empty-element", { message }))).toBe(blocking);
    });
});
