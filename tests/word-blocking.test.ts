/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";

const issue = (code: string, rest: Partial<ValidationIssue> = {}): ValidationIssue => ({ severity: "error", code, message: "", ...rest });

describe("isWordBlockingIssue", () => {
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
        [undefined, false],
    ] as const)("judges xml-syntax, xml-undeclared-prefix and xml-encoding in %s as blocking: %s", (path, blocking) => {
        expect(isWordBlockingIssue(issue("xml-syntax", { path }))).toBe(blocking);
        expect(isWordBlockingIssue(issue("xml-undeclared-prefix", { path }))).toBe(blocking);
        // Word 16.115 refuses a word/document.xml declaring an unknown encoding (#397).
        expect(isWordBlockingIssue(issue("xml-encoding", { path }))).toBe(blocking);
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
