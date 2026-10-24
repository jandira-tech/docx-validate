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
        ["word/document.xml", "main document", ["word/document.xml"], true],
        ["WORD/document.xml", "main document, case-insensitively", ["word/document.xml"], true],
        ["word/styles.xml", "eagerly parsed styles part", ["word/styles.xml"], true],
        ["customXml/item1.xml", "eagerly parsed customXml item outside word/", ["customXml/item1.xml"], true],
        ["word/footer1.xml", "lazily parsed footer the document uses", ["word/footer1.xml"], true],
        ["word/theme1.xml", "unobserved type under word/", ["word/theme1.xml"], true],
        ["word/footer1.xml", "lazily parsed footer the document does not use", ["word/styles.xml"], false],
        ["word/comments.xml", "lazily parsed comments part the document does not use", [], false],
        ["word/xtra.xml", "part named by no relationship", [], false],
        ["word/xtra.xml", "part named only by an unknown relationship type", ["word/document.xml"], false],
        ["word/document.xml", "part named by nothing, even as the usual main spelling", [], false],
    ] as const)("judges xml-syntax in %s as blocking when reached as the %s: %s", (path, _why, readParts, blocking) => {
        const read = new Set(readParts.map(canonicalPartName));
        expect(isWordBlockingIssue(issue("xml-syntax", { path }), new Set(), read)).toBe(blocking);
        expect(isWordBlockingIssue(issue("xml-undeclared-prefix", { path }), new Set(), read)).toBe(blocking);
        // Word 16.115 refuses a word/document.xml declaring an unknown encoding (#397).
        expect(isWordBlockingIssue(issue("xml-encoding", { path }), new Set(), read)).toBe(blocking);
    });

    it.each([
        ["word/document.xml", false],
        ["Word/Comments.XML", false],
        ["%77ord/document.xml", false],
        ["wordy/document.xml", false],
        ["DOCPROPS/APP.XML", false],
        ["docProps/app.xml", false],
        ["docProps/custom.xml", false],
        ["metadata/core.xml", false],
        ["word/_rels/document.xml.rels", false],
        [undefined, false],
    ] as const)("no longer judges xml-syntax in %s blocking by folder alone: %s", (path, blocking) => {
        // #437: lying under word/ no longer blocks; the content-types stream,
        // the package relationships part and the properties parts still do.
        expect(isWordBlockingIssue(issue("xml-syntax", { path }), new Set(), new Set())).toBe(blocking);
        expect(isWordBlockingIssue(issue("xml-undeclared-prefix", { path }), new Set(), new Set())).toBe(blocking);
        expect(isWordBlockingIssue(issue("xml-encoding", { path }), new Set(), new Set())).toBe(blocking);
    });

    it("judges xml-syntax, xml-undeclared-prefix and xml-encoding in the always-blocking parts", () => {
        for (const path of [
            "[Content_Types].xml",
            "[CONTENT_TYPES].XML",
            "[content_types].xml",
            "_rels/.rels",
            "_RELS/.RELS",
            "_rels/.%72els",
        ]) {
            expect.soft(isWordBlockingIssue(issue("xml-syntax", { path }), new Set(), new Set()), path).toBe(true);
            expect.soft(isWordBlockingIssue(issue("xml-undeclared-prefix", { path }), new Set(), new Set()), path).toBe(true);
            expect.soft(isWordBlockingIssue(issue("xml-encoding", { path }), new Set(), new Set()), path).toBe(true);
        }
        // The content-types stream is a ZIP item, matched as System.IO.Packaging
        // matches it: ASCII case-insensitively, with no escape decoding (review Q).
        expect(isWordBlockingIssue(issue("xml-syntax", { path: "%5BContent_Types%5D.xml" }), new Set(), new Set())).toBe(false);
        expect(isWordBlockingIssue(issue("xml-syntax", { path: "_rels/x.rels" }), new Set(), new Set())).toBe(false);
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
            expect.soft(isWordBlockingIssue(issue(code, { path }), properties, new Set()), "referenced").toBe(true);
            expect.soft(isWordBlockingIssue(issue(code, { path }), new Set(), new Set()), "unreferenced").toBe(false);
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

    describe("xml-syntax by reachability, not by folder (#437)", () => {
        const reach = (parts: readonly string[]): ReadonlySet<string> => new Set(parts.map(canonicalPartName));
        const noReach = reach([]);

        it.each([
            // Word 16.115 opens these clean: a part named by nothing, by an
            // unknown relationship type, or by a lazy type the document does
            // not use (sessions 3c, 5, 8b, 9a).
            ["word/xtra.xml", noReach, false],
            ["word/footer1.xml", noReach, false],
            ["word/comments.xml", noReach, false],
            ["customXml/item9.xml", noReach, false],
            // The main document blocks wherever it is stored (rule 2).
            ["maindoc/document.xml", reach(["maindoc/document.xml"]), true],
            ["maindoc/document.xml", noReach, false],
            // Eagerly parsed relationship targets block, used or not (rule 3).
            ["customXml/item1.xml", reach(["customXml/item1.xml"]), true],
            ["customXml/item1.xml", noReach, false],
            ["word/numbering.xml", reach(["word/numbering.xml"]), true],
            // A lazily parsed target blocks once the document uses it (rule 4).
            ["word/footer1.xml", reach(["word/footer1.xml"]), true],
            ["word/footnotes.xml", reach(["word/footnotes.xml"]), true],
            // Anything under word/ named by a known Office relationship type
            // of an unobserved kind keeps blocking (rule 5).
            ["word/header1.xml", reach(["word/header1.xml"]), true],
            ["word/theme/theme1.xml", reach(["word/theme/theme1.xml"]), true],
            ["word/theme/theme1.xml", noReach, false],
            ["word/header1.xml", noReach, false],
        ] as const)("judges %s with the given reachability as blocking: %s", (path, wordRead, blocking) => {
            for (const code of ["xml-syntax", "xml-undeclared-prefix", "xml-encoding"] as const) {
                expect.soft(isWordBlockingIssue(issue(code, { path }), new Set(), wordRead), code).toBe(blocking);
            }
        });

        it("compares the reached part names as OPC part names", () => {
            const wordRead = reach(["customXml/item1.xml", "maindoc/document.xml"]);
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "CUSTOMXML/%69tem1.xml" }), new Set(), wordRead)).toBe(true);
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "MAINDOC/document.xml" }), new Set(), wordRead)).toBe(true);
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "customXml/item2.xml" }), new Set(), wordRead)).toBe(false);
        });

        it("warns on a relationships sidecar under word/ that no relationship names", () => {
            // A sidecar is never a relationship target, so none of the five
            // reachability rules covers it (#437). Its structural defects keep
            // their own blocking code (`opc-relationships-structure`, #433).
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "word/_rels/document.xml.rels" }), new Set(), noReach)).toBe(false);
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "customXml/_rels/item1.xml.rels" }), new Set(), noReach)).toBe(false);
        });
    });
});
