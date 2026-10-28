/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { canonicalPartName } from "../src/lib/content-types";
import type { ValidationIssue } from "../src/lib/types";
import { parseXml } from "../src/lib/xml-helpers";
import { registerRelsBrokenTarget, registerRelsEmptyElementShape } from "../src/scripts/office/rels-finding-facts";
import {
    hasWordRelationshipReference,
    isWordBlockingIssue,
    namedByNoRelationship,
    unobservedReferenceAttributeValues,
} from "../src/scripts/office/validators/word-blocking";

const issue = (code: string, rest: Partial<ValidationIssue> = {}): ValidationIssue => ({ severity: "error", code, message: "", ...rest });

describe("isWordBlockingIssue", () => {
    it.each(["_rels/.rels", "word/_rels/document.xml.rels"])("allows markup removed from relationships at %s (#632)", (path) => {
        expect(isWordBlockingIssue(issue("opc-relationships-markup-compatibility", { path }))).toBe(false);
    });

    it("blocks duplicate relationship Ids in the package and story-root sidecars (#615)", () => {
        // #617 renamed this set to the story roots: every main document plus
        // every reached glossary document. This case keeps a mains-only set.
        const roots = new Set([canonicalPartName("word/main.xml"), canonicalPartName("stories/body.xml")]);
        const read = new Set(["word/_rels/header9.xml.rels"]);
        for (const path of ["_rels/.rels", "_RELS/.RELS", "_rels\\.rels", "word/_rels/main.xml.rels", "STORIES/_RELS/body.xml.RELS"]) {
            expect.soft(isWordBlockingIssue(issue("rels-id-duplicate", { path }), new Set(), read, roots), path).toBe(true);
        }
        for (const path of [
            undefined,
            "word/_rels/document.xml.rels",
            "word/_rels/header9.xml.rels",
            "word/glossary/_rels/document.xml.rels",
        ]) {
            expect.soft(isWordBlockingIssue(issue("rels-id-duplicate", { path }), new Set(), read, roots), path).toBe(false);
        }
    });

    it("blocks duplicate relationship Ids in a reached glossary's own sidecar (#617)", () => {
        // Session 30 (s30a): Word 16.115 offered recovery for the rIdZ pair in
        // a reached glossary's sidecar, so a glossary a main-document
        // glossaryDocument relationship reaches is a story root whose sidecar
        // blocks; an unreached glossary's sidecar keeps its word-valid warning.
        const roots = new Set([canonicalPartName("word/document.xml"), canonicalPartName("word/glossary/document.xml")]);
        for (const path of ["word/glossary/_rels/document.xml.rels", "WORD/GLOSSARY/_RELS/document.xml.RELS"]) {
            expect.soft(isWordBlockingIssue(issue("rels-id-duplicate", { path }), new Set(), new Set(), roots), path).toBe(true);
        }
        const mainsOnly = new Set([canonicalPartName("word/document.xml")]);
        expect(
            isWordBlockingIssue(
                issue("rels-id-duplicate", { path: "word/glossary/_rels/document.xml.rels" }),
                new Set(),
                new Set(),
                mainsOnly,
            ),
        ).toBe(false);
        expect(
            isWordBlockingIssue(issue("rels-id-duplicate", { path: "word/_rels/header9.xml.rels" }), new Set(), new Set(), mainsOnly),
        ).toBe(false);
    });

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
        // A wne:toolbarData r:id naming a relationship of another Type, or
        // none, gets the recovery offer whatever the target holds (#716, session 48).
        "word-attached-toolbar-relationship",
        "word-row-without-cell",
        "word-cell-ends-with-table",
        "word-textbox-deleted-text",
        "word-move-from-deleted-text",
        "word-choice-requires-undeclared",
        "word-text-contains-paragraph",
        "word-hyperlink-relationship",
        // Word 16.115 offers recovery for every story-reference type mismatch probed (#606, session 26).
        "word-header-footer-relationship",
        // Word 16.115 offers recovery for a c:chart, dgm:relIds r:dm or o:OLEObject
        // whose r:id names a relationship of another type, or none (#705,
        // sessions 45 and 46).
        "word-part-reference-relationship",
        "word-duplicate-theme-relationship",
        "word-duplicate-font-table-relationship",
        "word-duplicate-web-settings-relationship",
        "word-duplicate-settings-relationship",
        // Two styles or two numbering relationships in one .rels part make
        // Word offer recovery (#694, session 45).
        "word-duplicate-styles-relationship",
        "word-duplicate-numbering-relationship",
        // Word session 62 (s62c, s62e) offers recovery for two endnotes
        // relationships in the main document's sidecar when the body
        // references an endnote; the same pair with no reference opens
        // clean (#807).
        "word-duplicate-endnotes-relationship",
        // Word session 63 (s63d) offers recovery for two footnotes
        // relationships in the main document's sidecar when the body
        // references a footnote; the unreferenced pair opened clean
        // (session 45 s45g, #827).
        "word-duplicate-footnotes-relationship",
        "word-numbering-picture-bullet",
        "word-footnote-reference-missing",
        // Word session 66 (s66a) offers recovery for a comments part both
        // the main document and a reached glossary use, and session 67
        // (s67a) for the shared footnotes part and its endnotes twin (#850).
        "word-shared-story-part",
        "documentbuilder-insert-unresolved",
        "docprops-boolean-invalid",
        // Word 16.115 offers recovery for every duplicate spelling probed (#351).
        "opc-equivalent-part-names",
    ])("keeps %s an error wherever it is found", (code) => {
        expect(isWordBlockingIssue(issue(code))).toBe(true);
    });

    it.each(["word/document.xml", "WORD/footnotes.XML", "word/footnotes.%78ml"])(
        "keeps the markup-compatibility errors of %s an error under word-valid",
        (path) => {
            // The part key compares canonical part names, so the escaped
            // spelling of a word/*.xml part blocks as its twin does (#639).
            for (const code of ["mce-malformed", "mce-must-understand"]) {
                expect(isWordBlockingIssue(issue(code, { path })), `${code} ${path}`).toBe(true);
            }
        },
    );

    it.each(["mce-malformed", "mce-must-understand"])(
        "%s blocks package and reached story-root sidecars by OPC identity (#644, #639)",
        (code) => {
            const roots = new Set(["word/doc2.xml", "word/glossary/document.xml", "stories/main.xml"]);
            for (const path of [
                "_rels/.rels",
                "_RELS/.RELS",
                "_rels\\.rels",
                "_rels/.%72els",
                "word/_rels/doc2.xml.rels",
                "WORD/_RELS/doc2.xml.RELS",
                "word/_rels\\doc2.xml.rels",
                "word/_rels/d%6Fc2.xml.rels",
                "word/glossary/_rels/document.xml.rels",
                "stories/_rels/main.xml.rels",
            ]) {
                expect.soft(isWordBlockingIssue(issue(code, { path }), new Set(), new Set(), roots), path).toBe(true);
            }
            for (const path of [
                undefined,
                "word/_rels/document.xml.rels",
                "word/_rels/styles.xml.rels",
                "word/_rels/footer1.xml.rels",
                "word/_rels/header1.xml.rels",
                "word/_rels/settings.xml.rels",
                "word/_rels/numbering.xml.rels",
                "word/glossary/glossary/_rels/document.xml.rels",
                "docProps/_rels/.rels",
                "word/_rels/doc2.xml%2Frels",
            ]) {
                expect.soft(isWordBlockingIssue(issue(code, { path }), new Set(), new Set(), roots), String(path)).toBe(false);
            }
            // No reached story roots means even the default main sidecar is stray.
            expect(isWordBlockingIssue(issue(code, { path: "word/_rels/document.xml.rels" }))).toBe(false);
            expect(
                isWordBlockingIssue(
                    issue(code, { path: "word/_rels/document.xml.rels" }),
                    new Set(),
                    new Set(),
                    new Set(["word/document.xml"]),
                ),
            ).toBe(true);
        },
    );

    it.each(["mce-malformed", "mce-must-understand"])(
        "%s warns in a word/ part no relationship names, and blocks in one a relationship names (#1116, session 83)",
        (code) => {
            for (const path of ["word/header83.xml", "WORD/header83.XML", "word/header%38%33.xml", "word\\header83.xml"]) {
                const finding = issue(code, { path });
                expect
                    .soft(isWordBlockingIssue(finding, new Set(), new Set(), new Set(), new Set(), { namedByNoRelationship: true }), path)
                    .toBe(false);
                expect
                    .soft(isWordBlockingIssue(finding, new Set(), new Set(), new Set(), new Set(), { namedByNoRelationship: false }), path)
                    .toBe(true);
                expect.soft(isWordBlockingIssue(finding), path).toBe(true);
            }
            // No relationship names a relationships part, so the fact must not
            // reach the sidecar half: there the story roots decide (#644).
            const roots = new Set(["word/document.xml"]);
            for (const path of ["_rels/.rels", "word/_rels/document.xml.rels"]) {
                expect
                    .soft(
                        isWordBlockingIssue(issue(code, { path }), new Set(), new Set(), roots, new Set(), { namedByNoRelationship: true }),
                        path,
                    )
                    .toBe(true);
            }
        },
    );

    it("reads a part as named by no relationship only when the named list is known and lacks it (#1116)", () => {
        const named = new Set([canonicalPartName("word/document.xml"), canonicalPartName("word/header1.xml")]);
        expect(namedByNoRelationship(issue("mce-malformed", { path: "word/header83.xml" }), named)).toBe(true);
        // OPC-equivalent spellings, and a backslash separator, name the listed part.
        for (const path of ["word/header1.xml", "WORD/Header1.XML", "word/header%31.xml", "word\\header1.xml"]) {
            expect.soft(namedByNoRelationship(issue("mce-malformed", { path }), named), path).toBe(false);
        }
        // An unknown list, or a finding with no part, says nothing.
        expect(namedByNoRelationship(issue("mce-malformed", { path: "word/header83.xml" }), undefined)).toBe(false);
        expect(namedByNoRelationship(issue("mce-malformed"), named)).toBe(false);
    });

    it("blocks an unrecognized content type only on the observed reach (#707, #718)", () => {
        // The verdict reads the observed set the callers resolve once per run
        // (`wordObservedContentTypeParts`), never the message: it embeds the
        // declared content type and the stored part name verbatim, so a
        // package could carry the recovery sentence itself and flip an
        // unobserved finding to blocking (#718).
        const observed = new Set([canonicalPartName("word/styles.bin")]);
        const blocking = issue("word-part-content-type-unrecognized", {
            path: "word/styles.bin",
            message:
                'word/styles.bin, named by the main document\'s styles relationship, is typed "application/octet-stream" by a Default; it expects "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml". Word 16.115 offers to recover the package (sessions 46 and 49).',
        });
        expect(isWordBlockingIssue(blocking, new Set(), new Set(), new Set(), observed)).toBe(true);
        // An OPC-equivalent spelling names the same part.
        expect(
            isWordBlockingIssue(
                issue("word-part-content-type-unrecognized", { path: "WORD/STYLES.BIN" }),
                new Set(),
                new Set(),
                new Set(),
                observed,
            ),
        ).toBe(true);
        const unobserved = issue("word-part-content-type-unrecognized", {
            path: "word/settings.bin",
            // The crafted sentence rides the message, not the reach.
            message:
                'word/settings.bin, named by the main document\'s settings relationship, is typed "application/x-Word 16.115 offers to recover the package" by an Override.',
        });
        expect(isWordBlockingIssue(unobserved, new Set(), new Set(), new Set(), observed)).toBe(false);
        const glossary = issue("word-part-content-type-unrecognized", {
            path: "word/glossary/styles.bin",
            message:
                'word/glossary/styles.bin, named by the glossary\'s styles relationship, is typed "application/octet-stream" by a Default; it expects "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the styles part.',
        });
        expect(isWordBlockingIssue(glossary, new Set(), new Set(), new Set(), observed)).toBe(false);
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
        ["word/theme1.xml", "eagerly parsed theme part", ["word/theme1.xml"], true],
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
        { path: "word/document.xml", read: false, story: true, reference: true, blocking: true },
        { path: "word/document.xml", read: true, story: false, reference: true, blocking: true },
        { path: "word/document.xml", read: true, story: true, reference: false, blocking: false },
        { path: "word/styles.xml", read: true, story: false, reference: false, blocking: false },
        { path: "word/comments.xml", read: true, story: false, reference: true, blocking: true },
        { path: "word/xtra.xml", read: false, story: false, reference: true, blocking: false },
        { path: "word/header1.xml", read: true, story: false, reference: true, blocking: false },
        { path: "word/FOOTER.xml", read: true, story: false, reference: true, blocking: false },
        { path: "word/header1.xml", read: false, story: true, reference: true, blocking: true },
        { path: "WORD/%68eader1.XML", read: false, story: true, reference: true, blocking: true },
        { path: "word/header1.xml", read: true, story: true, reference: false, blocking: false },
        { path: undefined, read: false, story: false, reference: true, blocking: false },
    ])("judges a missing sidecar with per-issue processed-reference facts: $path/$read/$story/$reference", (specimen) => {
        const canonical = canonicalPartName(specimen.path ?? "");
        expect(
            isWordBlockingIssue(
                issue("rels-missing-sidecar", { path: specimen.path, message: "arbitrary prose" }),
                new Set(),
                new Set(specimen.read ? [canonical] : []),
                new Set(specimen.story ? [canonical] : []),
                new Set(),
                { hasEffectiveRelationshipReference: specimen.reference },
            ),
        ).toBe(specimen.blocking);
    });

    it.each([
        { path: "charts/chart1.xml", unobserved: true, reference: true, blocking: true },
        { path: "word/charts/chart1.xml", unobserved: true, reference: true, blocking: true },
        { path: "charts/chart1.xml", unobserved: true, reference: false, blocking: false },
        { path: "charts/chart1.xml", unobserved: false, reference: true, blocking: false },
        { path: "word/header1.xml", unobserved: true, reference: true, blocking: false },
    ])(
        "keeps a missing sidecar blocking in a part an unobserved Office relationship names (#638): $path/$unobserved/$reference",
        (specimen) => {
            expect(
                isWordBlockingIssue(issue("rels-missing-sidecar", { path: specimen.path }), new Set(), new Set(), new Set(), new Set(), {
                    hasEffectiveRelationshipReference: specimen.reference,
                    reachedByUnobservedRelationship: specimen.unobserved,
                }),
            ).toBe(specimen.blocking);
        },
    );

    it.each([
        // #645, session 36 (s36b/s36d): a part every reaching relationship
        // names through `…/relationships/chart` blocks only when a Word-read
        // part's effective XML references one of those Ids.
        { path: "word/charts/chart1.xml", unobserved: true, reference: true, story: false, chart: "referenced", blocking: true },
        { path: "charts/chart1.xml", unobserved: true, reference: true, story: false, chart: "referenced", blocking: true },
        { path: "word/charts/chart1.xml", unobserved: true, reference: true, story: false, chart: "unreferenced", blocking: false },
        { path: "charts/chart1.xml", unobserved: true, reference: true, story: false, chart: "unreferenced", blocking: false },
        // A chart-only part that is a story root keeps the story-root
        // precedence (#625), and a genuine header/footer keeps its basename
        // tolerance even when referenced.
        { path: "charts/chart1.xml", unobserved: true, reference: true, story: true, chart: "unreferenced", blocking: true },
        { path: "word/header1.xml", unobserved: true, reference: true, story: false, chart: "referenced", blocking: false },
        // A non-chart part reached the same way stays as today: the
        // unreferenced chart-only rule never applies to it.
        { path: "word/xtra.xml", unobserved: true, reference: true, story: false, chart: "unreferenced", blocking: false },
        { path: "word/xtra.xml", unobserved: true, reference: true, story: false, chart: "referenced", blocking: true },
        // The fact stays undefined for a part any non-chart relationship
        // reaches, or for an unrelated part, so reach alone decides.
        { path: "word/charts/chart1.xml", unobserved: true, reference: true, story: false, chart: undefined, blocking: true },
        { path: "word/xtra.xml", unobserved: false, reference: true, story: false, chart: undefined, blocking: false },
    ] as const)("judges a chart-only missing sidecar by whether the relationship is referenced (#645): $path/$chart", (specimen) => {
        const canonical = canonicalPartName(specimen.path);
        expect(
            isWordBlockingIssue(
                issue("rels-missing-sidecar", { path: specimen.path }),
                new Set(),
                new Set(),
                new Set(specimen.story ? [canonical] : []),
                new Set(),
                {
                    hasEffectiveRelationshipReference: specimen.reference,
                    reachedByUnobservedRelationship: specimen.unobserved,
                    referenceLoadedReach: specimen.chart,
                },
            ),
        ).toBe(specimen.blocking);
    });

    it("does not infer a missing-sidecar reference from message prose or a part name (#638)", () => {
        expect(
            isWordBlockingIssue(
                issue("rels-missing-sidecar", { path: "word/document.xml", message: "r:id on w:footerReference" }),
                new Set(),
                new Set(["word/document.xml"]),
                new Set(["word/document.xml"]),
            ),
        ).toBe(false);
    });

    it.each([
        // The shapes Word was shown: vectors.docx's broken `../customXml/item1.xml`
        // (unreadable content) and a nested `.rels` sidecar, the package's own
        // relationships part included. The registered resolved part decides.
        ["customXml/item1.xml", true],
        ["CUSTOMXML/%69tem1.xml", true],
        ["customXml/_rels/item1.xml.rels", true],
        ["_rels/.rels", true],
        ["word/_rels/document.xml.rels", true],
        // #734's injected spellings: the raw Target carried `../customXml/`,
        // `/_rels/` or `.rels` while the part it names is none of those.
        ["word/media/_rels/x.png", false],
        ["word/notes.rels", false],
        ["word/customXml/x.png", false],
        ["media/image1.png", false],
        // A target naming a folder, registered with its final `/` (the package
        // root as `/`): Word refused each one it was shown, whatever the
        // relationship type and whether anything used it (sessions 70 to 73,
        // #975).
        ["word/styles.xml/", true],
        ["word/settings.xml/", true],
        ["word/media/image1.png/", true],
        ["word/header1.xml/", true],
        ["word/", true],
        ["docProps/core.xml/", true],
        ["/", true],
    ])("judges rels-broken resolving to %j as blocking: %s", (part, blocking) => {
        const found = issue("rels-broken", { path: "word/_rels/document.xml.rels", message: `Broken reference to ${part}` });
        registerRelsBrokenTarget(found, part);
        expect(isWordBlockingIssue(found, new Set(), new Set(), new Set(["word/document.xml"]))).toBe(blocking);
    });

    it.each([
        // The relationships parts Word reads to reach the stories: the
        // package's own and each story root's sidecar, compared as OPC
        // compares part names.
        ["_rels/.rels", true],
        ["_RELS/.rels", true],
        ["word/_rels/document.xml.rels", true],
        ["WORD/_rels/document.xml.rels", true],
        ["word/glossary/_rels/document.xml.rels", true],
        // Word opened a folder target clean in the sidecar of an unreferenced
        // part (session 79), a used and an unused header and the styles part
        // (session 80), and in no relationships part at all.
        ["word/_rels/stray.xml.rels", false],
        ["word/_rels/header1.xml.rels", false],
        ["word/_rels/styles.xml.rels", false],
        ["word/notes.rels", false],
        [undefined, false],
    ])("blocks a folder target reported in %j only where Word reads it: %s (#987)", (relsPart, blocking) => {
        const found = issue("rels-broken", { path: relsPart, message: "Broken reference to media/image1.png/." });
        registerRelsBrokenTarget(found, "word/media/image1.png/");
        const storyRoots = new Set(["word/document.xml", "word/glossary/document.xml"]);
        expect(isWordBlockingIssue(found, new Set(), storyRoots, storyRoots)).toBe(blocking);
    });

    it("keeps a story root's sidecar finding a warning when the read model lists no story root (#987)", () => {
        const found = issue("rels-broken", { path: "word/_rels/document.xml.rels", message: "Broken reference to media/image1.png/." });
        registerRelsBrokenTarget(found, "word/media/image1.png/");
        expect(isWordBlockingIssue(found)).toBe(false);
    });

    it("leaves an unregistered rels-broken a warning, whatever its message (#734)", () => {
        expect(isWordBlockingIssue(issue("rels-broken", { message: "Broken reference to ../customXml/item1.xml" }))).toBe(false);
    });

    it.each([
        ["missing-attribute", true],
        ["body-content", false],
    ] as const)("judges rels-empty-element %s by its registered shape (#735)", (shape, blocking) => {
        // The body-content message carries the poisoned Id: the shape, not the
        // text, decides.
        const found = issue("rels-empty-element", {
            message: '<Relationship Id="missing required attribute"> must be a self-closing empty element',
        });
        registerRelsEmptyElementShape(found, shape);
        expect(isWordBlockingIssue(found)).toBe(blocking);
    });

    it("leaves an unregistered rels-empty-element a warning, whatever its message (#735)", () => {
        expect(isWordBlockingIssue(issue("rels-empty-element", { message: "<Relationship> missing required attribute(s): Id" }))).toBe(
            false,
        );
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
            ["word/charts/chart1.xml", reach(["word/charts/chart1.xml"]), true],
            ["word/charts/chart1.xml", noReach, false],
            ["word/header1.xml", noReach, false],
            ["word/theme1.xml", noReach, false],
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

        it("warns on a relationships sidecar the run's reachability set does not name", () => {
            // A sidecar is never a relationship target, so the predicate
            // itself adds no rule for it; the run's `wordReadParts` names the
            // main document's own sidecar (#514), and no other (#437). A
            // sidecar's structural defects keep their own blocking code
            // (`opc-relationships-structure`, #433).
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "word/_rels/document.xml.rels" }), new Set(), noReach)).toBe(false);
            expect(
                isWordBlockingIssue(
                    issue("xml-syntax", { path: "word/_rels/document.xml.rels" }),
                    new Set(),
                    reach(["word/_rels/document.xml.rels"]),
                ),
            ).toBe(true);
            expect(isWordBlockingIssue(issue("xml-syntax", { path: "customXml/_rels/item1.xml.rels" }), new Set(), noReach)).toBe(false);
        });
    });
});

describe("hasWordRelationshipReference", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const AS = "http://purl.oclc.org/ooxml/drawingml/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    // Representative declarations exercise inherited WML references,
    // DrawingML attribute groups and VML, including cross-conformance ids.
    const references = [
        [W, "footerReference", "id"],
        [W, "subDoc", "id"],
        [W, "top", "id"],
        [WS, "top", "id"],
        [WS, "subDoc", "id"],
        ["http://schemas.openxmlformats.org/drawingml/2006/chart", "chart", "id"],
        ["http://purl.oclc.org/ooxml/drawingml/chart", "chart", "id"],
        ["urn:foreign", "footerReference", "id"],
        [A, "unknown", "embed"],
        ["http://schemas.microsoft.com/office/drawing/2016/SVG/main", "svgBlip", "embed"],
        [WS, "headerReference", "id"],
        [W, "objectLink", "id"],
        [W, "embedBold", "id"],
        [W, "hyperlink", "id"],
        [A, "blip", "embed"],
        [AS, "blip", "link"],
        [A, "hlinkClick", "id"],
        [A, "wavAudioFile", "embed"],
        [A, "audioFile", "link"],
        ["urn:schemas-microsoft-com:vml", "imagedata", "id"],
        ["urn:schemas-microsoft-com:office:office", "OLEObject", "id"],
    ];
    it.each(references)("counts a declared or unknown %s %s %s with either relationship namespace", (ns, name, attribute) => {
        for (const relationship of [R, RS]) {
            const document = parseXml(`<root><e:${name} xmlns:e="${ns}" xmlns:alias="${relationship}" alias:${attribute}="rId9"/></root>`);
            expect(hasWordRelationshipReference(document)).toBe(true);
        }
    });
    it.each([
        [W, "name", "id", "rId9"],
        [W, "name", "embed", "rId9"],
        [W, "name", "link", "rId9"],
        [WS, "name", "id", "rId9"],
        [W, "footerReference", "embed", "rId9"],
        [A, "blip", "id", "rId9"],
        [W, "footerReference", "id", ""],
    ])("ignores a non-reference %s %s %s=%s", (ns, name, attribute, value) => {
        const document = parseXml(`<e:${name} xmlns:e="${ns}" xmlns:r="${R}" r:${attribute}="${value}"/>`);
        expect(hasWordRelationshipReference(document)).toBe(false);
    });
    it("counts the root and later descendants but never an id in a foreign namespace", () => {
        expect(hasWordRelationshipReference(parseXml(`<w:footerReference xmlns:w="${W}" xmlns:r="${RS}" r:id="rId9"/>`))).toBe(true);
        expect(hasWordRelationshipReference(parseXml(`<w:footerReference xmlns:w="${W}" xmlns:x="urn:foreign" x:id="rId9"/>`))).toBe(false);
        expect(
            hasWordRelationshipReference(
                parseXml(`<root xmlns:w="${W}" xmlns:r="${R}"><w:name r:id="bad"/><w:hyperlink r:id="good"/></root>`),
            ),
        ).toBe(true);
    });
});

describe("unobservedReferenceAttributeValues (#647, #702)", () => {
    // The matcher for the reference-loaded types whose referenced shapes
    // Word was never shown but a picture's a:blip r:embed, which session 45
    // opened clean (#702 S1): every attribute in either relationships
    // namespace, and VML's o:relid, on any element but that one; nothing
    // else, and no empty value.
    it("collects every relationships-namespace attribute and o:relid, and nothing else", () => {
        const document = parseXml(
            `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
                `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
                `xmlns:rs="http://purl.oclc.org/ooxml/officeDocument/relationships" ` +
                `xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
                `<w:body><dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="" rs:cs="rId3"/><w:p o:relid="rId4" o:spid="rId5" w:rsidR="rId6" id="rId7"/></w:body></w:document>`,
        );
        expect([...unobservedReferenceAttributeValues(document)].sort()).toEqual(["rId1", "rId2", "rId3", "rId4"]);
    });

    it("leaves out a:blip r:embed, Transitional and ISO Strict, and keeps every other reference", () => {
        const document = parseXml(
            `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
                `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
                `xmlns:rs="http://purl.oclc.org/ooxml/officeDocument/relationships" ` +
                `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:as="http://purl.oclc.org/ooxml/drawingml/main" ` +
                `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:x" xmlns:o="urn:schemas-microsoft-com:office:office">` +
                `<w:body><a:blip r:embed="rId1" r:link="rId2"/><as:blip rs:embed="rId3"/><x:blip r:embed="rId4"/>` +
                `<c:chart r:id="rId5"/><w:p o:relid="rId6"/></w:body></w:document>`,
        );
        expect([...unobservedReferenceAttributeValues(document)].sort()).toEqual(["rId2", "rId4", "rId5", "rId6"]);
    });
});
