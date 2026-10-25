/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { canonicalPartName } from "../src/lib/content-types";
import type { ValidationIssue } from "../src/lib/types";
import { parseXml } from "../src/lib/xml-helpers";
import { hasWordRelationshipReference, isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";

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
        "word-row-without-cell",
        "word-cell-ends-with-table",
        "word-textbox-deleted-text",
        "word-move-from-deleted-text",
        "word-choice-requires-undeclared",
        "word-text-contains-paragraph",
        "word-hyperlink-relationship",
        // Word 16.115 offers recovery for every story-reference type mismatch probed (#606, session 26).
        "word-header-footer-relationship",
        "word-duplicate-theme-relationship",
        "word-duplicate-font-table-relationship",
        "word-duplicate-web-settings-relationship",
        "word-duplicate-settings-relationship",
        "word-numbering-picture-bullet",
        "word-footnote-reference-missing",
        "documentbuilder-insert-unresolved",
        "docprops-boolean-invalid",
        // Word 16.115 offers recovery for every duplicate spelling probed (#351).
        "opc-equivalent-part-names",
    ])("keeps %s an error wherever it is found", (code) => {
        expect(isWordBlockingIssue(issue(code))).toBe(true);
    });

    it.each(["word/document.xml", "WORD/footnotes.XML"])(
        "keeps the markup-compatibility errors of %s an error under word-valid",
        (path) => {
            for (const code of ["mce-malformed", "mce-must-understand"]) {
                expect(isWordBlockingIssue(issue(code, { path })), `${code} ${path}`).toBe(true);
            }
        },
    );

    it("blocks mc:MustUnderstand on the package relationships part only (#636, s34c)", () => {
        // Word 16.115 offered recovery for mc:MustUnderstand of an unknown
        // namespace on the _rels/.rels root, then failed to open the file
        // (session 34, s34c, docs/plans/evidence/word-probes-2026-10-24.md).
        for (const path of ["_rels/.rels", "_RELS/.RELS", "_rels\\.rels", "_rels/.%72els"]) {
            expect.soft(isWordBlockingIssue(issue("mce-must-understand", { path })), path).toBe(true);
        }
        // Unobserved: every other relationships part only warns under word-valid.
        for (const path of [
            undefined,
            "word/_rels/document.xml.rels",
            "word/_rels/header1.xml.rels",
            "word/glossary/_rels/document.xml.rels",
            "docProps/_rels/.rels",
        ]) {
            expect.soft(isWordBlockingIssue(issue("mce-must-understand", { path })), String(path)).toBe(false);
        }
    });

    it("blocks mce-malformed in the main document's relationships part only (#636, s34d)", () => {
        // Word 16.115 offered recovery for an AlternateContent whose Choice
        // Requires prefix is undeclared in word/_rels/document.xml.rels
        // (session 34, s34d). Other mce-malformed shapes in that part share
        // the code and are unobserved; every other relationships part is
        // unobserved too and only warns under word-valid.
        for (const path of ["word/_rels/document.xml.rels", "WORD/_RELS/document.xml.RELS", "word/_rels\\document.xml.rels"]) {
            expect.soft(isWordBlockingIssue(issue("mce-malformed", { path })), path).toBe(true);
        }
        for (const path of [
            undefined,
            "_rels/.rels",
            "word/_rels/header1.xml.rels",
            "word/glossary/_rels/document.xml.rels",
            "docProps/_rels/.rels",
        ]) {
            expect.soft(isWordBlockingIssue(issue("mce-malformed", { path })), String(path)).toBe(false);
        }
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
                isWordBlockingIssue(issue("rels-missing-sidecar", { path: specimen.path }), new Set(), new Set(), new Set(), {
                    hasEffectiveRelationshipReference: specimen.reference,
                    reachedByUnobservedRelationship: specimen.unobserved,
                }),
            ).toBe(specimen.blocking);
        },
    );

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
