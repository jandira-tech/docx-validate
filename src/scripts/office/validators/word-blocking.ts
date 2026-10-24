/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Which errors stay errors under the `word-valid` profile: those that predict
 * Word's open-error or unreadable-content dialogs. Every other error is
 * reported as a warning. A full DOCX run and a partial run both apply this
 * one predicate, so they cannot drift (#381).
 */

import { canonicalPartName, isContentTypesStreamName, PART_NAME_FOLDER_SEGMENT_DOT } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";

/**
 * Whether a part lies under `word/`, comparing part names as OPC does: ASCII
 * case-insensitively, with escapes of unreserved characters decoded, so
 * `WORD/document.xml` names the part `word/document.xml` does (#392).
 */
const isWordFolderPart = (partName: string | undefined): boolean =>
    partName !== undefined && canonicalPartName(partName).startsWith("word/");

/**
 * Whether a part is the package relationships part, compared as OPC compares
 * part names. Word 16.115 offers recovery when `_rels/.rels` is zero bytes,
 * truncated or uses an undeclared prefix (#388; hash-bound observations in
 * docs/plans/evidence/word-probes-2026-10-24.md).
 */
const isPackageRelationshipsPart = (partName: string | undefined): boolean =>
    partName !== undefined && canonicalPartName(partName) === "_rels/.rels";

export const isWordBlockingIssue = (
    issue: ValidationIssue,
    propertiesParts: ReadonlySet<string> = new Set(),
    wordReadParts: ReadonlySet<string> = new Set(),
): boolean => {
    switch (issue.code) {
        // NB: run-content-misplaced is NOT Word-blocking — a real Word probe
        // (open -g, no focus steal) opens bare <w:tab>/<w:br>/<w:drawing> under
        // <w:p> CLEANLY (Word silently relocates/drops the run content). It
        // stays an error in lenient/strict (structural invalidity + silent
        // data loss) but the word-valid openability profile tolerates it.
        case "comment-thread-commentid-paraid-orphan":
        case "comment-thread-commentid-missing-paraid":
        case "comment-thread-commentid-missing-durableid":
        case "comment-thread-commentid-duplicate-paraid":
        case "comment-thread-commentid-duplicate-durableid":
        case "comment-thread-durableid-orphan":
        case "comment-thread-durableid-missing":
        case "comment-thread-durableid-duplicate":
        case "id-durable-overflow":
        case "word-math-parse":
        case "word-math-spre-body":
        case "word-content-type-invalid":
        // #317, Word 16.115: a foreign main root prevents opening; an ignored main root silently drops all content.
        case "word-document-root":
        // Word 16.115 offers recovery for a package without
        // [Content_Types].xml — it cannot type a single part (#435; session
        // 3c in docs/plans/evidence/word-probes-2026-10-24.md).
        case "ct-missing":
        // Word 16.115 offers recovery for an empty archive and for a package
        // whose main document part is gone (#380, session 3c, same file), and
        // for one whose officeDocument relationship is removed (session 4).
        // Multiple internal officeDocument relationships also get recovery,
        // even with the same target (#515, Word probes session 15, s15b/s15d).
        case "word-main-document-part":
        // A part with no Override and no Default makes Word offer recovery
        // ([trash]/0000.dat). A part that has a content type but no
        // relationship is tolerated, so rels-unreferenced stays a warning.
        case "ct-uncovered-part":
        // #431: Word 16.115 offers recovery for all five duplicate Default/
        // Override shapes, including identical, conflicting and case variants;
        // MINIMAL_PARTS opens clean (session 3, word-probes-2026-10-24.md).
        case "ct-duplicate-declaration":
        // #459: Word 16.115 offers recovery for every Markup Compatibility
        // usage in the content-types stream — mc:Ignorable, an ignored
        // element, or AlternateContent in either orientation (session 14a in
        // docs/plans/evidence/word-probes-2026-10-24.md).
        case "ct-markup-compatibility":
        // paragraph-sectpr-breaks.docx: empty vt:filetime causes Word unreadable-content recovery; empty lpwstr in complex2.docx opens clean.
        case "word-custom-property-filetime-empty":
        // header-with-rid-no-sidecar.docx: Word open-error with missing wp:extent; the complete scaffold in validators-docx.test.ts opens clean.
        case "word-drawing-extent-missing":
        // Historical WORD_BLOCKING_MISPLACED_LOCALS extent probe comment: misplaced extents are rejected; positive-extent drawing control opens clean.
        case "word-drawing-extent-misplaced":
        // header-with-rid-no-sidecar.docx: Word open-error without pic:nvPicPr; the complete header scaffold in validators-docx.test.ts is tolerated.
        case "word-picture-nvpicpr-missing":
        // header-with-rid-no-sidecar.docx: Word open-error without pic:spPr; the complete header scaffold in validators-docx.test.ts is tolerated.
        case "word-picture-sppr-missing":
        // word-rejects-misplaced-run.docx: a w:r directly inside w:tc makes Word refuse the file (OPEN_ERROR probe); the same runs under w:p/w:ins/w:del open clean.
        case "word-run-misplaced":
        // 5Errors.docx (word: unreadable-content-warning): a second w:rPr inside one run makes Word offer recovery; a single leading w:rPr opens clean.
        case "word-run-properties-misplaced":
        // invalid-underline.docx (word: unreadable-content-warning): w:u directly inside w:pPr makes Word offer recovery; w:u inside w:rPr opens clean.
        case "word-underline-misplaced":
        // 5Errors.docx (word: unreadable-content-warning): w:pPr nested inside w:pPr makes Word offer recovery; the duplicate-pPr-in-one-paragraph control (word-tolerated-duplicate-ppr.docx) opens clean.
        case "word-paragraph-properties-misplaced":
        // Corpus probe (word-invalid-fixtures/README.md, CLEAN-FP 0/1304): a w:sectPr where the schema forbids it is a Word rejection class; legal body/pPr placements (word-tolerated-misplaced-pgsz.docx) open clean.
        case "word-section-properties-misplaced":
        // table_tests_v2.docx (probed UNREADABLE_CONTENT): w:tblBorders ordered after w:tblLayout inside w:tblPr makes Word offer recovery; Word-opened jubarte k07/k08/k09 keep tcPr tolerated.
        case "word-table-properties-misplaced":
        // UnknownElement.docx (word: open-error): an unknown w:p2 block in the body makes Word refuse the file; unknown names in property containers (annotations_import_2.docx, clean-open) stay tolerated.
        case "word-unknown-element":
        // hyperlink-exported.docx (word: open-error): a namespace-less <link/> inside w:rPr makes Word refuse the file; the namespaced w:link style child opens clean (word-tolerated-misplaced-link.docx).
        case "word-link-without-namespace":
        case "word-drawing-scalar-whitespace":
        case "word-drawing-arc-radius":
        case "word-empty-drawing":
        case "word-drawing-extent-negative":
        case "word-anchor-missing-attribute":
        case "word-position-offset-and-align":
        case "word-sdt-id-empty":
        case "word-attached-toolbar-utf8":
        case "word-row-without-cell":
        case "word-cell-ends-with-table":
        case "word-textbox-deleted-text":
        case "word-move-from-deleted-text":
        case "word-choice-requires-undeclared":
        case "word-text-contains-paragraph":
        case "word-hyperlink-relationship":
        case "word-duplicate-theme-relationship":
        case "word-duplicate-font-table-relationship":
        case "word-duplicate-web-settings-relationship":
        case "word-duplicate-settings-relationship":
        case "word-numbering-picture-bullet":
        case "word-footnote-reference-missing":
        // Both root-caused against real Word (minimal-repair + re-probe) and
        // verified disjoint from the Word-clean corpus (0/350): an unresolved
        // DocumentBuilder <Insert> directive -> OPEN_ERROR; a docProps boolean
        // with whitespace -> unreadable-content recovery prompt.
        case "documentbuilder-insert-unresolved":
        case "docprops-boolean-invalid":
        // Word 16.115 offers recovery for every duplicate spelling probed, even
        // with identical bytes: case variants, an unreserved escape, a
        // backslash, and two content-types streams (#351; evidence in
        // docs/plans/evidence/word-probes-2026-10-24.md).
        case "opc-equivalent-part-names":
        // Word 16.115 offers recovery for every unreserved escape probed
        // (letters in either case, a digit, -, ., _ and ~), even on an
        // unreferenced item, and opens the %20 control clean (#408; evidence
        // in docs/plans/evidence/word-probes-2026-10-24.md).
        case "opc-part-name-escaped-unreserved":
        // Word 16.115 offers recovery for an unreferenced `word%2Fxtra.xml`,
        // refuses `word%5Cxtra.xml`, and opens the well-formed `word/xtra.xml`
        // control clean — the stored name alone predicts the dialog (#436;
        // evidence in docs/plans/evidence/word-probes-2026-10-24.md).
        // Word 16.115 recovers escaped RFC 2396 marks, high octets and literal
        // non-ASCII names, even unreferenced (#408, #443, #451; sessions 6–12
        // in docs/plans/evidence/word-probes-2026-10-24.md).
        case "opc-part-name-word-refused":
        case "ct-override-partname-word-refused":
        case "opc-part-name-escaped-slash":
            return true;
        // Word 16.115 offers recovery for every malformed Override PartName
        // probed, even when nothing is stored under that name and the content
        // type is well formed (#414; probe sessions 3c, 4, 5 and 7 in
        // docs/plans/evidence/word-probes-2026-10-24.md). A well-formed
        // PartName for a missing part opens clean (#401), so the PartName
        // alone decides it. A folder segment ending with a dot opens clean
        // too (#450, session 7); a dot-only one does not, and carries its own
        // rule (#453, session 12). The message ends with the rule broken, after
        // the quoted PartName, so only its end is read.
        case "ct-override-partname-invalid":
            return !issue.message.endsWith(`it ${PART_NAME_FOLDER_SEGMENT_DOT}.`);
        // Word 16.115 offers to recover a package whose Word-read part is
        // typed only by the generic application/xml Default (#432, session
        // 8a): the main document, styles, settings, fontTable and webSettings.
        case "word-part-content-type-generic":
            return true;
        // NB: comment-orphan-start/end and comment-marker-missing are NOT
        // promoted here. Empirically Word TOLERATES an orphan comment-range
        // marker (a commentRangeStart with no End, referencing a missing
        // comment) — verified: comments.unmatched-comment-marker_from_html.docx
        // carries the exact same orphan and opens cleanly in real Word. So an
        // orphan marker does not predict Word rejection; keep it a warning.
        case "rels-missing-sidecar":
            return !/^word\/(?:header|footer)\d*\.xml$/i.test(issue.path ?? "");
        case "opc-relationships-structure": {
            // #433: Word 16.115 offers recovery for a wrong root or a foreign
            // child in _rels/.rels (session 3c) and in word/_rels/document.xml.rels
            // (session 4c). Relationships parts outside word/ are unobserved and
            // only warn, as xml-syntax does there.
            const path = issue.path?.replace(/\\/g, "/");
            return isPackageRelationshipsPart(path) || isWordFolderPart(path);
        }
        case "xml-syntax":
        // An undeclared namespace prefix is the same class of defect as a
        // malformed part — Word reports unreadable content — so it follows
        // the identical reachability policy (#437).
        case "xml-undeclared-prefix":
        // A part whose text cannot be decoded is unreadable too: Word 16.115
        // refuses a word/document.xml declaring an unknown encoding and opens
        // the windows-1252 control clean (#397). It also refuses real labels
        // libxml2 cannot decode: macintosh, x-mac-cyrillic, IBM866, UTF-7
        // (#410). The other paths follow the xml-syntax policy below.
        case "xml-encoding":
            // #437: Word 16.115 reads the parts its relationships reach, not
            // every part under word/ — an unreferenced truncated word/xtra.xml
            // opens clean even with its own Override, a related-but-unused
            // malformed footer, header, comments, footnotes or endnotes part
            // opens clean (#516: an unused header or endnotes part under
            // word/ too — session 15), while an eagerly parsed part (styles,
            // settings, fontTable, webSettings, numbering, customXml, theme,
            // glossaryDocument — #516, session 15) and a used lazy one get
            // repair — wherever they are stored. Unobserved relationship
            // types under word/ (chart, image, …) keep blocking, as before.
            // The content-types stream, the package relationships part and the
            // properties parts (#409) always block; so does each main
            // document's own relationships part, which `wordReadParts`
            // reaches wherever the main document is stored (#514; Word offers
            // recovery for a zero-byte and a truncated document.xml.rels,
            // session 15). A malformed sidecar no rule covers now warns.
            return (
                (issue.path !== undefined && propertiesParts.has(canonicalPartName(issue.path))) ||
                (issue.path !== undefined && wordReadParts.has(canonicalPartName(issue.path))) ||
                isPackageRelationshipsPart(issue.path) ||
                (issue.path !== undefined && isContentTypesStreamName(issue.path))
            );
        case "rels-broken":
            // NB: a broken media/ rel (image Target absent) is NOT word-blocking:
            // real Word shows a missing-image placeholder and opens cleanly
            // (verified: word-tolerated-broken-media-rel.docx). Broken media alone
            // does not predict rejection — files unreadable "with" a broken media
            // ref fail for a separate, undetected reason. Stays an error in
            // lenient/strict; only the word-valid openability profile tolerates it.
            // Structural rels (customXml, nested .rels) remain word-blocking.
            return issue.message.includes("../customXml/") || /\/_rels\/|\.rels$/i.test(issue.message);
        case "rels-empty-element":
            return issue.message.includes("missing required attribute");
        case "mce-malformed":
        case "mce-must-understand":
            return true;
        default:
            return false;
    }
};
