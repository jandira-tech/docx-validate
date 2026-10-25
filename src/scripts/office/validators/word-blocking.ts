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
import { NS, type ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameAll } from "../../../lib/xml-helpers";
import {
    SCHEMA_DECLARED_ELEMENTS,
    SCHEMA_REFERENCE_ELEMENTS,
    SCHEMA_UNRESOLVED_ELEMENTS,
} from "../../../lib/relationship-reference-elements";

/** Per-issue facts supplied by the caller that owns the processed XML view. */
export type WordBlockingFacts = {
    hasEffectiveRelationshipReference?: boolean;
    reachedByUnobservedRelationship?: boolean;
};

/**
 * Count non-empty id/embed/link in either relationships namespace in the
 * caller's effective XML. Only a schema-declared element with no declaration
 * carrying that attribute can exclude it (w:name is CT_String). Generated
 * bundled-XSD sets union ambiguous declarations and both conformance classes;
 * undeclared elements/namespaces and unresolved types keep counting. Thus
 * unobserved extensions keep the pre-#638 blocking decision. No XSD is loaded
 * or compiled on this path. Session 35, chunks s35a/s35b, supplies the w:name
 * and unselected-branch evidence; it did not probe arbitrary extensions.
 */
export const hasWordRelationshipReference = (document: Document): boolean => {
    for (const element of getElementsByTagNameAll(document)) {
        const identity = `{${element.namespaceURI ?? ""}}${element.localName}`;
        const unknown = !SCHEMA_DECLARED_ELEMENTS.has(identity) || SCHEMA_UNRESOLVED_ELEMENTS.has(identity);
        for (const attribute of ["id", "embed", "link"]) {
            if (
                (unknown || SCHEMA_REFERENCE_ELEMENTS.get(identity)?.has(attribute)) &&
                (element.getAttributeNS(NS.R, attribute) ||
                    element.getAttributeNS("http://purl.oclc.org/ooxml/officeDocument/relationships", attribute))
            )
                return true;
        }
    }
    return false;
};

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

/**
 * Whether a part is one of the XML parts under `word/`, comparing part names
 * as OPC does: ASCII case-insensitively, with escapes of unreserved
 * characters decoded. Relationships parts under `word/` end in `.rels` and
 * are excluded: their markup-compatibility policy is observed per part
 * (#636), not inherited from the folder.
 */
const isWordXmlPart = (partName: string | undefined): boolean =>
    partName !== undefined && /^word\/.+\.xml$/.test(canonicalPartName(partName));

/**
 * Whether a part is the main document's relationships part at its default
 * spelling, compared as OPC compares part names. Word 16.115 offered
 * recovery for an AlternateContent whose mc:Choice Requires prefix is
 * undeclared there (session 34, s34d, #636). A main document stored under
 * another name keeps an unobserved warning, as every other sidecar does.
 */
const isDocumentRelationshipsPart = (partName: string | undefined): boolean =>
    partName !== undefined && canonicalPartName(partName) === "word/_rels/document.xml.rels";

/**
 * Whether `issue` stays an error under `word-valid`: `true` when Word 16.115
 * is observed to refuse or offer to recover a file carrying it. The three sets
 * hold canonical part names, computed once per run by the caller:
 * `propertiesParts` (the core, app and custom properties parts the package
 * relationships name, `propertiesParts()`), `wordReadParts` (the parts Word
 * reads at open, `wordReadParts()`) and `storyRootParts` (every main document
 * plus every reached glossary document, `wordStoryRootParts()`). An empty set
 * makes the findings that depend on it non-blocking.
 */
export const isWordBlockingIssue = (
    issue: ValidationIssue,
    propertiesParts: ReadonlySet<string> = new Set(),
    wordReadParts: ReadonlySet<string> = new Set(),
    storyRootParts: ReadonlySet<string> = new Set(),
    facts: WordBlockingFacts = {},
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
        // Word 16.115 offers recovery for a story sectPr's w:headerReference or
        // w:footerReference whose r:id names a relationship of another type —
        // header→footer, footer→header, a reached glossary's references, and an
        // unrelated type such as styles — while a matching reference and
        // relationship opens clean even when the target's root element is the
        // other story's (#606; session 26 in
        // docs/plans/evidence/word-probes-2026-10-24.md).
        case "word-header-footer-relationship":
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
        // Word 16.115 offers to recover the package when a Word-read part's
        // content type is not exactly the part's type ignoring ASCII case —
        // any MIME parameter or padding refuses, an upper-case spelling
        // opens clean (#540, #550; sessions 16–17, 2026-10-24, the table
        // living with issue #550).
        case "word-part-content-type-inexact":
        // Session 24: a read header/footer declared with the other story's type (#600).
        case "word-part-content-type-mismatch":
            return true;
        case "rels-id-duplicate": {
            // #615, #617: Word 16.115 offers recovery for duplicate Ids in
            // the package relationships part and in the sidecar of a story
            // root — the main document and every glossary a main-document
            // `glossaryDocument` relationship reaches — even unreferenced or
            // identical (sessions 29 and 30, s29a/s29b and s30a in
            // docs/plans/evidence/word-probes-2026-10-24.md). A header's own
            // sidecar opened CLEAN (session 29) and so did an orphan
            // `word/document.xml`'s (session 30): every other sidecar keeps
            // its strict error and word-valid warning, even if Word reads
            // their owner.
            const part = issue.path?.replace(/\\/g, "/");
            if (isPackageRelationshipsPart(part)) return true;
            const source = part === undefined ? undefined : canonicalPartName(part).match(/^(.*\/)?_rels\/([^/]+)\.rels$/);
            return source !== undefined && source !== null && storyRootParts.has(`${source[1] ?? ""}${source[2]}`);
        }
        // NB: comment-orphan-start/end and comment-marker-missing are NOT
        // promoted here. Empirically Word TOLERATES an orphan comment-range
        // marker (a commentRangeStart with no End, referencing a missing
        // comment) — verified: comments.unmatched-comment-marker_from_html.docx
        // carries the exact same orphan and opens cleanly in real Word. So an
        // orphan marker does not predict Word rejection; keep it a warning.
        case "rels-missing-sidecar": {
            // #638, session 35, chunks s35a/s35b: Word opens unread parts,
            // unselected MC branches and rs:id on w:name cleanly, but offers
            // recovery for rs:id on a footerReference inside sectPrChange.
            // The caller supplies the processed-reference fact; never infer
            // it from prose or stored bytes. Generated schema declarations
            // exclude only known non-references; undeclared elements and
            // namespaces keep blocking. The finding itself is unchanged.
            const part = issue.path?.replace(/\\/g, "/");
            if (part === undefined || !facts.hasEffectiveRelationshipReference) return false;
            const canonical = canonicalPartName(part);
            // #625, session 33 s33b: story roots retain precedence over the
            // header/footer basename tolerance, wherever the root is stored.
            if (storyRootParts.has(canonical)) return true;
            // A part an unobserved Office relationship type names keeps the
            // pre-#638 decision wherever it is stored: Word was never shown
            // one outside word/ (wordReadParts holds only those under it).
            const read = wordReadParts.has(canonical) || facts.reachedByUnobservedRelationship === true;
            return read && !/^word\/(?:header|footer)\d*\.xml$/i.test(issue.path ?? "");
        }
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
            // repair — wherever they are stored. A reached glossary document
            // is a second story root: its own relationships part is read
            // (#549, session 17), the lazy stories its references use are read
            // (#539, session 16; #549, session 17), and its styles, settings,
            // fontTable, webSettings, numbering and theme are eagerly parsed
            // (#562, session 19; the theme #569, session 20) — a glossary's
            // customXml and a nested glossaryDocument are observed, not read
            // (session 20: malformed, both open clean). Unobserved relationship
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
        case "mce-must-understand": {
            // Word 16.115 offered recovery for mc:MustUnderstand of a
            // namespace it does not understand on the _rels/.rels root and,
            // declined, reported an open error (session 34, s34c, #636); the
            // XML parts under word/ keep the blocking classification this
            // code had when only they were judged. Every other relationships
            // part is unobserved and only warns under word-valid.
            const path = issue.path?.replace(/\\/g, "/");
            return isWordXmlPart(path) || isPackageRelationshipsPart(path);
        }
        case "mce-malformed": {
            // Word 16.115 offered recovery for an mc:Choice whose Requires
            // prefix is undeclared in word/_rels/document.xml.rels (session
            // 34, s34d, #636); the XML parts under word/ keep their blocking
            // classification. Other mce-malformed shapes in that part share
            // the code and are unobserved; the package relationships part and
            // every other sidecar are unobserved too and only warn.
            const path = issue.path?.replace(/\\/g, "/");
            return isWordXmlPart(path) || isDocumentRelationshipsPart(path);
        }
        default:
            return false;
    }
};
