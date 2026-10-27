/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 */

/**
 * Plain-language explanations for the issue codes that the `word-valid` profile
 * treats as Word-blocking, plus advisory compatibility findings (see
 * `isWordBlockingIssue` in `validators/word-blocking.ts`).
 *
 * Each entry explains why Microsoft Word refuses or accepts the markup and
 * how to fix the defect. These are surfaced by the
 * CLI under each error and are grounded in real-Word probing
 * (`broken-word/FINDINGS.md`, `validation-results/word-probe-all.jsonl`), not the
 * ISO schema — Word is stricter than the OpenXML SDK on some of these and more
 * lenient on others, so the wording reflects Word's actual behaviour.
 */
export const WORD_ERROR_EXPLANATIONS: Readonly<Record<string, string>> = Object.freeze({
    "mce-malformed":
        "Markup compatibility has an invalid structure. AlternateContent needs Choices with declared, non-empty Requires prefixes followed by at most one Fallback; unwrapping a part root must leave exactly one element. If root unwrapping fails, validation keeps the stored part and its root findings. Word 16.115 offered recovery for a _rels/.rels root that unwraps to two Relationship elements (session 36) and for one that unwraps to none, its Relationships element in an unselected Choice (session 38); it opened clean roots that unwrap to the Relationships element, which word-valid accepts (#646). Other zero- and multiple-element root shapes are unobserved. In relationships parts, Word offered recovery for malformed compatibility in _rels/.rels and the main document or reached glossary sidecar, including a renamed main document; a stray, styles, footer, settings, numbering or header sidecar opened clean (sessions 36 and 37). Other sidecars are unobserved and warn under word-valid. Restore a valid stored root and compatibility structure.",
    "mce-must-understand":
        "A required namespace is not understood by the selected compatibility view. Word 16.115 offered recovery in _rels/.rels and the main document or reached glossary sidecar, including a renamed main document (sessions 34 and 36); a stray, styles or footer sidecar opened clean (sessions 36 and 37). Other sidecars are unobserved and warn under word-valid. XML parts under word/ keep their blocking classification. Remove the unsupported requirement or use a namespace the selected view understands.",
    "ignorable-undeclared":
        "Word refuses to open the file: mc:Ignorable lists a namespace prefix that is never declared on the element. Declare the prefix (xmlns:…) on the root, or remove it from mc:Ignorable.",
    "word-math-spre-body":
        "Word cannot open a document with a body-level <m:oMathPara> that contains <m:sPre>. Wrap the display math in a paragraph (<w:p>) instead of placing it directly in <w:body>.",
    "word-math-parse":
        "word/document.xml is not well-formed XML, so Word reports unreadable content. Fix the XML syntax error in the document body.",
    "ct-missing":
        "The package has no [Content_Types].xml stream, so Word cannot type any part and offers to recover the file. Restore the content-types stream with a Default or Override covering every part.",
    "word-main-document-part":
        "The package needs exactly one internal officeDocument relationship in _rels/.rels targeting an existing main document part. Missing or multiple main relationships make Word offer recovery. Add the missing relationship and its part, or restore a missing target. Remove extra internal officeDocument relationships, even if their targets are identical.",
    "opc-relationships-target-mode":
        "A relationship in the package or main document relationships part has a TargetMode other than exactly Internal or External. Word 16.115 offered recovery for external on a chart or diagramData relationship, and external, External with a trailing space or an empty value on an unreferenced hyperlink in the main sidecar; external in the package relationships part was refused too (session 55). The header-sidecar hyperlink opened clean; other relationships parts are unshown and keep their verdict. Use the exact Internal or External spelling, or omit TargetMode for an internal relationship.",
    "opc-relationships-markup-compatibility":
        "Word applies markup compatibility to relationships parts and opened each tested shape of this markup clean (a Fallback over a Choice it does not understand, an understood w14 Choice, and an Ignorable foreign child, in the main document's and the package's relationships); other shapes are untested. For strict OPC conformance, store those Relationship children directly under Relationships and remove the compatibility markup.",
    "word-content-type-invalid":
        "A part's content type in [Content_Types].xml is missing or wrong, so Word cannot load that part. Add the correct <Override>/<Default> content-type entry.",
    "word-part-content-type-unrecognized":
        "A part Word reads has a content type that is neither its own Word type, nor generic XML, nor another Word part's type, so Word cannot load that part. Word 16.115 offered recovery for the main document and the parts its own relationships relate by styles, settings, fontTable and numbering: typed application/octet-stream (styles also by a Default; sessions 46, 49 and 50 — session 50 showed that spelling on settings, fontTable and numbering), and typed APPLICATION/OCTET-STREAM, application/octet-stream; charset=utf-8, image/png or application/foo+xml on the main document and styles (sessions 49 and 50) and on settings, fontTable and numbering (session 53); other reaches are unobserved. Give the part the content type its relationship expects.",
    "ct-uncovered-part":
        "A package part has no content type in [Content_Types].xml. Word offers to recover the file. Remove the part, or add a Default or Override that covers it.",
    "word-drawing-scalar-whitespace":
        "A DrawingML numeric value (e.g. <wp:posOffset>, <wp14:pctWidth>) contains whitespace, which Word rejects as an invalid number. Trim the value so it is digits only.",
    "word-text-contains-paragraph":
        "A w:t text run contains a w:p paragraph. Word cannot open the file. Move the paragraph out of the text run.",
    "word-hyperlink-relationship":
        "A w:hyperlink r:id does not reference a Relationship whose Type is the Transitional or the ISO Strict hyperlink URI with TargetMode External. Word offers to recover the file. Add that relationship, or remove the r:id.",
    "word-header-footer-relationship":
        "A w:headerReference or w:footerReference has an empty or missing id in both relationships namespaces, names no Relationship with its r:id, or names a Relationship whose Type is neither the Transitional nor the ISO Strict URI the reference needs. Word cannot open the file when the id is missing; an id in a third namespace counts as missing. Word offers to recover the file for an empty id, a dangling id or a wrong Type. Give the reference a non-empty r:id naming a Relationship of the matching Type: add that Relationship when none exists, or retype the one it names.",
    "word-part-reference-relationship":
        "A c:chart, a dgm:relIds r:dm or an o:OLEObject in a part Word reads names a relationship whose Type is not the one the element reads, an External relationship where an internal one is required, an Id no relationship has — whether the part's own relationships part lacks the Id or the part has no relationships part at all, both refused (session 55 s55a) — or an empty r:id or r:dm — Word 16.115 offered recovery for an image, diagramData, package, chart, External hyperlink and External chart relationship alike and refused every near-miss Type spelling (a trailing slash, a capital letter), whatever the target part, while the matching internal relationship opened clean (sessions 45, 46 and 52); an External package relationship an o:OLEObject names opened clean too (session 52 B1). A c:chart carrying an id attribute in both relationships namespaces (r:id and the ISO Strict rs:id, whatever their values) is an open error with no recovery offer (session 55, s55b and s55d); rs:id alone opens clean and draws the chart, and the same pair on the other references is unobserved. Word offers to recover the file; for an empty dgm:relIds r:dm it reported an open error on one specimen and offered recovery on another. Point the reference at a Relationship of the Type it reads — chart for c:chart, an internal diagramData for dgm:relIds r:dm, package or oleObject for o:OLEObject — adding one with that Id and Type when the part has none, and set a non-empty id. Word tolerates dgm:relIds r:lo, r:qs and r:cs naming another type, so those stay unchecked.",
    "word-duplicate-theme-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one theme. Word offers to recover the file — a reached glossary's own sidecar with two theme relationships was refused (session 54) — while the same pair in _rels/.rels or a referenced header's relationships part opens clean. Only the exact theme Type counts: a …/relationships/custom/theme relationship beside the real one opens clean, while two real theme relationships get the recovery offer (session 52). Keep a single theme relationship.",
    "word-duplicate-font-table-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one font table. Word offers to recover the file; the same pair in _rels/.rels (session 54) or a referenced header's relationships part opens clean. Keep a single font table relationship.",
    "word-duplicate-web-settings-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one web settings part. Word offers to recover the file — a reached glossary's own sidecar with two webSettings relationships was refused and the same pair in _rels/.rels opened clean (session 54) — while the pair in a referenced header's relationships part opens clean. Keep a single web settings relationship.",
    "word-duplicate-settings-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one settings part. Word offers to recover the file; the same pair in _rels/.rels or a referenced header's relationships part opens clean. Keep a single settings relationship.",
    "word-duplicate-styles-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one styles part. Word offers to recover the file; the same pair in _rels/.rels or a referenced header's relationships part opens clean. Keep a single styles relationship.",
    "word-duplicate-numbering-relationship":
        "The main document's or a reached glossary document's relationships part lists more than one numbering part. Word offers to recover the file; the same pair in _rels/.rels or a referenced header's relationships part opens clean. Keep a single numbering relationship.",
    "word-duplicate-endnotes-relationship":
        "The main document's relationships part lists more than one endnotes part while the document references an endnote. Word offers to recover the file; the same pair with no endnote reference opens clean (session 62), and a reached glossary's own pair is unprobed. Keep a single endnotes relationship.",
    "word-duplicate-footnotes-relationship":
        "The main document's relationships part lists more than one footnotes part while the document references a footnote. Word offers to recover the file (session 63, s63d); the unreferenced pair opened clean (session 45 s45g), and a reached glossary's own pair is unprobed. Keep a single footnotes relationship.",
    "word-drawing-arc-radius":
        "An a:arcTo radius (wR or hR) is not an integer. Word cannot open the file. Replace the radius with an integer.",
    "word-empty-drawing":
        "A w:drawing element has no child element. Word does not open the file cleanly. Remove the drawing, or give it a wp:anchor or wp:inline child.",
    "word-drawing-extent-negative":
        "A wp:extent cx or cy is a negative integer. Word cannot open the file. Use a positive extent. Zero and positive extents are not this defect.",
    "word-anchor-missing-attribute":
        "A wp:anchor is missing locked, layoutInCell, or allowOverlap. Word cannot open the file. Add the missing attribute.",
    "word-position-offset-and-align":
        "A wp:positionH or wp:positionV contains both wp:posOffset and wp:align. Word cannot open the file. Keep only one of those children.",
    "word-sdt-id-empty": 'A w:sdt has w:id w:val="". Word cannot open the file. Set w:val to a number.',
    "word-attached-toolbar-relationship":
        "A wne:toolbarData r:id names a relationship whose Type is not exactly http://schemas.microsoft.com/office/2006/relationships/attachedToolbars, or no relationship at all. Word 16.115 offers to recover the file whatever the target's bytes hold — clean toolbar bytes are refused once the Type changes (session 48). Point the r:id at a Relationship of that exact Type, or retype the one it names.",
    "word-attached-toolbar-utf8":
        "An attached toolbar referenced by wne:toolbarData contains the UTF-8 replacement sequence EF BF BD. Word offers to recover the file. Replace each sequence with the original byte. The same bytes in a part that toolbarData does not reference are not this defect.",
    "word-row-without-cell": "A table row contains no w:tc cell. Word cannot open the file. Remove the row or add a cell.",
    "word-cell-ends-with-table":
        "A table cell ends with w:tbl and has no following w:p. Word offers to recover the file. Add a w:p after that table.",
    "word-numbering-picture-bullet":
        "A numbering level's w:lvlPicBulletId does not match a w:numPicBullet in that numbering part. Word offers to recover the file. Remove the w:lvlPicBulletId, or add the picture bullet it names.",
    "word-footnote-reference-missing":
        "A w:footnoteReference id does not match a w:footnote in the footnotes part. Word offers to recover the file. Remove the reference, or add the footnote it names.",
    "word-endnote-reference-missing":
        "A w:endnoteReference id does not match a w:endnote in the endnotes part. Word offers to recover the file (session 58, s58a; #784). Remove the reference, or add the endnote it names.",
    "word-shared-story-part":
        "A comments, footnotes or endnotes part is read through a relationship of its own segment of both the main document and a reached glossary document. Word offers to recover the file (session 66, s66a, for comments; session 67, s67a, for footnotes and endnotes); the same part used by one root alone, or related by both roots but used by one or by neither, opens clean. Shared header and footer parts, a part two glossaries share, and a part one root reads under one of these segments and the other under another are unprobed. Give each story root its own part.",
    "word-textbox-deleted-text":
        "A text box contains w:delText or w:delInstrText that is not inside a w:del in that text box. Word offers to recover the file. Remove that deleted text from the text box. Deleted text inside an mc:Fallback text box, or inside a w:del in the text box, is not this defect. A w:del outside the text box does not cover it.",
    "word-move-from-deleted-text":
        "A w:moveFrom element contains w:delText. Word offers to recover the file. Use w:t for the moved text. Deleted text outside a move source is not this defect.",
    "word-choice-requires-undeclared":
        "mc:Choice Requires names a namespace prefix that is not declared on that element or an ancestor. Word offers to recover the file. Declare the prefix there. A declaration on a descendant, or only on the fallback, does not count.",
    "id-durable-overflow": "A w16cid durableId exceeds the value range Word accepts. Regenerate the durableId within range.",
    "comment-thread-commentid-paraid-orphan":
        "The comment threading metadata is inconsistent: a commentsIds paraId does not resolve to a real comment. Word refuses the file. Rebuild word/commentsIds.xml / commentsExtended.xml so every reference resolves.",
    "comment-thread-commentid-missing-paraid":
        "A comment is missing the paraId that the threading parts reference, so Word cannot reconcile the thread. Add the matching w14:paraId.",
    "comment-thread-commentid-missing-durableid":
        "A comment is missing the durableId that the threading parts reference. Add the matching w16cid:durableId.",
    "comment-thread-commentid-duplicate-paraid":
        "Two comments share the same paraId, so Word cannot tell them apart and refuses to open. Make every paraId unique.",
    "comment-thread-commentid-duplicate-durableid": "Two comments share the same durableId. Make every durableId unique.",
    "comment-thread-durableid-orphan":
        "A durableId in the comment-threading metadata points to no existing comment. Remove the orphan reference or add the comment.",
    "comment-thread-durableid-missing":
        "Comment-threading metadata expected a durableId that is not present. Add it or drop the threading entry.",
    "comment-thread-durableid-duplicate": "Duplicate durableId in comment-threading metadata. Make each unique.",
    "rels-id-duplicate":
        "Duplicate relationship Ids in _rels/.rels or a story root's relationships — the main document's or a reached glossary document's — make Word offer recovery, even when identical or unreferenced. Ids are counted after markup compatibility, including selected Fallback relationships. Keep each Id unique within that relationships part and update any references when renaming it. Duplicates in other relationships parts remain non-blocking under word-valid.",
    "rels-missing-sidecar":
        "A part uses relationship IDs but its _rels sidecar is missing. Word-valid blocks only a Word-read part, or one that a known Office relationship of an unobserved type names, with a counted reference in its processed markup-compatibility view. A part every reaching relationship names through the chart type blocks only when a Word-read part's effective XML references one of those relationship Ids — Word opens a related-but-unreferenced chart clean. Generated bundled-schema declarations exclude only known non-references; undeclared elements and namespaces keep counting. A genuine header/footer basename outside the story roots remains tolerated. Other occurrences warn under word-valid. Add the matching .rels file when the reference is used.",
    "rels-broken":
        "A relationship targets a part that does not exist in the package. Add the missing target part or remove the dangling relationship. (Note: Word tolerates a missing media/ image — it shows a placeholder — but not a missing customXml or required structural part.)",
    "rels-empty-element":
        "A <Relationship> entry is missing a required attribute (Id, Type, or Target), or a complete one carries body content — this code reports both malformed shapes. The missing-attribute shape is the one Word refuses — annotations_import.docx reports unreadable content; supply the missing attribute. The body-content shape opens clean in Word (EmptyRelationshipElement.docx), so word-valid only warns on it; make the element self-closing.",
    "xml-syntax": "An XML part is not well-formed, so Word reports unreadable content. Fix the XML syntax error in the named part.",
    "xsd-error":
        "An element or attribute is not allowed in this context per the schema, and Word rejects it (not all schema violations block Word — this one does). Correct the offending element/attribute shown in the message.",
    "zip-bytes-after-eocd":
        "Bytes follow the ZIP end-of-central-directory record beyond the comment length that record declares. Word offers to recover the document. Delete those bytes. A ZIP comment whose length is declared in the record is not this defect. A second copy of the archive appended after the record is this defect even when that copy ends with its own end-of-central-directory record.",
    "zip-bytes-before-archive":
        "Bytes sit in front of the ZIP archive's first entry: junk, a self-extractor stub, or a whole prepended archive, whether the archive's offsets are relative or absolute. Word offers to recover the document. Strip them so the archive begins at the file's first byte. A ZIP comment declared inside the record is not this defect.",
    "zip-eocd-comment-overrun":
        "The ZIP end-of-central-directory comment length claims more bytes than remain in the file. Word offers to recover the document. Set that length to the number of bytes actually present after the record.",
});

/** Human explanation for an issue `code`, if one is known. */
export function explainWordError(code: string | undefined): string | undefined {
    return code ? WORD_ERROR_EXPLANATIONS[code] : undefined;
}
