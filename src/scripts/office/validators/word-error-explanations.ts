/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 */

/**
 * Plain-language explanations for the issue codes that the `word-valid` profile
 * treats as Word-blocking (see `isWordBlockingIssue` in `validators/docx.ts`).
 *
 * Each entry answers two questions for a human reader: *why Microsoft Word
 * refuses to open the document* and *how to fix it*. These are surfaced by the
 * CLI under each error and are grounded in real-Word probing
 * (`broken-word/FINDINGS.md`, `validation-results/word-probe-all.jsonl`), not the
 * ISO schema — Word is stricter than the OpenXML SDK on some of these and more
 * lenient on others, so the wording reflects Word's actual behaviour.
 */
export const WORD_ERROR_EXPLANATIONS: Readonly<Record<string, string>> = Object.freeze({
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
    "word-content-type-invalid":
        "A part's content type in [Content_Types].xml is missing or wrong, so Word cannot load that part. Add the correct <Override>/<Default> content-type entry.",
    "ct-uncovered-part":
        "A package part has no content type in [Content_Types].xml. Word offers to recover the file. Remove the part, or add a Default or Override that covers it.",
    "word-drawing-scalar-whitespace":
        "A DrawingML numeric value (e.g. <wp:posOffset>, <wp14:pctWidth>) contains whitespace, which Word rejects as an invalid number. Trim the value so it is digits only.",
    "word-text-contains-paragraph":
        "A w:t text run contains a w:p paragraph. Word cannot open the file. Move the paragraph out of the text run.",
    "word-hyperlink-relationship":
        "A w:hyperlink r:id does not reference a Relationship whose Type ends in /hyperlink and whose TargetMode is External. Word offers to recover the file. Add that relationship, or remove the r:id.",
    "word-header-footer-relationship":
        "A w:headerReference or w:footerReference names a Relationship whose Type is neither the Transitional nor the ISO Strict URI the reference needs. Word offers to recover the file. Point the r:id at a Relationship of the matching type, or retype that Relationship.",
    "word-duplicate-theme-relationship":
        "A relationships part lists more than one theme. Word offers to recover the file. Keep a single theme relationship.",
    "word-duplicate-font-table-relationship":
        "A relationships part lists more than one font table. Word offers to recover the file. Keep a single font table relationship.",
    "word-duplicate-web-settings-relationship":
        "A relationships part lists more than one web settings part. Word offers to recover the file. Keep a single web settings relationship.",
    "word-duplicate-settings-relationship":
        "A relationships part lists more than one settings part. Word offers to recover the file. Keep a single settings relationship.",
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
    "word-attached-toolbar-utf8":
        "An attached toolbar referenced by wne:toolbarData contains the UTF-8 replacement sequence EF BF BD. Word offers to recover the file. Replace each sequence with the original byte. The same bytes in a part that toolbarData does not reference are not this defect.",
    "word-row-without-cell": "A table row contains no w:tc cell. Word cannot open the file. Remove the row or add a cell.",
    "word-cell-ends-with-table":
        "A table cell ends with w:tbl and has no following w:p. Word offers to recover the file. Add a w:p after that table.",
    "word-numbering-picture-bullet":
        "A numbering level's w:lvlPicBulletId does not match a w:numPicBullet in that numbering part. Word offers to recover the file. Remove the w:lvlPicBulletId, or add the picture bullet it names.",
    "word-footnote-reference-missing":
        "A w:footnoteReference id does not match a w:footnote in the footnotes part. Word offers to recover the file. Remove the reference, or add the footnote it names.",
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
    "rels-missing-sidecar":
        "A part uses relationship IDs but its _rels sidecar is missing, so Word cannot resolve the targets. Add the matching .rels file.",
    "rels-broken":
        "A relationship targets a part that does not exist in the package. Add the missing target part or remove the dangling relationship. (Note: Word tolerates a missing media/ image — it shows a placeholder — but not a missing customXml or required structural part.)",
    "rels-empty-element":
        "A <Relationship> entry is missing a required attribute (Id, Type, or Target), so the package is malformed. Supply the missing attribute.",
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
