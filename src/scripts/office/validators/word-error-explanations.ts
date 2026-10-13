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
    "word-content-type-invalid":
        "A part's content type in [Content_Types].xml is missing or wrong, so Word cannot load that part. Add the correct <Override>/<Default> content-type entry.",
    "word-drawing-scalar-whitespace":
        "A DrawingML numeric value (e.g. <wp:posOffset>, <wp14:pctWidth>) contains whitespace, which Word rejects as an invalid number. Trim the value so it is digits only.",
    "word-text-contains-paragraph":
        "A w:t text run contains a w:p paragraph. Word cannot open the file. Move the paragraph out of the text run.",
    "word-hyperlink-relationship":
        "A w:hyperlink r:id does not reference a Relationship whose Type ends in /hyperlink and whose TargetMode is External. Word offers to recover the file. Add that relationship, or remove the r:id.",
    "word-duplicate-theme-relationship":
        "A relationships part lists more than one theme. Word offers to recover the file. Keep a single theme relationship.",
    "word-drawing-arc-radius":
        "An a:arcTo radius (wR or hR) is not an integer. Word cannot open the file. Replace the radius with an integer.",
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
    "zip-eocd-comment-overrun":
        "The ZIP end-of-central-directory comment length claims more bytes than remain in the file. Word offers to recover the document. Set that length to the number of bytes actually present after the record.",
});

/** Human explanation for an issue `code`, if one is known. */
export function explainWordError(code: string | undefined): string | undefined {
    return code ? WORD_ERROR_EXPLANATIONS[code] : undefined;
}
