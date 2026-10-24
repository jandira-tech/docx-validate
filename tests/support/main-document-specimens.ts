/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { MINIMAL_PARTS } from "./minimal-docx";

// #515: the shapes from session 15's s15builder.mts, built from MINIMAL_PARTS.
// Distinct target: am-f6-second-officedoc-wellformed-control (s15b).
// Same target: am-f6-duplicate-officedoc-same-target (s15d).
const OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

export const ORPHAN_MAIN_DOCUMENT_PARTS = {
    ...MINIMAL_PARTS,
    "alt/document.xml": MINIMAL_PARTS["word/document.xml"]!,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        '<Override PartName="/alt/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
};

export const DUPLICATE_MAIN_DOCUMENT_PARTS = [
    { label: "distinct target", target: "alt/document.xml", parts: ORPHAN_MAIN_DOCUMENT_PARTS },
    { label: "same target", target: "word/document.xml", parts: MINIMAL_PARTS },
].map(({ label, target, parts }) => ({
    label,
    parts: {
        ...parts,
        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${OFFICE_DOCUMENT}" Target="${target}"/></Relationships>`,
        ),
    },
}));
