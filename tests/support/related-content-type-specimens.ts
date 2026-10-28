/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Sessions 49 and 50 (2026-10-26): the minimal package plus a settings,
 * fontTable or numbering part (session 47's empty roots), related from the
 * main document's own relationships part, its Override carrying another
 * type (#720, #722). Shapes mirror the sessions' builders; Word's outcomes
 * are recorded with the sessions, not here.
 */
import { MINIMAL_PARTS } from "./minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

/** The related parts the sessions typed: their segment, session-47 root and own content type. */
export const RELATED_PART_SPECIMENS = [
    {
        segment: "settings",
        root: `${XML_DECL}<w:settings xmlns:w="${W}"/>`,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    },
    {
        segment: "fontTable",
        root: `${XML_DECL}<w:fonts xmlns:w="${W}"/>`,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    },
    {
        segment: "numbering",
        root: `${XML_DECL}<w:numbering xmlns:w="${W}"/>`,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    },
] as const;

export type RelatedPartSpecimen = (typeof RELATED_PART_SPECIMENS)[number];

/** The minimal package with `specimen`'s part related as rId2, its Override carrying `type`. */
export const relatedContentTypeParts = (specimen: RelatedPartSpecimen, type: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [`word/${specimen.segment}.xml`]: specimen.root,
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="rId2" Type="${R}/${specimen.segment}" Target="${specimen.segment}.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/word/${specimen.segment}.xml" ContentType="${type}"/></Types>`,
    ),
});
