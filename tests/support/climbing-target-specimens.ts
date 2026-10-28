/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { MINIMAL_PARTS } from "./minimal-docx";

// #996: the shapes from Word session 81's s81builder.mts, built from
// MINIMAL_PARTS. Each names a part by a target whose `..` segments climb past
// the package root, which every read-model walk stops at the root since #986.
// Word read each such target at the root and refused the reached part for the
// reason it refuses the in-package twin, and opened every calibration clean
// (s81a/s81b, two runs each, all readings agree).

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const CT = "[Content_Types].xml";
const SIDECAR = "word/_rels/document.xml.rels";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const P = MINIMAL_PARTS as Record<string, string>;

const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};
const relationship = (id: string, segment: string, target: string) =>
    `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"/>`;
const appendRelationships = (rels: string, extra: string) => replaceOnce(rels, "</Relationships>", `${extra}</Relationships>`);
const overrides = (types: string, ...pairs: [string, string][]) =>
    pairs.reduce(
        (text, [part, type]) => replaceOnce(text, "</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`),
        types,
    );

/**
 * s81a: a footer holding Foot stored at /footer1.xml, named from word/ by
 * `../../footer1.xml` and referenced by the section, typed `type` by its
 * Override. Typed footer, Word opened it clean and rendered `Foot`; typed
 * styles, it offered recovery, as session 53 did for the in-package twin.
 */
export const climbingFooterTyped = (type: "footer" | "styles"): Record<string, string> => ({
    ...P,
    "word/document.xml": replaceOnce(
        replaceOnce(P["word/document.xml"]!, `<w:document xmlns:w="${W}"`, `<w:document xmlns:w="${W}" xmlns:r="${R}"`),
        "<w:sectPr>",
        '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/>',
    ),
    "footer1.xml": `${DECL}<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Foot</w:t></w:r></w:p></w:ftr>`,
    [SIDECAR]: appendRelationships(P[SIDECAR]!, relationship("rIdF", "footer", "../../footer1.xml")),
    [CT]: overrides(P[CT]!, ["footer1.xml", `${WML}.${type}+xml`]),
});

const CORE_OPEN =
    `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>S81</dc:title>`;

/**
 * s81a: docProps/core.xml named from `_rels/.rels` by `../docProps/core.xml`,
 * well-formed or with its root unclosed. Word opened the well-formed part
 * clean and offered recovery for the unclosed one (#409's reading).
 */
export const climbingCore = (wellFormed: boolean): Record<string, string> => ({
    ...P,
    "_rels/.rels": appendRelationships(
        P["_rels/.rels"]!,
        '<Relationship Id="rIdP" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="../docProps/core.xml"/>',
    ),
    "docProps/core.xml": wellFormed ? `${CORE_OPEN}</cp:coreProperties>` : CORE_OPEN,
    [CT]: overrides(P[CT]!, ["docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"]),
});

/**
 * s81a: the styles part moved to /styles.xml and named from word/ by
 * `../../styles.xml`, well-formed or with its root unclosed. Word opened the
 * well-formed part clean (session 75's reading) and offered recovery for the
 * unclosed one, an eagerly parsed part (session 15).
 */
export const climbingStyles = (wellFormed: boolean): Record<string, string> => {
    const parts: Record<string, string> = {
        ...P,
        "styles.xml": wellFormed ? P["word/styles.xml"]! : replaceOnce(P["word/styles.xml"]!, "</w:styles>", ""),
        [SIDECAR]: replaceOnce(P[SIDECAR]!, 'Target="styles.xml"', 'Target="../../styles.xml"'),
        [CT]: replaceOnce(P[CT]!, 'PartName="/word/styles.xml"', 'PartName="/styles.xml"'),
    };
    delete parts["word/styles.xml"];
    return parts;
};

const NOTE_REFERENCE = '<w:r><w:footnoteReference w:id="1"/></w:r>';

/**
 * s81a/s81b: word/footnotes.xml holding footnote 1, which the main document
 * relates by `footnotes.xml` and references, and a glossary
 * word/glossary/document.xml the main document relates by `glossaryTarget`,
 * whose sidecar relates the same notes part by `notesTarget`; the glossary's
 * docPart references footnote 1 when `glossaryUsesNote`. Word offered
 * recovery whenever both roots used the note, the glossary reached by
 * `../../word/glossary/document.xml` or its notes by
 * `../../../word/footnotes.xml`, as session 67 did in-package, and opened the
 * climbing glossary that uses no note clean, rendering the body's note.
 */
export const climbingSharedNotes = (glossaryTarget: string, notesTarget: string, glossaryUsesNote: boolean): Record<string, string> => ({
    ...P,
    "word/document.xml": replaceOnce(
        P["word/document.xml"]!,
        "<w:r><w:t>Hello</w:t></w:r>",
        `<w:r><w:t>Hello</w:t></w:r>${NOTE_REFERENCE}`,
    ),
    "word/footnotes.xml": `${DECL}<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    "word/glossary/document.xml":
        `${DECL}<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
        `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Text</w:t></w:r>${glossaryUsesNote ? NOTE_REFERENCE : ""}</w:p></w:docPartBody>` +
        `</w:docPart></w:docParts></w:glossaryDocument>`,
    "word/glossary/_rels/document.xml.rels": `${DECL}<Relationships xmlns="${PR}">${relationship("rId1", "footnotes", notesTarget)}</Relationships>`,
    [SIDECAR]: appendRelationships(
        P[SIDECAR]!,
        relationship("rIdN", "footnotes", "footnotes.xml") + relationship("rIdG", "glossaryDocument", glossaryTarget),
    ),
    [CT]: overrides(P[CT]!, ["word/footnotes.xml", `${WML}.footnotes+xml`], ["word/glossary/document.xml", `${WML}.document.glossary+xml`]),
});
