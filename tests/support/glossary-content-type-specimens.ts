/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Session 23's builder shapes, rebuilt from the committed evidence appendix (#595, #596). */
import { MINIMAL_PARTS } from "./minimal-docx";
import { GENERIC_NUMBERING_THEME_SPECIMENS } from "./generic-numbering-theme-specimens";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
export const GLOSSARY_THEME_PART = "word/glossary/theme/theme1.xml";

const glossary = (body: string) =>
    `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
    `<w:docPartBody>${body}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

const rel = (segment: string, target: string, id = "rIdE") => `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"/>`;
const rels = (body: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const override = (part: string, type: string) => `<Override PartName="/${part}" ContentType="${type}"/>`;
const viaGlossary = (xml: string, relationships: string, part: string, payload: string, type?: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml": xml,
    "word/glossary/_rels/document.xml.rels": rels(relationships),
    [part]: payload,
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `${rel("glossaryDocument", "glossary/document.xml", "rIdG")}</Relationships>`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `${override("word/glossary/document.xml", `${WML}.document.glossary+xml`)}${type === undefined ? "" : override(part, type)}</Types>`,
    ),
});

/** s23a: theme only the reached glossary relates; the optional footer tests #593's relationship ordering. */
export const glossaryThemeParts = (exact = false, order?: "theme-first" | "footer-first"): Record<string, string> => {
    const theme = GENERIC_NUMBERING_THEME_SPECIMENS[1];
    const themeRel = rel("theme", "theme/theme1.xml");
    const footerRel = rel("footer", "theme/theme1.xml", "rIdF");
    return viaGlossary(
        glossary("<w:p><w:r><w:t>Part</w:t></w:r></w:p>"),
        order === undefined ? themeRel : order === "theme-first" ? themeRel + footerRel : footerRel + themeRel,
        GLOSSARY_THEME_PART,
        theme.xml,
        exact ? theme.contentType : undefined,
    );
};
