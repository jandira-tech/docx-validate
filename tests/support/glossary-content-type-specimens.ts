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

export const GLOSSARY_STORY_SEGMENTS = ["footer", "footnotes", "endnotes", "comments"] as const;
export type GlossaryStorySegment = (typeof GLOSSARY_STORY_SEGMENTS)[number];
export const glossaryStoryPart = (segment: GlossaryStorySegment): string =>
    `word/glossary/${segment === "footer" ? "footer9" : segment}.xml`;

const FOOTER = `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`;
const FOOTNOTES =
    `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:id="1"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const ENDNOTES =
    `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
    `<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>` +
    `<w:endnote w:id="1"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`;
const COMMENTS = `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`;

/** s23a/s23b: exact story XML, without adding paraIds or repairing the comments control. */
export const glossaryStoryParts = (
    segment: GlossaryStorySegment,
    exact = false,
    referenced = true,
    order?: "story-first" | "header-first",
): Record<string, string> => {
    const part = glossaryStoryPart(segment);
    const file = part.slice("word/glossary/".length);
    const id = segment === "footer" ? "rIdF" : "rIdS";
    const references = {
        footer: "",
        footnotes: `<w:r><w:footnoteReference w:id="1"/></w:r>`,
        endnotes: `<w:r><w:endnoteReference w:id="1"/></w:r>`,
        comments: `<w:r><w:commentReference w:id="0"/></w:r>`,
    };
    const body =
        `<w:p><w:r><w:t>Part</w:t></w:r>${referenced ? references[segment] : ""}</w:p>` +
        `<w:sectPr>${referenced && segment === "footer" ? `<w:footerReference w:type="default" r:id="rIdF"/>` : ""}</w:sectPr>`;
    const storyRel = rel(segment, file, id);
    const headerRel = rel("header", file, "rIdH");
    return viaGlossary(
        glossary(body),
        order === undefined ? storyRel : order === "story-first" ? storyRel + headerRel : headerRel + storyRel,
        part,
        { footer: FOOTER, footnotes: FOOTNOTES, endnotes: ENDNOTES, comments: COMMENTS }[segment],
        `${WML}.${segment}+xml${exact ? "" : "; charset=utf-8"}`,
    );
};
