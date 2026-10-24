/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { MINIMAL_PARTS } from "./minimal-docx";
import { glossaryStoryParts } from "./glossary-content-type-specimens";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const storyContentType = (segment: string): string =>
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${segment}+xml`;

/** Sessions 24 and 26: the minimal package plus a main-document or reached-glossary header/footer. */
export const storyContentTypeParts = (
    segment: "header" | "footer",
    declared = storyContentType(segment === "header" ? "footer" : "header"),
    via: "main" | "glossary" = "main",
    referenced = true,
): Record<string, string> => {
    if (via === "glossary") {
        const parts = glossaryStoryParts("footer", true, referenced);
        if (segment === "header") {
            parts["word/glossary/header9.xml"] = parts["word/glossary/footer9.xml"]!.replaceAll("w:ftr", "w:hdr").replace(
                "Footer",
                "Header",
            );
            delete parts["word/glossary/footer9.xml"];
            parts["word/glossary/document.xml"] = parts["word/glossary/document.xml"]!.replace("w:footerReference", "w:headerReference");
            parts["word/glossary/_rels/document.xml.rels"] = parts["word/glossary/_rels/document.xml.rels"]!.replace(
                '/footer"',
                '/header"',
            ).replace("footer9.xml", "header9.xml");
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace("footer9.xml", "header9.xml");
        }
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(storyContentType("footer"), declared);
        return parts;
    }
    const file = `${segment}9.xml`;
    const root = segment === "header" ? "hdr" : "ftr";
    return {
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace(
            `<w:document xmlns:w="${W}"`,
            `<w:document xmlns:w="${W}" xmlns:r="${R}"`,
        ).replace("<w:sectPr>", `<w:sectPr>${referenced ? `<w:${segment}Reference w:type="default" r:id="rIdX"/>` : ""}`),
        [`word/${file}`]: `<w:${root} xmlns:w="${W}"><w:p><w:r><w:t>${segment === "header" ? "Header" : "Footer"}</w:t></w:r></w:p></w:${root}>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdX" Type="${R}/${segment}" Target="${file}"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/${file}" ContentType="${declared}"/></Types>`,
        ),
    };
};

export const MISMATCH_STORY_CASES = [
    { segment: "footer", via: "main", declared: storyContentType("header") },
    { segment: "header", via: "main", declared: storyContentType("footer") },
    { segment: "footer", via: "glossary", declared: storyContentType("header") },
    { segment: "footer", via: "glossary", declared: `${storyContentType("header")}; charset=utf-8` },
    { segment: "header", via: "glossary", declared: storyContentType("footer") },
] as const;

/** Session 26: a read footer also targeted by an unused header relationship (#607). */
export const footerWithHeaderRelationship = (
    declared = storyContentType("header"),
    headerReferenced = false,
    headerFirst = false,
): Record<string, string> => {
    const parts = storyContentTypeParts("footer", declared);
    const rels = parts["word/_rels/document.xml.rels"]!;
    const header = `<Relationship Id="rIdH" Type="${R}/header" Target="footer9.xml"/>`;
    parts["word/_rels/document.xml.rels"] = headerFirst
        ? rels.replace('<Relationship Id="rIdX"', `${header}<Relationship Id="rIdX"`)
        : rels.replace("</Relationships>", `${header}</Relationships>`);
    if (headerReferenced) {
        parts["word/document.xml"] = parts["word/document.xml"]!.replace(
            "<w:sectPr>",
            '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>',
        );
    }
    return parts;
};
