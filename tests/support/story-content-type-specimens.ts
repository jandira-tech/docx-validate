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

/** Session 24: the minimal package plus a header/footer, or session 23's reached glossary footer (#600). */
export const storyContentTypeParts = (
    segment: "header" | "footer",
    declared = storyContentType(segment === "header" ? "footer" : "header"),
    via: "main" | "glossary" = "main",
    referenced = true,
): Record<string, string> => {
    if (via === "glossary") {
        const parts = glossaryStoryParts("footer", true, referenced);
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
] as const;
