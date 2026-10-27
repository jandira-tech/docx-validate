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

const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";
/**
 * `parts` with a Strict main document and styles part, and Strict
 * relationship types and ids: every session behind a read-story co-reach
 * citation stored a Transitional main document, so this package is
 * unprobed (#857, #861). Story parts keep the roots `parts` gives them.
 */
export const uniformlyStrict = (parts: Record<string, string>): Record<string, string> => ({
    ...parts,
    "word/document.xml": parts["word/document.xml"]!.replaceAll(W, STRICT_W).replaceAll(R, STRICT_R),
    "word/styles.xml": parts["word/styles.xml"]!.replaceAll(W, STRICT_W),
    "_rels/.rels": parts["_rels/.rels"]!.replaceAll(R, STRICT_R),
    "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"]!.replaceAll(R, STRICT_R),
});

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

const GLOSSARY_TYPE = storyContentType("document.glossary");
const relsBody = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

/** Session 26's story payloads (#606). */
export const STORY_HEADER_XML = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
export const STORY_FOOTER_XML = `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`;

/** Session 26's `docWith` shape (#606): the main document's body sectPr carries the reference. */
const docWithStoryReference = (reference: "headerReference" | "footerReference", id = "rIdX"): string =>
    MINIMAL_PARTS["word/document.xml"]!.replace(`<w:document xmlns:w="${W}"`, `<w:document xmlns:w="${W}" xmlns:r="${R}"`).replace(
        "<w:sectPr>",
        `<w:sectPr><w:${reference} w:type="default" r:id="${id}"/>`,
    );

/**
 * Session 26's `mainPart` shape (#606): the sectPr `reference` names rIdX, a
 * `segment`-type relationship to `word/file`, typed `type`.
 */
export const storyReferenceParts = (
    reference: "headerReference" | "footerReference",
    segment: string,
    file: string,
    xml: string,
    type: string,
): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": docWithStoryReference(reference),
    [`word/${file}`]: xml,
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="rIdX" Type="${R}/${segment}" Target="${file}"/></Relationships>`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/word/${file}" ContentType="${type}"/></Types>`,
    ),
});

/** Session 26's `headerRefToStyles` shape (#606): the sectPr headerReference names rId1, the styles relationship MINIMAL_PARTS already carries. */
export const storyReferenceToOwnStylesParts = (): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/document.xml": docWithStoryReference("headerReference", "rId1"),
});

/**
 * Session 26's `glossaryRef` shape (#606): a reached glossary's docPartBody
 * sectPr carries the `reference` (rIdF), rIdF a `relType` relationship to
 * `word/glossary/file`.
 */
export const glossaryStoryReferenceParts = (
    reference: "headerReference" | "footerReference",
    relType: string,
    file: string,
    xml: string,
    type: string,
): Record<string, string> => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml":
        `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p><w:sectPr><w:${reference} w:type="default" r:id="rIdF"/></w:sectPr></w:docPartBody>` +
        `</w:docPart></w:docParts></w:glossaryDocument>`,
    "word/glossary/_rels/document.xml.rels": relsBody(`<Relationship Id="rIdF" Type="${R}/${relType}" Target="${file}"/>`),
    [`word/glossary/${file}`]: xml,
    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/word/glossary/document.xml" ContentType="${GLOSSARY_TYPE}"/><Override PartName="/word/glossary/${file}" ContentType="${type}"/></Types>`,
    ),
});
