/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { MINIMAL_PARTS } from "./minimal-docx";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Move the minimal package's main document and its sidecar, keeping its styles related. */
export const compatibilityMainParts = (main: string, xml = MINIMAL_PARTS["word/document.xml"]!): Record<string, string> => {
    const parts = { ...MINIMAL_PARTS };
    const folder = path.posix.dirname(main);
    delete parts["word/document.xml"];
    delete parts["word/_rels/document.xml.rels"];
    parts[main] = xml;
    parts[`${folder}/_rels/${path.posix.basename(main)}.rels`] = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
        'Target="styles.xml"',
        `Target="${path.posix.relative(folder, "word/styles.xml")}"`,
    );
    parts["_rels/.rels"] = parts["_rels/.rels"]!.replace("word/document.xml", main);
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace("/word/document.xml", `/${main}`);
    return parts;
};

/** Relate a story from the main document, optionally adding the reference that makes Word read it. */
export const compatibilityStoryParts = (part: string, segment: string, xml: string, used = true): Record<string, string> => {
    const parts = { ...MINIMAL_PARTS, [part]: xml };
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `<Relationship Id="story" Type="${R}/${segment}" Target="${path.posix.relative("word", part)}"/></Relationships>`,
    );
    const type = segment === "glossaryDocument" ? "document.glossary" : segment;
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/></Types>`,
    );
    if (used) {
        if (segment === "header" || segment === "footer") {
            parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                "<w:sectPr>",
                `<w:sectPr><w:${segment}Reference xmlns:r="${R}" r:id="story" w:type="default"/>`,
            );
        } else if (segment !== "glossaryDocument") {
            const reference = { comments: "comment", footnotes: "footnote", endnotes: "endnote" }[segment];
            parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                "<w:r><w:t>Hello</w:t></w:r>",
                `<w:r><w:${reference}Reference w:id="1"/></w:r>`,
            );
        }
    }
    return parts;
};
