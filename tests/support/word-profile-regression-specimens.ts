/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";

import { MINIMAL_PARTS } from "./minimal-docx";

/** Renamed main parts outside the ct-undeclared-part pass (#535, AR F1). */
export const GENERIC_RENAMED_MAIN_PARTS = ["word/document.bin", "docProps/document.xml"].flatMap((part) =>
    ["application/xml", "text/xml"].map((contentType) => {
        const parts = { ...MINIMAL_PARTS };
        delete parts["word/document.xml"];
        delete parts["word/_rels/document.xml.rels"];
        parts[part] = MINIMAL_PARTS["word/document.xml"]!;
        parts[`${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`] = MINIMAL_PARTS[
            "word/_rels/document.xml.rels"
        ]!.replace('Target="styles.xml"', 'Target="/word/styles.xml"');
        parts["_rels/.rels"] = MINIMAL_PARTS["_rels/.rels"]!.replace("word/document.xml", part);
        parts["[Content_Types].xml"] = MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            /<Override PartName="\/word\/document.xml"[^>]*\/>/,
            "",
        )
            .replace('ContentType="application/xml"', `ContentType="${contentType}"`)
            .replace("</Types>", `<Default Extension="bin" ContentType="${contentType}"/></Types>`);
        return { part, contentType, parts };
    }),
);

/** Lazy usage is unknown when the main declaration fails byte parsing (#536, AR F2). */
export const BYTE_REJECTED_MAIN_PARTS = ["header", "endnotes", "footer", "comments", "footnotes"].map((segment) => {
    const part = `word/${segment}.xml`;
    const root = segment === "header" ? "hdr" : segment === "footer" ? "ftr" : segment;
    return {
        segment,
        part,
        parts: {
            ...MINIMAL_PARTS,
            "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"'),
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                "</Relationships>",
                `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${segment}" Target="${segment}.xml"/></Relationships>`,
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${segment}+xml"/></Types>`,
            ),
            [part]: `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`,
        },
    };
});
