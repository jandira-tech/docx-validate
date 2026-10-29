/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
export const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
export const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
export const WNE = "http://schemas.microsoft.com/office/word/2006/wordml";
export const document = (body: string, w = W): string =>
    `<w:document xmlns:w="${w}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:mc="${MC}" xmlns:r="${R}" xmlns:rs="${RS}" xmlns:m="${M}" xmlns:wne="${WNE}" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"><w:body>${body}</w:body></w:document>`;
export const rels = (body: string): string =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
export const relation = (id: string, type: string, mode = "", target = "part.xml"): string =>
    `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ""}/>`;
