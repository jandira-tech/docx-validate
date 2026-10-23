/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Builds a minimal PPTX (one master, one layout, one slide) that passes the
 * full PPTX validation. Tests derive single-defect specimens from it, so a
 * failure points at the one defect the test introduced.
 */

import JSZip from "jszip";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PML = "application/vnd.openxmlformats-officedocument.presentationml";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
const EMPTY_TREE =
    `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` + `<p:grpSpPr/></p:spTree></p:cSld>`;

const relationships = (...rels: readonly string[]): string =>
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;

export const MINIMAL_PPTX_PARTS: Readonly<Record<string, string>> = Object.freeze({
    "[Content_Types].xml":
        `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="${PML}.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${PML}.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${PML}.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="${PML}.slide+xml"/>` +
        `</Types>`,
    "_rels/.rels": relationships(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`),
    "ppt/presentation.xml":
        `${XML_DECL}<p:presentation ${NS}>` +
        `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
        `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>` +
        `</p:presentation>`,
    "ppt/_rels/presentation.xml.rels": relationships(
        `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`,
        `<Relationship Id="rId2" Type="${REL}/slide" Target="slides/slide1.xml"/>`,
    ),
    "ppt/slideMasters/slideMaster1.xml":
        `${XML_DECL}<p:sldMaster ${NS}>${EMPTY_TREE}` +
        `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ` +
        `accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
        `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
        `</p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": relationships(
        `<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml":
        `${XML_DECL}<p:sldLayout ${NS} type="obj">${EMPTY_TREE}` + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relationships(
        `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`,
    ),
    "ppt/slides/slide1.xml": `${XML_DECL}<p:sld ${NS}>${EMPTY_TREE}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": relationships(
        `<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
    ),
});

/** Zip `parts` (defaults to the minimal package) into PPTX bytes. */
export async function buildPptx(parts: Readonly<Record<string, string | Uint8Array>> = MINIMAL_PPTX_PARTS): Promise<Buffer> {
    const zip = new JSZip();
    for (const [name, content] of Object.entries(parts)) zip.file(name, content);
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
