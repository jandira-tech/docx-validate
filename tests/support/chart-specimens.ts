/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Session 36's chart specimens (#645): a chart part whose
 * `<c:externalData r:id="rId1"/>` names nothing because the chart has no
 * `.rels` beside it, related from the main document's sidecar as `rId9`
 * (`…/relationships/chart`), stored under `word/` and outside it, referenced
 * by an inline drawing or by nothing. Word 16.115 opens a related-but-
 * unreferenced chart clean and offers recovery for the referenced shape
 * (chunks s36b/s36d in docs/plans/evidence/word-probes-2026-10-24.md).
 */

import { glossaryStoryParts } from "./glossary-content-type-specimens";
import { MINIMAL_PARTS } from "./minimal-docx";
import { storyContentType, storyContentTypeParts } from "./story-content-type-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";

const CT = "[Content_Types].xml";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";

export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

/** `text` with the one occurrence of `from` replaced; throws when `from` is absent, so a shape cannot silently stay unchanged. */
const replaceOnce = (text: string, from: string, to: string): string => {
    if (!text.includes(from)) throw new Error(`missing ${from}`);
    return text.replace(from, to);
};

/** The s36b bare plot area: `c:externalData`'s `r:id` names nothing while the chart has no sidecar. */
export const BARE_CHART_XML =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${C}" xmlns:r="${R}"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>` +
    `<c:externalData r:id="rId1"/></c:chartSpace>`;

const BAR_PLOT_AREA =
    '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
    '<c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
    '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea>';

/** The s36d schema-valid clustered bar chart; `externalData` keeps the dangling `rId1`, its absence is the drawing's control. */
export const barChartXml = (externalData: boolean): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${BAR_PLOT_AREA}</c:chart>` +
    `${externalData ? '<c:externalData r:id="rId1"/>' : ""}</c:chartSpace>`;

/** The inline `wp:inline` drawing whose `c:chart` names the chart relationship `rId9` (s36d). */
export const CHART_DRAWING =
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Chart 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';

/**
 * Session 36's `chartParts` shape (s36b): the bare chart stored at `part`,
 * related by the main document's sidecar as `rId9` when `target` is given
 * and referenced by no drawing. Omit `target` for the unrelated control.
 */
export const chartSpecimenParts = (part: string, target?: string): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [part]: BARE_CHART_XML,
    ...(target === undefined
        ? {}
        : {
              [SIDECAR]: replaceOnce(
                  MINIMAL_PARTS[SIDECAR]!,
                  "</Relationships>",
                  `<Relationship Id="rId9" Type="${R}/chart" Target="${target}"/></Relationships>`,
              ),
          }),
    [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${CHART_CONTENT_TYPE}"/></Types>`),
});

/**
 * Session 36's `referencedChartParts` shape (s36d): the clustered bar chart
 * stored at `part`, related as `rId9` and referenced by `CHART_DRAWING` in
 * the main document. `externalData: false` is the drawing's own control.
 */
export const referencedChartSpecimenParts = (part: string, target: string, externalData: boolean): Record<string, string> => ({
    ...chartSpecimenParts(part, target),
    [part]: barChartXml(externalData),
    [MAIN]: replaceOnce(
        replaceOnce(MINIMAL_PARTS[MAIN]!, "<w:document ", `<w:document xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}" `),
        "<w:sectPr>",
        `${CHART_DRAWING}<w:sectPr>`,
    ),
});

/** `text` with the first match of `pattern` replaced; throws when nothing matches. */
const replaceFirst = (text: string, pattern: RegExp, to: string): string => {
    if (!pattern.test(text)) throw new Error(`missing ${pattern}`);
    return text.replace(pattern, to);
};

/** Where session 39 related the chart from (#661): a referenced header, or a reached glossary's footnotes part or document. */
export type ChartStorySource = "header" | "footnote" | "glossary";

/**
 * Session 39's story-sourced chart (s39a/s39c, #661): the s36d clustered bar
 * chart at `word/charts/chart1.xml`, related as `rId9` from the sidecar of a
 * header the main document references or of a reached glossary's footnotes
 * part or document. `referenced` puts `CHART_DRAWING` before that part's
 * first paragraph. Word 16.115 opened every unreferenced shape clean and
 * offered recovery for every referenced one.
 */
export const storyChartSpecimenParts = (source: ChartStorySource, referenced: boolean): Record<string, string> => {
    const chart = "word/charts/chart1.xml";
    const parts =
        source === "header" ? storyContentTypeParts("header", storyContentType("header")) : glossaryStoryParts("footnotes", true, true);
    const part = { header: "word/header9.xml", footnote: "word/glossary/footnotes.xml", glossary: "word/glossary/document.xml" }[source];
    const sidecar = {
        header: "word/_rels/header9.xml.rels",
        footnote: "word/glossary/_rels/footnotes.xml.rels",
        glossary: "word/glossary/_rels/document.xml.rels",
    }[source];
    parts[chart] = barChartXml(true);
    parts[CT] = replaceOnce(parts[CT]!, "</Types>", `<Override PartName="/${chart}" ContentType="${CHART_CONTENT_TYPE}"/></Types>`);
    parts[sidecar] = replaceOnce(
        parts[sidecar] ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR}"></Relationships>`,
        "</Relationships>",
        `<Relationship Id="rId9" Type="${R}/chart" Target="${source === "header" ? "charts/chart1.xml" : "../charts/chart1.xml"}"/></Relationships>`,
    );
    if (referenced) {
        let xml = replaceFirst(parts[part]!, /(<w:\w+)(\s)/, `$1 xmlns:w14="${W14}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}"$2`);
        xml = replaceFirst(xml.replaceAll(` xmlns:r="${R}"`, ""), /(<w:\w+)/, `$1 xmlns:r="${R}"`);
        parts[part] = replaceOnce(xml, "<w:p>", `${CHART_DRAWING}<w:p>`);
    }
    return parts;
};

/**
 * Session 39's Fallback-wrapped drawing (#661): `referencedChartSpecimenParts`
 * with the drawing in an `mc:Fallback`. `selected` puts an unselected `foo`
 * Choice before it, so Word reads the Fallback (recovery); otherwise an empty
 * `w14` Choice is selected and the drawing is never read (clean).
 */
export const fallbackChartSpecimenParts = (selected: boolean): Record<string, string> => {
    const parts = referencedChartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml", true);
    parts[MAIN] = replaceOnce(
        replaceOnce(
            parts[MAIN]!,
            CHART_DRAWING,
            `<mc:AlternateContent><mc:Choice Requires="${selected ? "foo" : "w14"}"/><mc:Fallback>${CHART_DRAWING}</mc:Fallback></mc:AlternateContent>`,
        ),
        "<w:document ",
        '<w:document xmlns:foo="urn:foo" ',
    );
    return parts;
};

/** Session 39 (#661): a second `chart` relationship `rId8` to the same part; `referenced` names `rId9` from the drawing. */
export const twoChartRelationshipsSpecimenParts = (referenced: boolean): Record<string, string> => {
    const parts = referenced
        ? referencedChartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml", true)
        : chartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml");
    parts[SIDECAR] = replaceOnce(
        parts[SIDECAR]!,
        "</Relationships>",
        `<Relationship Id="rId8" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`,
    );
    return parts;
};

/** Session 39 (#661): the relationship targets `CHARTS/%63hart1.xml`, an OPC-equivalent spelling of the stored chart name. */
export const equivalentTargetChartSpecimenParts = (referenced: boolean): Record<string, string> => {
    const parts = referenced
        ? referencedChartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml", true)
        : chartSpecimenParts("word/charts/chart1.xml", "charts/chart1.xml");
    parts[SIDECAR] = replaceOnce(parts[SIDECAR]!, 'Target="charts/chart1.xml"', 'Target="CHARTS/%63hart1.xml"');
    return parts;
};

const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/** The s44a inline picture whose `a:blip r:embed` names the image relationship `rId8`. */
export const PICTURE_DRAWING =
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

/** A relationship from the main document's sidecar in a session 44 specimen. */
export interface ReachingRelationship {
    readonly id: string;
    readonly type: string;
}

/**
 * Session 44's reach shapes (s44a–s44c, #647): `BARE_CHART_XML` (or, with
 * `externalData: false`, the same chart with no `c:externalData`) stored at
 * `part`, a chart Override, and one relationship per entry of `reach` from the
 * main document's sidecar, each targeting the chart; `drawing`, when given, is
 * put before the main document's `w:sectPr`. Word 16.115 opened every shape
 * whose only reference is a picture's `a:blip r:embed`, or that nothing
 * references, clean, and refused the chart a `c:chart r:id` references.
 */
export const reachedChartSpecimenParts = (
    part: string,
    reach: readonly ReachingRelationship[],
    drawing?: string,
    externalData = true,
): Record<string, string> => {
    const target = part.startsWith("word/") ? part.slice("word/".length) : `../${part}`;
    return {
        ...MINIMAL_PARTS,
        [part]: externalData ? BARE_CHART_XML : BARE_CHART_XML.replace('<c:externalData r:id="rId1"/>', ""),
        [SIDECAR]: replaceOnce(
            MINIMAL_PARTS[SIDECAR]!,
            "</Relationships>",
            `${reach.map(({ id, type }) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("")}</Relationships>`,
        ),
        [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${CHART_CONTENT_TYPE}"/></Types>`),
        ...(drawing === undefined
            ? {}
            : {
                  [MAIN]: replaceOnce(
                      replaceOnce(
                          MINIMAL_PARTS[MAIN]!,
                          "<w:document ",
                          `<w:document xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:c="${C}" xmlns:r="${R}" `,
                      ),
                      "<w:sectPr>",
                      `${drawing}<w:sectPr>`,
                  ),
              }),
    };
};
