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

import { MINIMAL_PARTS } from "./minimal-docx";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

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
