/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { withTempDir } from "../src/lib/run-cli";
import { WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { document, rels, relation } from "./support/word-rule-test-xml";

const CODE = "word-part-reference-relationship";
const HERE = path.dirname(fileURLToPath(import.meta.url));

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CS = "http://purl.oclc.org/ooxml/drawingml/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DGMS = "http://purl.oclc.org/ooxml/drawingml/diagram";
const O = "urn:schemas-microsoft-com:office:office";
const V = "urn:schemas-microsoft-com:vml";
const CT = "[Content_Types].xml";
const SIDECAR = "word/_rels/document.xml.rels";
const MAIN = "word/document.xml";
const CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHART_INSIDE = "word/charts/chart1.xml";
const CHART_OUTSIDE = "charts/chart1.xml";

const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};
const relationship = (id: string, type: string, target: string, external = false): string =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`;
const OFFICE = (segment: string): string => `${R}/${segment}`;
const withOverrides = (...overrides: [string, string][]): string =>
    replaceOnce(
        MINIMAL_PARTS[CT]!,
        "</Types>",
        `${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join("")}</Types>`,
    );

/**
 * Session 45's drawings (s45builder.mts in docs/plans/evidence's builder
 * appendix): the exact shapes Word 16.115 judged.
 */
const CHART_DRAWING = (id: string): string =>
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Chart 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${C}">` +
    `<c:chart xmlns:c="${C}" r:id="${id}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const DIAGRAM_DRAWING = (id: string): string =>
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Diagram 1"/><a:graphic><a:graphicData uri="${DGM}">` +
    `<dgm:relIds xmlns:dgm="${DGM}" r:dm="${id}" r:lo="${id}" r:qs="${id}" r:cs="${id}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const OLE_OBJECT = (id: string): string =>
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:object>' +
    `<o:OLEObject xmlns:o="${O}" Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="${id}"/></w:object></w:r></w:p>`;
const PICTURE_DRAWING = (id: string): string =>
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}"><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
const VML_IMAGEDATA = (id: string): string =>
    '<w:p w14:paraId="00000002" w14:textId="00000002"><w:r><w:pict>' +
    `<v:shape xmlns:v="${V}" xmlns:o="${O}" id="s1" style="width:10pt;height:10pt"><v:imagedata o:relid="${id}"/></v:shape></w:pict></w:r></w:p>`;

/** The minimal package plus `parts`, with `relationships` appended to the main document's sidecar. */
const extended = (
    parts: Record<string, string | Uint8Array>,
    relationships: string[],
    overrides: [string, string][],
    pngDefault = false,
): Record<string, string | Uint8Array> => ({
    ...MINIMAL_PARTS,
    ...parts,
    [SIDECAR]: replaceOnce(MINIMAL_PARTS[SIDECAR]!, "</Relationships>", `${relationships.join("")}</Relationships>`),
    [CT]: replaceOnce(
        withOverrides(...overrides),
        '<Default Extension="rels"',
        `${pngDefault ? '<Default Extension="png" ContentType="image/png"/>' : ""}<Default Extension="rels"`,
    ),
});
const withDrawing = (parts: Record<string, string | Uint8Array>, drawing: string): Record<string, string | Uint8Array> => ({
    ...parts,
    [MAIN]: replaceOnce(
        replaceOnce(
            MINIMAL_PARTS[MAIN]!,
            "<w:document ",
            `<w:document xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:c="${C}" xmlns:r="${R}" `,
        ),
        "<w:sectPr>",
        `${drawing}<w:sectPr>`,
    ),
});

const PNG = Uint8Array.from(
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64"),
);
/** Session 45's #701 cause package: a real PNG reached by an image relationship (`rId8`). */
const pngReached = (): Record<string, string | Uint8Array> =>
    extended({ "word/media/image1.png": PNG }, [relationship("rId8", OFFICE("image"), "media/image1.png")], [], true);

const BAR_PLOT_AREA =
    '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
    '<c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
    '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea>';
const barChart = (externalData: boolean): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${BAR_PLOT_AREA}</c:chart>${externalData ? '<c:externalData r:id="rId1"/>' : ""}</c:chartSpace>`;
/** Session 36d's bar chart at `word/charts/chart1.xml` (no sidecar), reached from the main document's sidecar as `id` of `segment`. */
const barReached = (externalData: boolean, id: string, segment: string): Record<string, string | Uint8Array> =>
    extended(
        { [CHART_INSIDE]: barChart(externalData) },
        [relationship(id, OFFICE(segment), "charts/chart1.xml")],
        [[CHART_INSIDE, CHART_CT]],
    );

/** Session 45's bare chart (a plotArea-less chartSpace with `c:externalData`), stored at `part` with no sidecar. */
const BARE_CHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:r="${R}"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart><c:externalData r:id="rId1"/></c:chartSpace>`;
const targetOf = (part: string): string => (part === CHART_INSIDE ? "charts/chart1.xml" : `../${part}`);
const typed = (part: string, id: string, segment: string, external = false): string =>
    relationship(id, OFFICE(segment), external ? "charts/chart1.xml" : targetOf(part), external);
const reached = (part: string, ...relationships: string[]): Record<string, string | Uint8Array> => ({
    ...MINIMAL_PARTS,
    [part]: BARE_CHART,
    [SIDECAR]: replaceOnce(MINIMAL_PARTS[SIDECAR]!, "</Relationships>", `${relationships.join("")}</Relationships>`),
    [CT]: withOverrides([part, CHART_CT]),
});

/**
 * Session 45's refused specimens (docs/plans/evidence/word-probes-2026-10-24.md):
 * every row Word 16.115 refused in both runs, built on `MINIMAL_PARTS`.
 */
const REFUSED_TRANSITIONAL: Record<string, () => Promise<Buffer>> = {
    // c:chart r:id → an image relationship to a real PNG (s45e, s45j).
    "png-image-relationship-chart-reference": async () => buildDocx(withDrawing(pngReached(), CHART_DRAWING("rId8"))),
    // c:chart r:id → an image relationship to the bar chart with no c:externalData (s45k).
    "bar-noexternal-image-reach-chart-reference": async () =>
        buildDocx(withDrawing(barReached(false, "rId8", "image"), CHART_DRAWING("rId8"))),
    // c:chart r:id → an External chart relationship whose target is the stored chart part (s45c).
    "chart-word-external-chart-relationship-chart-reference": async () =>
        buildDocx(withDrawing(reached(CHART_INSIDE, typed(CHART_INSIDE, "rId9", "chart", true)), CHART_DRAWING("rId9"))),
    // c:chart r:id → an External chart relationship to a URL, no part (s45e, s45i).
    "remote-external-chart-relationship-chart-reference": async () =>
        buildDocx(
            withDrawing(
                extended({}, [relationship("rId9", OFFICE("chart"), "http://example.com/chart1.xml", true)], []),
                CHART_DRAWING("rId9"),
            ),
        ),
    // dgm:relIds r:dm → a chart relationship, bar chart with c:externalData and no sidecar (s45i).
    "bar-chart-reach-relids-reference": async () => buildDocx(withDrawing(barReached(true, "rId9", "chart"), DIAGRAM_DRAWING("rId9"))),
    // dgm:relIds r:dm → a chart relationship, bar chart without c:externalData (s45h).
    "bar-noexternal-chart-reach-relids-reference": async () =>
        buildDocx(withDrawing(barReached(false, "rId9", "chart"), DIAGRAM_DRAWING("rId9"))),
    // dgm:relIds r:dm → a chart relationship, the bare chart under word/ (s45b).
    "chart-word-chart-reach-relids-reference": async () =>
        buildDocx(withDrawing(reached(CHART_INSIDE, typed(CHART_INSIDE, "rId9", "chart")), DIAGRAM_DRAWING("rId9"))),
    // dgm:relIds r:dm → a chart relationship, the bare chart outside word/ (s45d).
    "chart-outside-word-chart-reach-relids-reference": async () =>
        buildDocx(withDrawing(reached(CHART_OUTSIDE, typed(CHART_OUTSIDE, "rId9", "chart")), DIAGRAM_DRAWING("rId9"))),
    // dgm:relIds r:dm → an image relationship to a real PNG (s45e, s45j).
    "png-image-relationship-relids-reference": async () => buildDocx(withDrawing(pngReached(), DIAGRAM_DRAWING("rId8"))),
    // o:OLEObject r:id → an image relationship to a real PNG (s45e, s45k).
    "png-image-relationship-oleobject-reference": async () => buildDocx(withDrawing(pngReached(), OLE_OBJECT("rId8"))),
};

/** Session 45's clean controls: the rows Word 16.115 opened clean in both runs. */
const CLEAN_TRANSITIONAL: Record<string, () => Promise<Buffer>> = {
    // c:chart r:id → an internal chart relationship (bar chart, no sidecar) (s45h).
    "bar-noexternal-chart-reach-chart-reference": async () =>
        buildDocx(withDrawing(barReached(false, "rId9", "chart"), CHART_DRAWING("rId9"))),
    // a:blip r:embed → an image relationship whose target is the chart part (s45d).
    "chart-word-image-reach-blip-reference": async () =>
        buildDocx(withDrawing(reached(CHART_INSIDE, typed(CHART_INSIDE, "rId8", "image")), PICTURE_DRAWING("rId8"))),
    // v:imagedata o:relid → a chart relationship (s45b).
    "chart-word-chart-reach-vml-relid-reference": async () =>
        buildDocx(withDrawing(reached(CHART_INSIDE, typed(CHART_INSIDE, "rId9", "chart")), VML_IMAGEDATA("rId9"))),
};

/**
 * Session 46's specimens: one attribute of `word/document.xml` in
 * `tests/fixtures/working/word-clean-strict01.docx` changed, exactly as the
 * session's builder did. The fixture is an ISO Strict document Word opened
 * clean, holding a SmartArt (`dgm:relIds` `rId35`–`rId38`), a chart
 * (`rId8`), an embedded Excel object (`o:OLEObject r:id="rId19"`, a
 * `package` relationship, in the `mc:Choice` Word selects) and an EMF
 * picture (`rId18`, `image`).
 */
const STRICT01 = path.join(HERE, "fixtures", "working", "word-clean-strict01.docx");
const strictVariant = async (from?: string, to?: string): Promise<Buffer> => {
    const zip = await JSZip.loadAsync(readFileSync(STRICT01));
    const parts: Record<string, string | Uint8Array> = {};
    for (const [name, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        parts[name] = await entry.async("uint8array");
    }
    if (from !== undefined && to !== undefined) {
        parts[MAIN] = replaceOnce(Buffer.from(parts[MAIN] as Uint8Array).toString("utf8"), from, to);
    }
    return buildDocx(parts);
};
const RELIDS = 'r:dm="rId35" r:lo="rId36" r:qs="rId37" r:cs="rId38"';
const relIds = (attribute: "dm" | "lo" | "qs" | "cs", id: string): Promise<Buffer> =>
    strictVariant(RELIDS, RELIDS.replace(new RegExp(`r:${attribute}="rId\\d+"`), `r:${attribute}="${id}"`));
const CHART_REF = 'r:id="rId8"/></a:graphicData>';
const chartNames = (id: string): Promise<Buffer> => strictVariant(CHART_REF, `r:id="${id}"/></a:graphicData>`);
const OLE_REF = 'ObjectID="_1509346857" r:id="rId19"/>';
const oleNames = (id: string): Promise<Buffer> => strictVariant(OLE_REF, `ObjectID="_1509346857" r:id="${id}"/>`);

/** Session 46's refused specimens (s46a–s46c): every row Word 16.115 refused in both runs. */
const REFUSED_STRICT: Record<string, () => Promise<Buffer>> = {
    "strict01-relids-dm-image": () => relIds("dm", "rId18"),
    "strict01-relids-dm-dangling": () => relIds("dm", "rId999"),
    "strict01-chart-names-diagramdata": () => chartNames("rId35"),
    "strict01-chart-names-package": () => chartNames("rId19"),
    "strict01-chart-names-dangling": () => chartNames("rId999"),
    "strict01-oleobject-names-image": () => oleNames("rId18"),
    "strict01-oleobject-names-chart": () => oleNames("rId8"),
    "strict01-oleobject-names-dangling": () => oleNames("rId999"),
};

/** Session 46's clean controls: the fixture re-zipped unchanged, and the tolerated `r:lo`/`r:qs`/`r:cs` rows. */
const CLEAN_STRICT: Record<string, () => Promise<Buffer>> = {
    // The unchanged fixture: also the o:OLEObject `package` control Word opened clean.
    "strict01-rebuilt": () => strictVariant(),
    "strict01-relids-lo-image": () => relIds("lo", "rId18"),
    "strict01-relids-qs-image": () => relIds("qs", "rId18"),
    "strict01-relids-cs-image": () => relIds("cs", "rId18"),
    "strict01-relids-lo-chart": () => relIds("lo", "rId8"),
};

/** The story namespaces the session 51 drawings sit under, declared on a story root the minimal main does not provide. */
const STORY_ROOT_NS = `xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"`;

/**
 * Session 51's scope specimens (s51builder.mts in the evidence appendix):
 * the same chart reference — a `c:chart r:id` naming an image relationship
 * to a real PNG — in the parts Word reads and in the parts it does not
 * (#726; session 51, chunk s51a).
 */
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const HEADER_CT = `${WML}.header+xml`;
const GLOSSARY_CT = `${WML}.document.glossary+xml`;
const relsPart = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
/** A header whose only paragraph holds the chart drawing, its `rId1` an image relationship to the PNG. */
const headerWithChart = (): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14" ${STORY_ROOT_NS}>${CHART_DRAWING("rId1")}</w:hdr>`;
const headerImageSidecar = (): string => relsPart(relationship("rId1", OFFICE("image"), "media/image1.png"));
/** A header the body references through `w:headerReference r:id="rIdH"`, stored as `word/<file>`. */
const referencedHeaderParts = (file: string): Record<string, string | Uint8Array> =>
    extended(
        {
            "word/media/image1.png": PNG,
            [`word/${file}`]: headerWithChart(),
            [`word/_rels/${file}.rels`]: headerImageSidecar(),
            [MAIN]: replaceOnce(
                replaceOnce(MINIMAL_PARTS[MAIN]!, "<w:document ", `<w:document xmlns:r="${R}" `),
                "<w:sectPr>",
                '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>',
            ),
        },
        [relationship("rIdH", OFFICE("header"), file)],
        [[`word/${file}`, HEADER_CT]],
        true,
    );
/** The main document stored as `word/main.xml`, its body holding the same chart reference (`rIdI` an image relationship). */
const renamedMainParts = (): Record<string, string | Uint8Array> => {
    const parts: Record<string, string | Uint8Array> = {
        ...MINIMAL_PARTS,
        "word/media/image1.png": PNG,
        "word/main.xml": replaceOnce(
            replaceOnce(MINIMAL_PARTS[MAIN]!, "<w:document ", `<w:document ${STORY_ROOT_NS} `),
            "<w:sectPr>",
            `${CHART_DRAWING("rIdI")}<w:sectPr>`,
        ),
        "word/_rels/main.xml.rels": replaceOnce(
            MINIMAL_PARTS[SIDECAR]!,
            "</Relationships>",
            `${relationship("rIdI", OFFICE("image"), "media/image1.png")}</Relationships>`,
        ),
        "_rels/.rels": replaceOnce(MINIMAL_PARTS["_rels/.rels"]!, 'Target="word/document.xml"', 'Target="word/main.xml"'),
        [CT]: replaceOnce(
            replaceOnce(
                MINIMAL_PARTS[CT]!,
                '<Default Extension="rels"',
                '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"',
            ),
            '<Override PartName="/word/document.xml"',
            '<Override PartName="/word/main.xml"',
        ),
    };
    delete parts[MAIN];
    delete parts[SIDECAR];
    return parts;
};
/** A header the main document relates but no `w:headerReference` names, with its sidecar and without one. */
const unusedHeaderParts = (withSidecar: boolean): Record<string, string | Uint8Array> =>
    extended(
        {
            "word/media/image1.png": PNG,
            "word/header1.xml": headerWithChart(),
            ...(withSidecar ? { "word/_rels/header1.xml.rels": headerImageSidecar() } : {}),
        },
        [relationship("rIdH", OFFICE("header"), "header1.xml")],
        [["word/header1.xml", HEADER_CT]],
        true,
    );
/** A glossary document no relationship reaches, its docPartBody holding the same chart reference. */
const unreachedGlossaryParts = (): Record<string, string | Uint8Array> =>
    extended(
        {
            "word/media/image1.png": PNG,
            "word/glossary/document.xml":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:glossaryDocument xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14" ${STORY_ROOT_NS}>` +
                '<w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>' +
                '<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>' +
                `<w:docPartBody>${CHART_DRAWING("rId1")}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`,
            "word/glossary/_rels/document.xml.rels": relsPart(relationship("rId1", OFFICE("image"), "../media/image1.png")),
        },
        [],
        [["word/glossary/document.xml", GLOSSARY_CT]],
        true,
    );
/** The main document's chart drawing with the document sidecar deleted: a story root with no relationships part. */
const noSidecarMainParts = (): Record<string, string | Uint8Array> => {
    const parts = withDrawing({ ...MINIMAL_PARTS }, CHART_DRAWING("rId9"));
    delete parts[SIDECAR];
    return parts;
};

/** The check under test, resolved through the registry so these rows fail on assertions first. */
const check = () => {
    const found = WORD_RULE_CHECKS.find((def) => def.id === CODE);
    expect(found?.id).toBe(CODE);
    return found!;
};

const chartReference = (id: string): string =>
    `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="${id}"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`;
const relIdsReference = (dm: string, others = ""): string =>
    `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="${dm}"${others}/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`;
const oleObjectReference = (id: string): string =>
    `<w:p><w:r><w:object><o:OLEObject xmlns:o="${O}" Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="${id}"/></w:object></w:r></w:p>`;

/** A package whose main document is `word/document.xml`, with the given body and sidecar relationships (#726's read set). */
const mainPackage = (body: string, docRels: string): Record<string, string> => ({
    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
    "word/document.xml": document(body),
    "word/_rels/document.xml.rels": rels(docRels),
});

describe(`${CODE} (unit rows)`, () => {
    it("reports the Word-rejected chart shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries(mainPackage(chartReference("rId1"), relation("rId1", "image")));
        expect(await check().run(ctx)).toEqual([
            {
                severity: "error",
                code: CODE,
                path: "word/document.xml",
                message:
                    `<c:chart> r:id="rId1" does not reference an internal chart relationship (Relationship 'rId1' has Type '${R}/image'). ` +
                    "Microsoft Word offers to recover the file. " +
                    `Use a Relationship of Type ${R}/chart (or the ISO Strict ${RS}/chart) internal to the package, or point the r:id at one.`,
            },
        ]);
    });

    it("accepts an internal chart relationship in either chart and relationships namespace", async () => {
        const transitional = ValidationContext.fromEntries(mainPackage(chartReference("rId1"), relation("rId1", "chart")));
        expect(await check().run(transitional)).toEqual([]);
        const strict = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document(
                `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${CS}"><c:chart xmlns:c="${CS}" rs:id="rId1"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`,
            ),
            "word/_rels/document.xml.rels": rels(`<Relationship Id="rId1" Type="${RS}/chart" Target="charts/chart1.xml"/>`),
        });
        expect(await check().run(strict)).toEqual([]);
    });

    it.each([
        ["dangling", "rId9", undefined, "no Relationship with Id 'rId9'"],
        ["wrong type", "rId1", relation("rId1", "image"), `Relationship 'rId1' has Type '${R}/image'`],
        ["external", "rId1", relation("rId1", "chart", "External"), "Relationship 'rId1' has TargetMode 'External'"],
    ])("explains why a chart reference is refused: %s", async (_name, rid, relationshipElement, why) => {
        const ctx = ValidationContext.fromEntries(mainPackage(chartReference(rid), relationshipElement ?? ""));
        const found = await check().run(ctx);
        expect(found).toHaveLength(1);
        expect(found[0]?.message).toContain(`(${why})`);
        expect(found[0]?.message).toContain("Microsoft Word offers to recover the file.");
    });

    it("reports dgm:relIds r:dm alone, not r:lo, r:qs or r:cs (#705, session 46)", async () => {
        const others = ' r:lo="rId2" r:qs="rId2" r:cs="rId2"';
        const mismatched = ValidationContext.fromEntries(
            mainPackage(relIdsReference("rId2", others), `${relation("rId1", "diagramData")}${relation("rId2", "image")}`),
        );
        expect(await check().run(mismatched)).toEqual([
            expect.objectContaining({
                code: CODE,
                path: "word/document.xml",
                severity: "error",
                message: expect.stringContaining('r:dm="rId2"'),
            }),
        ]);
        const tolerated = ValidationContext.fromEntries(
            mainPackage(
                relIdsReference("rId1", ' r:lo="rId2" r:qs="rId2" r:cs="rId2"'),
                `${relation("rId1", "diagramData")}${relation("rId2", "image")}`,
            ),
        );
        expect(await check().run(tolerated)).toEqual([]);
    });

    it("judges the ISO Strict dgm:relIds spelling too", async () => {
        const ctx = ValidationContext.fromEntries(
            mainPackage(
                `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${DGMS}"><dgm:relIds xmlns:dgm="${DGMS}" rs:dm="rId1"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`,
                `<Relationship Id="rId1" Type="${RS}/chart" Target="charts/chart1.xml"/>`,
            ),
        );
        expect(await check().run(ctx)).toEqual([expect.objectContaining({ code: CODE, severity: "error" })]);
    });

    it.each([
        ["package", ""],
        ["oleObject", ""],
        ["oleObject", "External"],
    ])("accepts an o:OLEObject naming a %s relationship%s", async (type, mode) => {
        const ctx = ValidationContext.fromEntries(mainPackage(oleObjectReference("rId1"), relation("rId1", type, mode)));
        expect(await check().run(ctx)).toEqual([]);
    });

    it.each([
        ["image", `Relationship 'rId1' has Type '${R}/image'`],
        ["chart", `Relationship 'rId1' has Type '${R}/chart'`],
        ["dangling", "no Relationship with Id 'rId9'"],
    ])("refuses an o:OLEObject naming a %s relationship", async (kind, why) => {
        const ctx = ValidationContext.fromEntries(
            mainPackage(oleObjectReference(kind === "dangling" ? "rId9" : "rId1"), kind === "dangling" ? "" : relation("rId1", kind)),
        );
        const found = await check().run(ctx);
        expect(found).toHaveLength(1);
        expect(found[0]?.message).toContain(`(${why})`);
    });

    it("a part with no relationships part is judged like one whose relationships part lacks the id (#751); an unparseable one is skipped", async () => {
        const absent = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document(chartReference("rId1")),
        });
        expect(await check().run(absent)).toEqual([
            expect.objectContaining({
                code: CODE,
                path: "word/document.xml",
                severity: "error",
                message: expect.stringContaining("no Relationship with Id 'rId1'"),
            }),
        ]);
        const unparseable = ValidationContext.fromEntries(mainPackage(chartReference("rId1"), "<Relationships"));
        expect(await check().run(unparseable)).toEqual([]);
    });

    it("leaves a chart reference with no r:id attribute alone, and judges an empty one (#728)", async () => {
        const ctx = ValidationContext.fromEntries(
            mainPackage(
                `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}"/><c:chart xmlns:c="${C}" r:id=""/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`,
                relation("rId1", "chart"),
            ),
        );
        const found = await check().run(ctx);
        expect(found).toHaveLength(1);
        expect(found[0]?.message).toContain('<c:chart> r:id=""');
        expect(found[0]?.message).toContain("(the id is empty)");
    });

    it("skips absent and unparseable story parts, and resolves an OPC-equivalent sidecar", async () => {
        expect(await check().run(ValidationContext.fromEntries({}))).toEqual([]);
        const unparseableRoot = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": "<broken",
        });
        expect(await check().run(unparseableRoot)).toEqual([]);
        const equivalent = ValidationContext.fromEntries({
            "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
            "word/document.xml": document(chartReference("rId1")),
            "word/_rels/document.xml.%72els": rels(relation("rId1", "chart")),
        });
        expect(await check().run(equivalent)).toEqual([]);
    });
});

describe(`${CODE} (session 45 Transitional specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    describe.each(Object.entries(REFUSED_TRANSITIONAL))("refused specimen %s", (name, make) => {
        it("word-valid fails with exactly one Word-blocking error", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/document.xml");
                expect(found[0]?.message).toContain("Microsoft Word offers to recover the file.");
            });
        });

        it("strict reports the code too (strict ⊇ word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
            });
        });

        it("lenient reports nothing for it", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.some((issue) => issue.code === CODE)).toBe(false);
            });
        });
    });

    describe.each(Object.entries(CLEAN_TRANSITIONAL))("clean specimen %s", (name, make) => {
        it("word-valid draws no such finding", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });
});

describe(`${CODE} (session 46 Strict specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    describe.each(Object.entries(REFUSED_STRICT))("refused specimen %s", (name, make) => {
        it("word-valid fails with exactly one Word-blocking error", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/document.xml");
                expect(found[0]?.message).toContain("Microsoft Word offers to recover the file.");
            });
        });

        it("strict reports the code too (strict ⊇ word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
            });
        });

        it("lenient reports nothing for it", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.some((issue) => issue.code === CODE)).toBe(false);
            });
        });
    });

    describe.each(Object.entries(CLEAN_STRICT))("clean specimen %s", (name, make) => {
        it("word-valid draws no such finding", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });
});

describe(`${CODE} (session 51 scope specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    it.each([
        ["word/header1.xml", () => referencedHeaderParts("header1.xml")],
        ["word/headerFirst.xml", () => referencedHeaderParts("headerFirst.xml")],
    ])("reports the chart reference in a header the body references, stored as %s", async (part, make) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "referenced-header.docx");
            writeFileSync(file, await buildDocx(make()));
            const result = await validate(file, { profile: "word-valid" });
            const found = result.issues.filter((issue) => issue.code === CODE);
            expect(result.valid).toBe(false);
            expect(found).toHaveLength(1);
            expect(found[0]?.path).toBe(part);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.some((issue) => issue.code === CODE && issue.path === part)).toBe(true);
        });
    });

    it("reports the chart reference in the main document stored as word/main.xml", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "main-renamed-chart-image.docx");
            writeFileSync(file, await buildDocx(renamedMainParts()));
            const result = await validate(file, { profile: "word-valid" });
            const found = result.issues.filter((issue) => issue.code === CODE);
            expect(result.valid).toBe(false);
            expect(found).toHaveLength(1);
            expect(found[0]?.path).toBe("word/main.xml");
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/main.xml")).toBe(true);
        });
    });

    // Session 51: Word 16.115 opened all three clean — Word reads neither an
    // unused header (sessions 15 and 51) nor an unreached glossary (sessions
    // 35 and 51), whatever their part names look like.
    it.each([
        ["a related header no w:headerReference names", () => unusedHeaderParts(true)],
        ["the same unused header with no relationships part", () => unusedHeaderParts(false)],
        ["a glossary document no relationship reaches", () => unreachedGlossaryParts()],
    ])("draws no finding for %s", async (_shape, make) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unread-part.docx");
            writeFileSync(file, await buildDocx(make()));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === CODE)).toEqual([]);
        });
    });

    it("a partial --checks word-part-reference-relationship run reaches the full run's verdict (#381)", async () => {
        await withTempDir(async (dir) => {
            const read = path.join(dir, "referenced-header.docx");
            writeFileSync(read, await buildDocx(referencedHeaderParts("headerFirst.xml")));
            const unread = path.join(dir, "unread-header.docx");
            writeFileSync(unread, await buildDocx(unusedHeaderParts(true)));
            for (const [file, expected] of [
                [read, true],
                [unread, false],
            ] as const) {
                const partial = await validate(file, { checks: [CODE] });
                expect(
                    partial.issues.some((issue) => issue.code === CODE),
                    file,
                ).toBe(expected);
            }
        });
    });

    it("judges a read part with no relationships part like one whose relationships part lacks the id (#751)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "main-no-sidecar.docx");
            writeFileSync(file, await buildDocx(noSidecarMainParts()));
            const result = await validate(file, { profile: "word-valid" });
            // Session 55 (s55a): a non-empty id a missing sidecar strands
            // names no relationship, so this rule reports it; the absent
            // sidecar itself stays rels-missing-sidecar's finding.
            expect(result.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            expect(result.issues.some((issue) => issue.code === "rels-missing-sidecar" && issue.path === "word/document.xml")).toBe(true);
        });
    });
});

/** Session 46's word-clean-strict01 package with its sidecar edited instead of its main document (session 52's B2). */
const DATA_REL =
    '<Relationship Id="rId35" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData" Target="diagrams/data1.xml"/>';
const strictSidecarVariant = async (from: string, to: string): Promise<Buffer> => {
    const zip = await JSZip.loadAsync(readFileSync(STRICT01));
    const parts: Record<string, string | Uint8Array> = {};
    for (const [name, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        parts[name] = await entry.async("uint8array");
    }
    parts[SIDECAR] = replaceOnce(Buffer.from(parts[SIDECAR] as Uint8Array).toString("utf8"), from, to);
    return buildDocx(parts);
};

/**
 * Session 52's specimens (s52builder.mts; #728): the empty-id shapes Word
 * 16.115 refused (C1–C3, the review's `r:dm`-only element for C3) and the
 * External `diagramData` relationship a `dgm:relIds r:dm` names (B2), with
 * the External `package` relationship Word opened clean (B1) and the
 * `r:dm`-alone confound as guards.
 */
const pngPackage = (): Record<string, string | Uint8Array> => extended({ "word/media/image1.png": PNG }, [], [], true);
const relIdsDmOnly = (dm: string): string =>
    `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="${dm}"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>`;
const REFUSED_728: Record<string, () => Promise<Buffer>> = {
    // s52b `chart-empty-id`: c:chart r:id="".
    "chart-empty-id": async () => buildDocx(withDrawing(pngPackage(), CHART_DRAWING(""))),
    // s52b `oleobject-empty-id`: o:OLEObject r:id="".
    "oleobject-empty-id": async () => buildDocx(withDrawing(pngPackage(), OLE_OBJECT(""))),
    // s52b `relids-empty-id`: dgm:relIds r:dm="" (the review's r:dm-only element).
    "relids-empty-id": async () => buildDocx(withDrawing(pngPackage(), relIdsDmOnly(""))),
    // s52b `relids-names-external-diagramdata`: r:dm names an External diagramData relationship.
    "relids-names-external-diagramdata": async () =>
        buildDocx(
            withDrawing(
                extended({}, [relationship("rId9", OFFICE("diagramData"), "http://example.com/data1.xml", true)], [], true),
                DIAGRAM_DRAWING("rId9"),
            ),
        ),
    // s52d `strict01-relids-dm-empty`: the SmartArt's r:dm emptied, r:lo/r:qs/r:cs intact.
    "strict01-relids-dm-empty": () => relIds("dm", ""),
    // s52d `strict01-relids-dm-external`: the SmartArt's diagramData relationship made External.
    "strict01-relids-dm-external": () =>
        strictSidecarVariant(
            DATA_REL,
            DATA_REL.replace('Target="diagrams/data1.xml"/>', 'Target="http://example.com/data1.xml" TargetMode="External"/>'),
        ),
};
const CLEAN_728: Record<string, () => Promise<Buffer>> = {
    // s52a `oleobject-names-external-package`: an External package relationship opens clean.
    "oleobject-names-external-package": async () =>
        buildDocx(
            withDrawing(
                extended({}, [relationship("rId9", OFFICE("package"), "http://example.com/o.xlsx", true)], [], true),
                OLE_OBJECT("rId9"),
            ),
        ),
    // s52d `strict01-relids-dm-only`: r:dm="rId35" alone names the real internal diagramData relationship.
    "strict01-relids-dm-only": () => strictVariant(RELIDS, 'r:dm="rId35"'),
};

describe(`${CODE} (session 52 specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    describe.each(Object.entries(REFUSED_728))("refused specimen %s", (name, make) => {
        it("word-valid fails with exactly one Word-blocking error", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/document.xml");
            });
        });

        it("strict reports the code too (strict \u2287 word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
            });
        });

        it("lenient reports nothing for it", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.issues.some((issue) => issue.code === CODE)).toBe(false);
            });
        });
    });

    it.each([
        ["chart-empty-id", '<c:chart> r:id=""', "the id is empty"],
        ["oleobject-empty-id", '<o:OLEObject> r:id=""', "the id is empty"],
        ["relids-empty-id", '<dgm:relIds> r:dm=""', "the id is empty"],
        ["strict01-relids-dm-empty", '<dgm:relIds> r:dm=""', "the id is empty"],
        ["relids-names-external-diagramdata", '<dgm:relIds> r:dm="rId9"', "Relationship 'rId9' has TargetMode 'External'"],
        ["strict01-relids-dm-external", '<dgm:relIds> r:dm="rId35"', "Relationship 'rId35' has TargetMode 'External'"],
    ] as const)("%s names the reference and why", async (name, reference, why) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${name}.docx`);
            writeFileSync(file, await REFUSED_728[name]!());
            const result = await validate(file, { profile: "word-valid" });
            const found = result.issues.filter((issue) => issue.code === CODE);
            expect(found[0]?.message).toContain(reference);
            expect(found[0]?.message).toContain(`(${why})`);
        });
    });

    it("says what session 52 observed for an empty r:dm, not one dialog (#752)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "relids-empty-id.docx");
            writeFileSync(file, await REFUSED_728["relids-empty-id"]!());
            const result = await validate(file, { profile: "word-valid" });
            const found = result.issues.filter((issue) => issue.code === CODE);
            expect(found[0]?.message).toContain("(the id is empty)");
            // s52b's `relids-empty-id` was an open error; only s52d's
            // strict01 mutation drew the recovery offer.
            expect(found[0]?.message).toContain("an open error");
            expect(found[0]?.message).not.toContain("Microsoft Word offers to recover the file");
        });
    });

    describe.each(Object.entries(CLEAN_728))("clean guard %s", (name, make) => {
        it("word-valid draws no such finding", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });
});

/**
 * Session 55's s55a specimens (s55builder.mts; #751): a header the body
 * references through `w:headerReference r:id="rIdH"`, stored as
 * `word/header1.xml` with NO relationships part, whose references name no
 * relationship. Word 16.115 offered recovery for all three shapes; its
 * calibrations — the same header with a relationships part that lacks the
 * id, and the header with no relationships part and no `r:` reference —
 * are the guards.
 */
const s55ChartDrawing = (ids: string): string =>
    '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Chart 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${C}">` +
    `<c:chart xmlns:c="${C}" ${ids}/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const s55DiagramDrawing = (ids: string): string =>
    '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Diagram 1"/><a:graphic><a:graphicData uri="${DGM}">` +
    `<dgm:relIds xmlns:dgm="${DGM}" ${ids}/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const s55OleObject = (ids: string): string =>
    "<w:p><w:r><w:object>" +
    `<o:OLEObject xmlns:o="${O}" Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" ${ids}/></w:object></w:r></w:p>`;

/**
 * A referenced `word/header1.xml` holding `body`, with `sidecar` as its relationships part — absent when undefined. The
 * image part ships only when the sidecar names it, as in s55a.
 */
const s55Header = (body: string, sidecar?: string): Record<string, string | Uint8Array> =>
    extended(
        {
            ...(sidecar?.includes("media/image1.png") ? { "word/media/image1.png": PNG } : {}),
            "word/header1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}" xmlns:dgm="${DGM}" xmlns:o="${O}" xmlns:r="${R}">${body}</w:hdr>`,
            ...(sidecar === undefined ? {} : { "word/_rels/header1.xml.rels": sidecar }),
            [MAIN]: replaceOnce(
                replaceOnce(MINIMAL_PARTS[MAIN]!, "<w:document ", `<w:document xmlns:r="${R}" `),
                "<w:sectPr>",
                '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>',
            ),
        },
        [relationship("rIdH", OFFICE("header"), "header1.xml")],
        [["word/header1.xml", HEADER_CT]],
        true,
    );

const REFUSED_751: Record<string, () => Promise<Buffer>> = {
    // s55a `header1-nosidecar-chart`: no relationships part, c:chart r:id="rId1" names none.
    "header1-nosidecar-chart": async () => buildDocx(s55Header(s55ChartDrawing('r:id="rId1"'))),
    // s55a `header1-nosidecar-oleobject`: the same holding o:OLEObject r:id="rId1".
    "header1-nosidecar-oleobject": async () => buildDocx(s55Header(s55OleObject('r:id="rId1"'))),
    // s55a `header1-nosidecar-relids`: the same holding dgm:relIds naming all four parts, so the schema holds.
    "header1-nosidecar-relids": async () => buildDocx(s55Header(s55DiagramDrawing('r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"'))),
    // s55a `header1-sidecar-no-rid1-calibration`: the relationships part is present but lacks rId1 — refused too (guard).
    "header1-sidecar-no-rid1-calibration": async () =>
        buildDocx(s55Header(s55ChartDrawing('r:id="rId1"'), relsPart(relationship("rId2", OFFICE("image"), "media/image1.png")))),
};
const CLEAN_751: Record<string, () => Promise<Buffer>> = {
    // s55a `header1-nosidecar-no-reference-calibration`: no relationships part and no r: reference opens clean.
    "header1-nosidecar-no-reference-calibration": async () => buildDocx(s55Header("<w:p><w:r><w:t>header</w:t></w:r></w:p>")),
};

describe(`${CODE} (session 55a no-sidecar specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    describe.each(Object.entries(REFUSED_751))("refused specimen %s", (name, make) => {
        it("word-valid fails with exactly one Word-blocking error, full and partial runs alike (#381)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/header1.xml");
                const partial = await validate(file, { checks: [CODE] });
                expect(partial.valid).toBe(false);
                expect(partial.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            });
        });

        it("strict reports the code too (strict ⊇ word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/header1.xml")).toBe(true);
            });
        });
    });

    describe.each(Object.entries(CLEAN_751))("clean specimen %s", (name, make) => {
        it("word-valid draws no such finding, full and partial runs alike", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                const partial = await validate(file, { checks: [CODE] });
                expect(partial.valid).toBe(true);
                expect(partial.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });
});

/**
 * Session 55's s55b/s55d specimens (#756): the minimal package's main
 * document declaring the ISO Strict relationships prefix `rs`, with a real
 * bar chart `rId9` names and `drawing` before `w:sectPr`. Word 16.115
 * cannot open a `c:chart` carrying an id attribute in both relationships
 * namespaces — an open error, no recovery offer — whatever the values,
 * while `rs:id` alone opens clean and draws the chart.
 */
const s55ChartIdsPackage = (drawing: string): Record<string, string | Uint8Array> => {
    const parts = extended(
        { [CHART_INSIDE]: barChart(false) },
        [relationship("rId9", OFFICE("chart"), "charts/chart1.xml")],
        [[CHART_INSIDE, CHART_CT]],
    );
    parts[MAIN] = replaceOnce(
        replaceOnce(
            MINIMAL_PARTS[MAIN]!,
            "<w:document ",
            `<w:document xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}" xmlns:rs="${RS}" `,
        ),
        "<w:sectPr>",
        `${drawing}<w:sectPr>`,
    );
    return parts;
};

const REFUSED_756: Record<string, () => Promise<Buffer>> = {
    // s55b `chart-empty-r-id-real-rs-id`: r:id="" rs:id="rId9".
    "chart-empty-r-id-real-rs-id": async () => buildDocx(s55ChartIdsPackage(s55ChartDrawing('r:id="" rs:id="rId9"'))),
    // s55b `chart-empty-rs-id-real-r-id`: rs:id="" r:id="rId9".
    "chart-empty-rs-id-real-r-id": async () => buildDocx(s55ChartIdsPackage(s55ChartDrawing('rs:id="" r:id="rId9"'))),
    // s55d `chart-both-ids-real`: r:id="rId9" rs:id="rId9" — both namespaces, both real.
    "chart-both-ids-real": async () => buildDocx(s55ChartIdsPackage(s55ChartDrawing('r:id="rId9" rs:id="rId9"'))),
};
const CLEAN_756: Record<string, () => Promise<Buffer>> = {
    // s55d `chart-rs-id-only`: rs:id="rId9" alone in the Transitional document opens clean, the chart draws.
    "chart-rs-id-only": async () => buildDocx(s55ChartIdsPackage(s55ChartDrawing('rs:id="rId9"'))),
    // s55b `chart-plain-calibration`: the plain r:id="rId9" names the real chart relationship (guard).
    "chart-plain-calibration": async () => buildDocx(s55ChartIdsPackage(s55ChartDrawing('r:id="rId9"'))),
};

describe(`${CODE} (session 55b/55d dual-namespace specimens, Word 16.115)`, { timeout: 240_000 }, () => {
    describe.each(Object.entries(REFUSED_756))("refused specimen %s", (name, make) => {
        it("word-valid fails with exactly one Word-blocking error, full and partial runs alike (#381)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                const found = result.issues.filter((issue) => issue.code === CODE);
                expect(result.valid).toBe(false);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("error");
                expect(found[0]?.path).toBe("word/document.xml");
                const partial = await validate(file, { checks: [CODE] });
                expect(partial.valid).toBe(false);
                expect(partial.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            });
        });

        it("strict reports the code too (strict ⊇ word-valid)", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
            });
        });
    });

    describe.each(Object.entries(CLEAN_756))("clean specimen %s", (name, make) => {
        it("word-valid draws no such finding, full and partial runs alike", async () => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `${name}.docx`);
                writeFileSync(file, await make());
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
                expect(result.issues.filter((issue) => issue.code === CODE)).toEqual([]);
                const partial = await validate(file, { checks: [CODE] });
                expect(partial.valid).toBe(true);
                expect(partial.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            });
        });
    });
});

describe(`${CODE} (selectors and the rels-id-missing hand-off)`, { timeout: 240_000 }, () => {
    it.each(["relationships", "drawings"])("a selected %s named-check run reports it", async (group) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "png-image-relationship-chart-reference.docx");
            writeFileSync(file, await REFUSED_TRANSITIONAL["png-image-relationship-chart-reference"]!());
            const partial = await validate(file, { checks: [group] });
            expect(partial.valid).toBe(false);
            expect(partial.issues.some((issue) => issue.code === CODE && issue.path === "word/document.xml")).toBe(true);
        });
    });

    it("reports a dangling Transitional r:id once in a full run, through the Word rule (#705)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "dangling-chart-reference.docx");
            // The same shape as the strict01 dangling rows, but Transitional:
            // a c:chart whose r:id names no Relationship in a present sidecar.
            writeFileSync(file, await buildDocx(withDrawing(pngReached(), CHART_DRAWING("rId9"))));
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            // The legacy walk's copy of the same element is handed off, as the
            // header/footer rule's is (withoutRepeatedHeaderFooterFindings).
            expect(wordValid.issues.filter((issue) => issue.code === "rels-id-missing")).toEqual([]);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            expect(strict.issues.filter((issue) => issue.code === "rels-id-missing")).toEqual([]);
            // Lenient runs no Word rule, so the legacy finding stands alone there.
            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.issues.filter((issue) => issue.code === CODE)).toEqual([]);
            expect(lenient.issues.some((issue) => issue.code === "rels-id-missing")).toBe(true);
        });
    });

    /**
     * #1074: outside lenient the legacy walk reads the Strict `rs:id` too
     * (#1065), so a session 55 dual-namespace chart whose ids differ yields
     * a legacy copy for each id the rule's one finding must drop; an `rs:id`
     * alone keeps #705's pairing.
     */
    it.each([
        ["word-valid", 'r:id="rId9" rs:id="rId8"'],
        ["strict", 'r:id="rId9" rs:id="rId8"'],
        ["word-valid", 'r:id="rId7" rs:id="rId8"'],
        ["strict", 'r:id="rId7" rs:id="rId8"'],
        ["word-valid", 'rs:id="rId8"'],
        ["strict", 'rs:id="rId8"'],
    ] as const)("reports the chart once under %s for %s (#1074)", async (profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "dual-namespace-chart.docx");
            writeFileSync(file, await buildDocx(s55ChartIdsPackage(s55ChartDrawing(ids))));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            expect(result.issues.filter((issue) => issue.code === "rels-id-missing")).toEqual([]);
        });
    });

    /**
     * #1079: the `o:OLEObject` twin of #1074's rows. The rule judges the
     * first non-empty id; the legacy copy of the other id pairs with its
     * finding too. `rId9` is the package's chart relationship, a mistyped
     * id for an OLE object.
     */
    it.each([
        ["word-valid", 'r:id="rId5" rs:id="rId6"'],
        ["strict", 'r:id="rId5" rs:id="rId6"'],
        ["word-valid", 'r:id="rId9" rs:id="rId6"'],
        ["strict", 'r:id="rId9" rs:id="rId6"'],
    ] as const)("reports the OLE object once under %s for %s (#1079)", async (profile, ids) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "dual-namespace-ole-object.docx");
            writeFileSync(file, await buildDocx(s55ChartIdsPackage(s55OleObject(ids))));
            const result = await validate(file, { profile });
            expect(result.issues.filter((issue) => issue.code === CODE)).toHaveLength(1);
            expect(result.issues.filter((issue) => issue.code === "rels-id-missing" || issue.code === "rels-id-mismatch")).toEqual([]);
        });
    });

    /**
     * #727: the element sits in both branches of one `mc:AlternateContent`
     * whose Choice Word does not select. The legacy walk reads the raw
     * part, so it sees two elements and makes two `rels-id-missing` copies;
     * the Word rule reads the selected view and reports one finding, which
     * must drop every legacy copy of the same Id on the same element.
     */
    const mcTwins = (element: string): Record<string, string | Uint8Array> => {
        const parts = withDrawing(
            pngReached(),
            `<mc:AlternateContent><mc:Choice Requires="x">${element}</mc:Choice><mc:Fallback>${element}</mc:Fallback></mc:AlternateContent>`,
        );
        parts[MAIN] = replaceOnce(parts[MAIN] as string, "<w:document ", '<w:document xmlns:x="http://example.com/unknown" ');
        return parts;
    };

    it.each([
        ["c:chart", () => mcTwins(CHART_DRAWING("rId9"))],
        ["o:OLEObject", () => mcTwins(OLE_OBJECT("rId9"))],
    ])("reports a dangling %s r:id once when its element has mc:Choice and mc:Fallback twins (#727)", async (element, make) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `twins-${element}.docx`);
            writeFileSync(file, await buildDocx(make()));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.filter((issue) => issue.code === CODE),
                    profile,
                ).toHaveLength(1);
                expect(
                    result.issues.filter((issue) => issue.code === "rels-id-missing"),
                    profile,
                ).toEqual([]);
            }
            // Lenient runs no Word rule; both twins keep their legacy copies.
            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.issues.filter((issue) => issue.code === "rels-id-missing")).toHaveLength(2);
        });
    });
});
