/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A graphic-frame or OLE reference in a story part whose relationship id
 * names a relationship of another Type than the element reads makes Word
 * 16.115 offer to recover the file. One such element is enough, and the
 * target part does not matter: a real PNG, a real chart with no sidecar and
 * no part at all are all refused. Word session 45 (#705; chunks s45b–s45k
 * in docs/plans/evidence/word-probes-2026-10-24.md) refused these on the
 * minimal Transitional package, and session 46 (s46a–s46c) refused the same
 * shapes as one-attribute mutations of the Word-clean ISO Strict
 * `tests/fixtures/working/word-clean-strict01.docx`, which opens clean
 * re-zipped unchanged:
 *
 * - `c:chart r:id` must name an internal `chart` relationship. An `image`
 *   relationship (to a real PNG or to a real chart part), a `diagramData`
 *   or `package` relationship, an External `chart` relationship (with or
 *   without a part), an External `hyperlink` relationship, or an Id no
 *   relationship has is refused.
 * - `dgm:relIds r:dm` must name an internal `diagramData` relationship; an
 *   `image` or `chart` relationship, an External `diagramData`
 *   relationship, or an Id no relationship has, is refused.
 *   `r:lo`, `r:qs` and `r:cs` are deliberately not checked: session 46
 *   opened them clean naming the image and the chart relationship, with
 *   the part's cached diagramDrawing rendered, so Word tolerates them.
 * - `o:OLEObject r:id` must name a `package` or an `oleObject`
 *   relationship. `package` is the type the Strict fixture itself carries,
 *   internal or External — session 52's B1 row opened an External
 *   `package` relationship (a linked package) clean; `oleObject`, internal
 *   or External (a linked object), is the spec's and was not shown to
 *   Word, so it is accepted unjudged. An `image` or `chart` relationship,
 *   or an Id no relationship has, is refused.
 *
 * Session 52 (#728; chunks s52a–s52d in the same evidence file, two runs,
 * all readings agree) filled the gaps those bullets left: Word refused
 * every Type shape tried under a reference it reads — an `o:OLEObject`
 * naming a `diagramData` relationship (A1), a `c:chart` naming an External
 * `hyperlink` relationship (A2), and `…/relationships/chart/` (A3) and
 * `…/relationships/Chart` (A4) naming a real chart that the exact Type
 * opens clean — so the exact-membership rule stands on Word rows. It also
 * refused an empty id present on the element — the recovery offer for
 * `c:chart r:id=""` (C1) and `o:OLEObject r:id=""` (C2), an open error
 * for `dgm:relIds r:dm=""` on the review's `r:dm`-only element and the
 * recovery offer for it on the SmartArt whose `r:lo`, `r:qs` and `r:cs`
 * stay intact (C3) — and an External `diagramData` relationship a
 * `dgm:relIds r:dm` names, in both shapes (B2). An id the element does
 * not carry at all is still not judged, and `dgm:relIds r:dm="rId35"`
 * alone, naming the real relationship, opens clean (the s52d confound),
 * as an `o:OLEObject` naming an External `package` relationship does (B1).
 *
 * A `c:chart` carrying an id attribute in BOTH relationships namespaces —
 * `r:id` and the ISO Strict `rs:id` together, whatever their values,
 * empty or real — cannot be opened: Word 16.115 reports an open error
 * with no recovery offer (session 55, chunks s55b and s55d, two runs, all
 * readings agree; #756), while `rs:id` alone on the Transitional element
 * opens clean and draws the chart. Word refuses the second namespace,
 * not the empty id. The same pair on `o:OLEObject`, `dgm:relIds` and
 * `w:headerReference`/`w:footerReference` was not shown to Word —
 * session 33 showed only a non-empty Strict id alone on a Transitional
 * header reference opening clean — so those elements keep their verdicts.
 *
 * Word reads the selected markup-compatibility view, so this check judges
 * `ctx.effectiveXml` and each part's own sidecar: an unparseable one is
 * skipped for xml-syntax to report, and a part with no relationships part
 * at all is judged like one whose relationships part lacks the id — every
 * non-empty id it holds names no relationship, and an empty id is wrong
 * before any lookup (#728) — while `rels-missing-sidecar` still reports
 * the absent sidecar itself (#751; session 55, chunk s55a, two runs: the
 * recovery offer for a referenced `word/header1.xml` with no
 * relationships part holding `c:chart r:id`, `o:OLEObject r:id` or
 * `dgm:relIds`, as for the same header whose relationships part lacks the
 * id, and a clean open for the same header with no `r:` reference). The
 * accepted Types are exact membership, byte for byte — the
 * Transitional `…/officeDocument/2006/relationships/<segment>` URI or its
 * ISO Strict `purl.oclc.org` counterpart — like the hyperlink rule's two
 * (#614); session 52 refused every other Type shape tried under these
 * references (A1–A4), and the `r:` attribute may sit in either
 * relationships namespace. `a:blip r:embed` and `v:imagedata o:relid` are not judged:
 * session 45 opened both naming a relationship of another type clean.
 *
 * The parts judged are the parts Word reads (#726; session 51, chunk s51a,
 * two runs), found by relationship as `wordReadTypedTargets` models them:
 * the main document and every reached glossary wherever they are stored
 * (`word/headerFirst.xml` and `word/main.xml` get the recovery offer), plus
 * the stories their references make Word read — headers and footers a
 * section reference names, the comments, footnotes and endnotes parts the
 * document uses. A part Word does not read is not judged however its name
 * looks: a header the main document relates but no `w:headerReference`
 * names, with its sidecar and without one, and a glossary no relationship
 * reaches, all open clean (session 51; as sessions 15 and 35 predicted).
 * A partial `--checks word-part-reference-relationship` run walks the same
 * read set, so its verdict is the full run's (#381).
 *
 * In a full word-valid or strict run, a dangling Transitional `r:id` on a
 * `c:chart` or an `o:OLEObject` would also be `rels-id-missing`'s finding;
 * the run keeps this rule's copy and drops that one
 * (`withoutRepeatedHeaderFooterFindings` in validators/docx.ts), as the
 * header/footer rule's hand-off does (#612, #619), dropping every copy of
 * the same Id on the same element, twins included (#727). `dgm:relIds
 * r:dm` never collides: the legacy walk reads only `r:id`, `r:embed` and
 * `r:link`, and an empty id is below its truthy read too (#728).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { sidecarOf, wordReadTypedTargets } from "../../../lib/word-read-parts";
import { registerStoryReferenceIdentity } from "../story-reference-identity";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE_RELATIONSHIPS_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/** Chart namespaces, Transitional first, ISO Strict second. */
const CHART_NAMESPACES: readonly [string, string] = [
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://purl.oclc.org/ooxml/drawingml/chart",
];
/** Diagram namespaces, Transitional first, ISO Strict second. */
const DIAGRAM_NAMESPACES: readonly [string, string] = [
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://purl.oclc.org/ooxml/drawingml/diagram",
];
/** The VML office namespace `o:OLEObject` lives in, shared by both conformance classes. */
const OFFICE_NAMESPACE = "urn:schemas-microsoft-com:office:office";

const relationshipTypes = (segment: string): readonly [string, string] => [
    `${OFFICE_RELATIONSHIPS_NAMESPACE}/${segment}`,
    `${OFFICE_RELATIONSHIPS_STRICT_NAMESPACE}/${segment}`,
];

/**
 * Sessions 45 and 46 (#705): the only relationship Types Word 16.115
 * accepts for each reference, byte for byte — the Transitional URI first,
 * then its ISO Strict counterpart.
 */
const ACCEPTED_CHART_RELATIONSHIP_TYPES = relationshipTypes("chart");
const ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES = relationshipTypes("diagramData");
const ACCEPTED_OLE_RELATIONSHIP_TYPES: readonly string[] = [...relationshipTypes("package"), ...relationshipTypes("oleObject")];

/**
 * The story segments whose parts Word reads because a story root's
 * references make it read them (#726): used headers, footers, comments,
 * footnotes and endnotes. A story root itself is typed `officeDocument`.
 */
const STORY_SEGMENTS: ReadonlySet<string> = new Set(["header", "footer", "comments", "footnotes", "endnotes"]);

/** One relationship of the part's own sidecar, as a reference reads it. */
type Relationship = { readonly type: string; readonly targetMode: string };

/**
 * Each relationship id the part's own sidecar names, with its Type and
 * TargetMode. An absent sidecar returns `undefined` — every non-empty id
 * then names no relationship (#751); an unparseable one returns `null`
 * and is diagnosed by xml-syntax.
 */
const relationshipIndex = (ctx: ValidationContext, xmlFile: string): Map<string, Relationship> | null | undefined => {
    const relsFile = sidecarOf(ctx.partNames, xmlFile);
    if (relsFile === undefined) {
        return undefined;
    }
    const parsed = ctx.effectiveXml(relsFile);
    if (!parsed?.ok) {
        return null;
    }
    const out = new Map<string, Relationship>();
    for (const rel of getRelationshipChildren(parsed.document)) {
        const id = rel.getAttribute("Id");
        if (!id) {
            continue;
        }
        out.set(id, { type: rel.getAttribute("Type") ?? "", targetMode: rel.getAttribute("TargetMode") ?? "" });
    }
    return out;
};

/**
 * The id the element carries `attribute` in, from the first relationships
 * namespace that has it: `null` only when neither namespace has the
 * attribute — an empty value present on the element is returned as an
 * empty id and judged (#728), with a non-empty id in either namespace
 * winning over an empty one, as the header/footer rule reads ids.
 */
const relationshipId = (elem: Element, attribute: string, namespaces: readonly string[]): { rid: string; namespaceURI: string } | null => {
    let empty: { rid: string; namespaceURI: string } | null = null;
    for (const relNs of namespaces) {
        if (!elem.hasAttributeNS(relNs, attribute)) {
            continue;
        }
        const value = elem.getAttributeNS(relNs, attribute);
        if (value) {
            return { rid: value, namespaceURI: relNs };
        }
        empty ??= { rid: "", namespaceURI: relNs };
    }
    return empty;
};

export const WORD_PART_REFERENCE_RELATIONSHIP = "word-part-reference-relationship";

export const wordPartReferenceRelationshipCheck: CheckDefinition = {
    id: WORD_PART_REFERENCE_RELATIONSHIP,
    description: "Chart, diagram and OLE object references name relationships of the type they read.",
    groups: ["relationships", "drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    /**
     * #726 (session 51): judge the parts Word reads — every main document
     * and reached glossary wherever stored, plus the stories their
     * references make Word read — not the parts whose name looks like a
     * story. `wordReadTypedTargets` gates the lazy segments by use, so an
     * unused header and an unreached glossary are absent from the walk.
     */
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, OFFICE_RELATIONSHIPS_STRICT_NAMESPACE];
        const seen = new Set<string>();
        for (const target of await wordReadTypedTargets(ctx, profile)) {
            if (target.segment !== "officeDocument" && !STORY_SEGMENTS.has(target.segment)) {
                continue;
            }
            const xmlFile = target.part;
            // A part read through two segments is judged once, through
            // whichever reached it first, as the header/footer rule judges
            // a root two types reach once.
            const canonical = canonicalPartName(xmlFile);
            if (seen.has(canonical)) {
                continue;
            }
            seen.add(canonical);
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = relationshipIndex(ctx, xmlFile);
            // An unparseable sidecar is xml-syntax's to report.
            if (rels === null) {
                continue;
            }
            for (const ns of CHART_NAMESPACES) {
                for (const elem of getElementsByTagNameNSAll(dom, ns, "chart")) {
                    const id = relationshipId(elem, "id", relationshipNamespaces);
                    if (!id) {
                        continue;
                    }
                    // Session 55 (s55b/s55d, #756): an id attribute in BOTH
                    // relationships namespaces is an open error whatever
                    // the values — the element is refused before its id is
                    // looked up.
                    if (
                        elem.hasAttributeNS(OFFICE_RELATIONSHIPS_NAMESPACE, "id") &&
                        elem.hasAttributeNS(OFFICE_RELATIONSHIPS_STRICT_NAMESPACE, "id")
                    ) {
                        const issue: ValidationIssue = {
                            severity: "error",
                            message:
                                `<c:chart> carries an id attribute in both relationships namespaces (r:id="${elem.getAttributeNS(OFFICE_RELATIONSHIPS_NAMESPACE, "id")}" rs:id="${elem.getAttributeNS(OFFICE_RELATIONSHIPS_STRICT_NAMESPACE, "id")}"). ` +
                                "Microsoft Word cannot open the file. " +
                                "Keep only one id attribute, naming an internal chart relationship.",
                            path: xmlFile,
                            code: WORD_PART_REFERENCE_RELATIONSHIP,
                        };
                        registerStoryReferenceIdentity(issue, {
                            part: xmlFile,
                            namespaceURI: ns,
                            localName: "chart",
                            rid: id.rid,
                            ridNamespaceURI: id.namespaceURI,
                        });
                        issues.push(issue);
                        continue;
                    }
                    const rel = rels?.get(id.rid);
                    // Session 45's External chart rows: the Type alone is not
                    // enough, the relationship must also be internal.
                    if (rel && ACCEPTED_CHART_RELATIONSHIP_TYPES.includes(rel.type) && rel.targetMode !== "External") {
                        continue;
                    }
                    const why =
                        id.rid === ""
                            ? "the id is empty"
                            : !rel
                              ? `no Relationship with Id '${id.rid}'`
                              : ACCEPTED_CHART_RELATIONSHIP_TYPES.includes(rel.type)
                                ? `Relationship '${id.rid}' has TargetMode '${rel.targetMode || "(none)"}'`
                                : `Relationship '${id.rid}' has Type '${rel.type || "(none)"}'`;
                    const issue: ValidationIssue = {
                        severity: "error",
                        message:
                            `<c:chart> r:id="${id.rid}" does not reference an internal chart relationship (${why}). ` +
                            "Microsoft Word offers to recover the file. " +
                            `Use a Relationship of Type ${ACCEPTED_CHART_RELATIONSHIP_TYPES[0]} ` +
                            `(or the ISO Strict ${ACCEPTED_CHART_RELATIONSHIP_TYPES[1]}) internal to the package, or point the r:id at one.`,
                        path: xmlFile,
                        code: WORD_PART_REFERENCE_RELATIONSHIP,
                    };
                    // #705: a dangling Transitional r:id here is also the
                    // legacy walk's rels-id-missing; the full run's hand-off
                    // pairs the two through this identity, never the prose.
                    registerStoryReferenceIdentity(issue, {
                        part: xmlFile,
                        namespaceURI: ns,
                        localName: "chart",
                        rid: id.rid,
                        ridNamespaceURI: id.namespaceURI,
                    });
                    issues.push(issue);
                }
            }
            for (const ns of DIAGRAM_NAMESPACES) {
                for (const elem of getElementsByTagNameNSAll(dom, ns, "relIds")) {
                    // Only r:dm: session 46 opened r:lo, r:qs and r:cs naming
                    // the image and the chart relationship clean.
                    const id = relationshipId(elem, "dm", relationshipNamespaces);
                    if (!id) {
                        continue;
                    }
                    const rel = rels?.get(id.rid);
                    // Session 52's B2 rows: like the chart rule, the Type
                    // alone is not enough — an External diagramData
                    // relationship is refused too.
                    if (rel && ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES.includes(rel.type) && rel.targetMode !== "External") {
                        continue;
                    }
                    const why =
                        id.rid === ""
                            ? "the id is empty"
                            : !rel
                              ? `no Relationship with Id '${id.rid}'`
                              : ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES.includes(rel.type)
                                ? `Relationship '${id.rid}' has TargetMode '${rel.targetMode || "(none)"}'`
                                : `Relationship '${id.rid}' has Type '${rel.type || "(none)"}'`;
                    issues.push({
                        severity: "error",
                        message:
                            `<dgm:relIds> r:dm="${id.rid}" does not reference an internal diagramData relationship (${why}). ` +
                            // Session 52: an open error for the empty r:dm
                            // on the minimal shape (s52b), the recovery
                            // offer for it on the strict01 SmartArt (s52d);
                            // the refused non-empty rows all drew the offer.
                            (id.rid === ""
                                ? "Microsoft Word reported an open error on one specimen and offered to recover another. "
                                : "Microsoft Word offers to recover the file. ") +
                            `Use a Relationship of Type ${ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES[0]} ` +
                            `(or the ISO Strict ${ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES[1]}) internal to the package, or point the r:dm at one.`,
                        path: xmlFile,
                        code: WORD_PART_REFERENCE_RELATIONSHIP,
                    });
                }
            }
            for (const elem of getElementsByTagNameNSAll(dom, OFFICE_NAMESPACE, "OLEObject")) {
                const id = relationshipId(elem, "id", relationshipNamespaces);
                if (!id) {
                    continue;
                }
                const rel = rels?.get(id.rid);
                // A linked object names an External oleObject relationship
                // and session 52's B1 opened an External `package` one clean,
                // so the TargetMode is not judged.
                if (rel && ACCEPTED_OLE_RELATIONSHIP_TYPES.includes(rel.type)) {
                    continue;
                }
                const why =
                    id.rid === ""
                        ? "the id is empty"
                        : !rel
                          ? `no Relationship with Id '${id.rid}'`
                          : `Relationship '${id.rid}' has Type '${rel.type || "(none)"}'`;
                const issue: ValidationIssue = {
                    severity: "error",
                    message:
                        `<o:OLEObject> r:id="${id.rid}" does not reference an OLE object relationship (${why}). ` +
                        "Microsoft Word offers to recover the file. " +
                        `Use a Relationship of Type ${OFFICE_RELATIONSHIPS_NAMESPACE}/package or ${OFFICE_RELATIONSHIPS_NAMESPACE}/oleObject ` +
                        `(or the ISO Strict ${OFFICE_RELATIONSHIPS_STRICT_NAMESPACE}/package and ${OFFICE_RELATIONSHIPS_STRICT_NAMESPACE}/oleObject), or point the r:id at one.`,
                    path: xmlFile,
                    code: WORD_PART_REFERENCE_RELATIONSHIP,
                };
                registerStoryReferenceIdentity(issue, {
                    part: xmlFile,
                    namespaceURI: OFFICE_NAMESPACE,
                    localName: "OLEObject",
                    rid: id.rid,
                    ridNamespaceURI: id.namespaceURI,
                });
                issues.push(issue);
            }
        }

        return issues;
    },
};
