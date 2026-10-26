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
 *   without a part), or an Id no relationship has is refused.
 * - `dgm:relIds r:dm` must name a `diagramData` relationship; an `image`
 *   or `chart` relationship, or an Id no relationship has, is refused.
 *   `r:lo`, `r:qs` and `r:cs` are deliberately not checked: session 46
 *   opened them clean naming the image and the chart relationship, with
 *   the part's cached diagramDrawing rendered, so Word tolerates them.
 * - `o:OLEObject r:id` must name a `package` or an `oleObject`
 *   relationship. `package` is the type the Strict fixture itself carries;
 *   `oleObject`, internal or External (a linked object), is the spec's and
 *   was not shown to Word, so it is accepted unjudged. An `image` or
 *   `chart` relationship, or an Id no relationship has, is refused.
 *
 * Word reads the selected markup-compatibility view, so this check judges
 * `ctx.effectiveXml` and each part's own sidecar: an absent sidecar leaves
 * every id unresolved, and an unparseable one is skipped for xml-syntax to
 * report. The accepted Types are exact membership, byte for byte — the
 * Transitional `…/officeDocument/2006/relationships/<segment>` URI or its
 * ISO Strict `purl.oclc.org` counterpart — like the hyperlink rule's two
 * (#614), and the `r:` attribute may sit in either relationships
 * namespace. `a:blip r:embed` and `v:imagedata o:relid` are not judged:
 * session 45 opened both naming a relationship of another type clean.
 *
 * In a full word-valid or strict run, a dangling Transitional `r:id` on a
 * `c:chart` or an `o:OLEObject` would also be `rels-id-missing`'s finding;
 * the run keeps this rule's copy and drops that one
 * (`withoutRepeatedHeaderFooterFindings` in validators/docx.ts), as the
 * header/footer rule's hand-off does (#612, #619). `dgm:relIds r:dm` never
 * collides: the legacy walk reads only `r:id`, `r:embed` and `r:link`.
 */

import path from "node:path";
import type { CheckDefinition } from "../../../lib/check-registry";
import { equivalentPartLookup } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
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

/** One relationship of the part's own sidecar, as a reference reads it. */
type Relationship = { readonly type: string; readonly targetMode: string };

/**
 * Each relationship id the part's own sidecar names, with its Type and
 * TargetMode. An absent sidecar returns an empty map (every id unresolved);
 * an unparseable one returns `null` and is diagnosed by xml-syntax.
 */
const relationshipIndex = (
    ctx: ValidationContext,
    xmlFile: string,
    findPart: (partName: string) => string | undefined,
): Map<string, Relationship> | null => {
    const relsFile = findPart(path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`));
    const parsed = relsFile === undefined ? undefined : ctx.effectiveXml(relsFile);
    if (parsed && !parsed.ok) {
        return null;
    }
    const out = new Map<string, Relationship>();
    if (!parsed) {
        return out;
    }
    for (const rel of getRelationshipChildren(parsed.document)) {
        const id = rel.getAttribute("Id");
        if (!id) {
            continue;
        }
        out.set(id, { type: rel.getAttribute("Type") ?? "", targetMode: rel.getAttribute("TargetMode") ?? "" });
    }
    return out;
};

/** The id in the first relationships namespace the element carries `attribute` in; `null` for none or an empty one. */
const relationshipId = (elem: Element, attribute: string, namespaces: readonly string[]): { rid: string; namespaceURI: string } | null => {
    for (const relNs of namespaces) {
        const value = elem.getAttributeNS(relNs, attribute);
        if (value) {
            return { rid: value, namespaceURI: relNs };
        }
    }
    return null;
};

export const WORD_PART_REFERENCE_RELATIONSHIP = "word-part-reference-relationship";

export const wordPartReferenceRelationshipCheck: CheckDefinition = {
    id: WORD_PART_REFERENCE_RELATIONSHIP,
    description: "Chart, diagram and OLE object references name relationships of the type they read.",
    groups: ["relationships", "drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        const findPart = equivalentPartLookup(ctx.partNames);

        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, OFFICE_RELATIONSHIPS_STRICT_NAMESPACE];
        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = relationshipIndex(ctx, xmlFile, findPart);
            if (!rels) {
                continue;
            }
            for (const ns of CHART_NAMESPACES) {
                for (const elem of getElementsByTagNameNSAll(dom, ns, "chart")) {
                    const id = relationshipId(elem, "id", relationshipNamespaces);
                    if (!id) {
                        continue;
                    }
                    const rel = rels.get(id.rid);
                    // Session 45's External chart rows: the Type alone is not
                    // enough, the relationship must also be internal.
                    if (rel && ACCEPTED_CHART_RELATIONSHIP_TYPES.includes(rel.type) && rel.targetMode !== "External") {
                        continue;
                    }
                    const why = !rel
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
                    const rel = rels.get(id.rid);
                    if (rel && ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES.includes(rel.type)) {
                        continue;
                    }
                    const why = !rel
                        ? `no Relationship with Id '${id.rid}'`
                        : `Relationship '${id.rid}' has Type '${rel.type || "(none)"}'`;
                    issues.push({
                        severity: "error",
                        message:
                            `<dgm:relIds> r:dm="${id.rid}" does not reference a diagramData relationship (${why}). ` +
                            "Microsoft Word offers to recover the file. " +
                            `Use a Relationship of Type ${ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES[0]} ` +
                            `(or the ISO Strict ${ACCEPTED_DIAGRAM_DATA_RELATIONSHIP_TYPES[1]}), or point the r:dm at one.`,
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
                const rel = rels.get(id.rid);
                // A linked object names an External oleObject relationship;
                // Word was never shown one, so the TargetMode is not judged.
                if (rel && ACCEPTED_OLE_RELATIONSHIP_TYPES.includes(rel.type)) {
                    continue;
                }
                const why = !rel ? `no Relationship with Id '${id.rid}'` : `Relationship '${id.rid}' has Type '${rel.type || "(none)"}'`;
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
