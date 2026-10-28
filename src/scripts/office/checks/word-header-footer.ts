/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A `w:headerReference` or `w:footerReference` in the main document's body
 * `w:sectPr`, or in a reached glossary's `w:docPartBody` `w:sectPr`, whose
 * `r:id` names a relationship whose Type is not exactly the Transitional or
 * the ISO Strict header (respectively footer) relationship URI makes Word
 * 16.115 offer to recover the file (#606, #613; sessions 26 and 27 — chunks
 * s26a/s26b/s26c, s27a/s27b/s27c — in
 * docs/plans/evidence/word-probes-2026-10-24.md). Word refused header→footer
 * and footer→header whether the target's root was `w:hdr` or `w:ftr` and
 * whatever its declared content type, refused both mirrors from a reached
 * glossary, and refused a reference to an unrelated type (`styles`); a
 * matching reference and relationship opens clean even when the target's
 * root element is the other story's. Session 27 pinned the accepted Types to
 * exactly two per element, byte for byte: the Transitional and ISO Strict
 * URIs below open clean, while `http://example.com/header`,
 * `…/relationships/custom/header`, `…/relationships/HEADER`, the URI padded
 * with spaces, the URI wrapped in tab and newline character references, and
 * the bare `header` are all refused (the footer pair likewise). The
 * relationship's type decides, so this check judges the Type alone and never
 * resolves the target.
 *
 * A reference whose `r:id` names no Relationship in the part's own
 * relationships part makes Word 16.115 offer recovery too, once that
 * relationships part exists and parses (#619; session 32): a `w:footerReference`
 * `r:id="rId9"` in the body `w:sectPr` beside a kept
 * `word/_rels/document.xml.rels` holding no `rId9` was refused. A part with
 * no relationships part keeps a non-empty `r:id` as
 * `rels-missing-sidecar`'s finding alone; an unparseable one is diagnosed by
 * xml-syntax for lookups.
 *
 * An empty or missing id in both relationships namespaces is this check's
 * finding even without a parseable sidecar (#623; session 33, chunks
 * s33a/s33c/s33d/s33e in the same evidence file). Word 16.115 offered
 * recovery for empty `r:id` on header and footer references in the main
 * document's body `w:sectPr`, and reported an error opening the file for
 * either reference with no id. A footer reference whose only id was in a
 * third namespace got the same open error, whether dangling or naming a
 * real footer relationship. A non-empty ISO Strict relationships id on a
 * Transitional reference naming a real footer relationship opened clean
 * and rendered the footer (#631 refuted). Other placements and reached
 * glossaries are on the same footing as sessions 26/27; session 33 observed
 * the empty/missing shapes only in the main document's body `w:sectPr`.
 * These findings register no pairing identity: the legacy walk reads only
 * truthy ids and reports neither shape.
 *
 * In a full word-valid or strict
 * run each element this check reports yields one relationship finding, this
 * check's: the run drops the matching `rels-id-mismatch` — or, for an `r:id` naming no
 * Relationship, `rels-id-missing` — copy of it (#612, #619;
 * `withoutRepeatedHeaderFooterFindings` in validators/docx.ts), paired through
 * the structured element identity the check registers on the issue, never
 * message prose (#616). Everywhere else those codes keep the element —
 * a standalone `validateAllRelationshipIds()` call, a foreign-namespace
 * `x:headerReference` this check does not match, a glossary no
 * main-document relationship reaches, and lenient, which runs no Word rule
 * (Word does not read an unreached glossary, so word-valid stays silent
 * there). Strict also keeps the schema's own `xsd-error` for a reference with
 * no `r:id`, which the XSD requires; it carries no identity to pair.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName, equivalentPartLookup } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";
import { otherRelationshipIds, registerStoryReferenceIdentity, relationshipAttributeName } from "../story-reference-identity";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const WORD_HEADER_FOOTER_RELATIONSHIP = "word-header-footer-relationship";

const REFERENCES: readonly { readonly localName: "headerReference" | "footerReference"; readonly expected: "header" | "footer" }[] = [
    { localName: "headerReference", expected: "header" },
    { localName: "footerReference", expected: "footer" },
];

/**
 * Session 27 (#613): the only relationship Types Word 16.115 accepts for each
 * reference, byte for byte — the Transitional URI first, then its ISO Strict
 * counterpart. Every other Type is refused, `…/custom/header`,
 * `http://example.com/header` and `…/HEADER` included.
 */
const ACCEPTED_RELATIONSHIP_TYPES: Readonly<Record<"header" | "footer", readonly [string, string]>> = {
    header: [`${OFFICE_RELATIONSHIPS_NAMESPACE}/header`, "http://purl.oclc.org/ooxml/officeDocument/relationships/header"],
    footer: [`${OFFICE_RELATIONSHIPS_NAMESPACE}/footer`, "http://purl.oclc.org/ooxml/officeDocument/relationships/footer"],
};

/**
 * Each relationship id the part's own sidecar names, with its Type. An
 * absent sidecar returns `undefined` (`rels-missing-sidecar`'s business); an
 * unparseable one returns `null` and is diagnosed by xml-syntax; a present,
 * parseable one returns its Ids, possibly none, so an `r:id` naming none of
 * them is this check's finding (#619).
 */
const relationshipTypes = (
    ctx: ValidationContext,
    xmlFile: string,
    findPart: (partName: string) => string | undefined,
): Map<string, string> | null | undefined => {
    const relsFile = findPart(path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`));
    if (relsFile === undefined) {
        return undefined;
    }
    const parsed = ctx.effectiveXml(relsFile);
    if (!parsed?.ok) {
        return null;
    }
    const out = new Map<string, string>();
    for (const rel of getRelationshipChildren(parsed.document)) {
        const id = rel.getAttribute("Id");
        if (!id) {
            continue;
        }
        out.set(id, rel.getAttribute("Type") ?? "");
    }
    return out;
};

export const wordHeaderFooterRelationshipCheck: CheckDefinition = {
    id: WORD_HEADER_FOOTER_RELATIONSHIP,
    description: "Header and footer references name header and footer relationships.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    /**
     * One error per story `w:headerReference`/`w:footerReference` whose
     * `r:id` is empty or missing, independently of the sidecar (#623), or
     * names no Relationship in its part's present, parseable
     * relationships part, or names one whose Type is not the accepted
     * Transitional or ISO Strict URI. Only findings for non-empty ids carry
     * the element identity the full run's hand-off pairs on.
     */
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        const findPart = equivalentPartLookup(ctx.partNames);

        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
        // The story roots session 26 observed: every main document, and every
        // glossary a main-document glossaryDocument relationship reaches.
        const roots = await wordTypedRelationshipTargets(ctx, new Set(["officeDocument", "glossaryDocument"]), profile);
        const seen = new Set<string>();
        for (const root of roots) {
            const partName = canonicalPartName(root.part);
            if (seen.has(partName)) {
                continue;
            }
            seen.add(partName);
            const parsed = ctx.effectiveXml(root.part);
            if (!parsed?.ok) {
                continue;
            }
            const types = relationshipTypes(ctx, root.part, findPart);
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const { localName, expected } of REFERENCES) {
                    const [transitionalType, strictType] = ACCEPTED_RELATIONSHIP_TYPES[expected];
                    for (const elem of getElementsByTagNameNSAll(parsed.document, ns, localName)) {
                        let rid: string | null = null;
                        let ridNamespaceURI = "";
                        let emptyIdNamespaceURI = "";
                        for (const relNs of relationshipNamespaces) {
                            const value = elem.getAttributeNS(relNs, "id");
                            if (value) {
                                rid = value;
                                ridNamespaceURI = relNs;
                                break;
                            }
                            if (!emptyIdNamespaceURI && elem.hasAttributeNS(relNs, "id")) emptyIdNamespaceURI = relNs;
                        }
                        // #623: the attribute is wrong before any lookup.
                        // The legacy walk skips empty/missing ids, so these
                        // findings have no copy to pair by identity.
                        if (rid === null) {
                            const hasEmptyId = emptyIdNamespaceURI !== "";
                            // #1109: an empty id is named as the part spells
                            // it, the first read; a missing one keeps r:id.
                            const attribute = relationshipAttributeName(elem, emptyIdNamespaceURI);
                            issues.push({
                                severity: "error",
                                message:
                                    `<w:${localName}> has ${hasEmptyId ? "an empty" : "a missing"} ${attribute}; ` +
                                    `a w:${localName} needs a Relationship of Type ${transitionalType} (or the ISO Strict ${strictType}). ` +
                                    (hasEmptyId ? "Microsoft Word offers to recover the file. " : "Microsoft Word cannot open the file. ") +
                                    `Set a non-empty ${attribute} naming such a Relationship, adding one with that Id and Type when the part has none.`,
                                path: root.part,
                                code: WORD_HEADER_FOOTER_RELATIONSHIP,
                            });
                            continue;
                        }
                        if (types === null || types === undefined) continue;
                        const type = types.get(rid);
                        const attribute = relationshipAttributeName(elem, ridNamespaceURI);
                        // An r:id naming no Relationship of a part whose
                        // sidecar exists and parses: this check's finding
                        // (#619). A part with no sidecar is
                        // rels-missing-sidecar's alone.
                        if (type === undefined) {
                            const issue: ValidationIssue = {
                                severity: "error",
                                message:
                                    `<w:${localName}> ${attribute}="${rid}" names no Relationship with Id '${rid}'; ` +
                                    `a w:${localName} needs a Relationship of Type ${transitionalType} (or the ISO Strict ${strictType}). ` +
                                    `Microsoft Word offers to recover the file. Point the ${attribute} at such a Relationship, or add one with that Id and Type.`,
                                path: root.part,
                                code: WORD_HEADER_FOOTER_RELATIONSHIP,
                            };
                            // #619: the full run's hand-off pairs this finding
                            // with rels-id-missing's copy through the
                            // registered element identity, never the prose.
                            registerStoryReferenceIdentity(issue, {
                                part: root.part,
                                namespaceURI: ns,
                                localName,
                                rid,
                                ridNamespaceURI,
                                otherRids: otherRelationshipIds(elem, relationshipNamespaces, rid),
                            });
                            issues.push(issue);
                            continue;
                        }
                        if (type === transitionalType || type === strictType) {
                            continue;
                        }
                        const issue: ValidationIssue = {
                            severity: "error",
                            message:
                                `<w:${localName}> ${attribute}="${rid}" names Relationship '${rid}' of Type '${type || "(none)"}'; ` +
                                `a w:${localName} needs a Relationship of Type ${transitionalType} (or the ISO Strict ${strictType}). ` +
                                `Microsoft Word offers to recover the file. Point the ${attribute} at such a Relationship, or retype that Relationship.`,
                            path: root.part,
                            code: WORD_HEADER_FOOTER_RELATIONSHIP,
                        };
                        // #616: the full run's hand-off pairs this finding with
                        // rels-id-mismatch's copy through the registered element
                        // identity, never the message prose.
                        registerStoryReferenceIdentity(issue, {
                            part: root.part,
                            namespaceURI: ns,
                            localName,
                            rid,
                            ridNamespaceURI,
                            otherRids: otherRelationshipIds(elem, relationshipNamespaces, rid),
                        });
                        issues.push(issue);
                    }
                }
            }
        }

        return issues;
    },
};
