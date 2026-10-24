/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A `w:headerReference` or `w:footerReference` in the main document's body
 * `w:sectPr`, or in a reached glossary's `w:docPartBody` `w:sectPr`, whose
 * `r:id` names a relationship whose Type does not end in `/header` or
 * `/footer` as the reference needs makes Word 16.115 offer to recover the
 * file (#606; session 26, chunks s26a/s26b/s26c, in
 * docs/plans/evidence/word-probes-2026-10-24.md). Word refused header→footer
 * and footer→header whether the target's root was `w:hdr` or `w:ftr` and
 * whatever its declared content type, refused both mirrors from a reached
 * glossary, and refused a reference to an unrelated type (`styles`); a
 * matching reference and relationship opens clean even when the target's
 * root element is the other story's. The relationship's type decides, so
 * this check judges the Type alone and never resolves the target.
 *
 * A reference whose `r:id` names no Relationship stays `rels-id-missing`'s
 * finding. In a full word-valid or strict run each element this check
 * reports yields one finding, this check's: the run drops the matching
 * `rels-id-mismatch` for it (#612; `withoutRepeatedHeaderFooterFindings` in
 * validators/docx.ts). Everywhere else `rels-id-mismatch` keeps the element —
 * a standalone `validateAllRelationshipIds()` call, a foreign-namespace
 * `x:headerReference` this check does not match, a glossary no
 * main-document relationship reaches, and lenient, which runs no Word rule
 * (Word does not read an unreached glossary, so word-valid stays silent
 * there).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName, equivalentPartLookup } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const WORD_HEADER_FOOTER_RELATIONSHIP = "word-header-footer-relationship";

const REFERENCES: readonly { readonly localName: "headerReference" | "footerReference"; readonly expected: "header" | "footer" }[] = [
    { localName: "headerReference", expected: "header" },
    { localName: "footerReference", expected: "footer" },
];

/**
 * Each relationship id the part's own sidecar names, with its Type. An
 * absent sidecar leaves every id unresolved (`rels-id-missing`'s business);
 * an unparseable one is diagnosed by xml-syntax.
 */
const relationshipTypes = (
    ctx: ValidationContext,
    xmlFile: string,
    findPart: (partName: string) => string | undefined,
): Map<string, string> | null => {
    const relsFile = findPart(path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`));
    const parsed = relsFile === undefined ? undefined : ctx.effectiveXml(relsFile);
    if (parsed && !parsed.ok) {
        return null;
    }
    const out = new Map<string, string>();
    if (!parsed) {
        return out;
    }
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
    run: async (ctx: ValidationContext): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        const findPart = equivalentPartLookup(ctx.partNames);

        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
        // The story roots session 26 observed: every main document, and every
        // glossary a main-document glossaryDocument relationship reaches.
        const roots = await wordTypedRelationshipTargets(ctx, new Set(["officeDocument", "glossaryDocument"]));
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
            if (types === null) {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const { localName, expected } of REFERENCES) {
                    for (const elem of getElementsByTagNameNSAll(parsed.document, ns, localName)) {
                        let rid: string | null = null;
                        for (const relNs of relationshipNamespaces) {
                            const value = elem.getAttributeNS(relNs, "id");
                            if (value) {
                                rid = value;
                                break;
                            }
                        }
                        // No r:id, or an r:id naming no Relationship: never
                        // this check's finding (rels-id-missing owns it).
                        const type = rid === null ? undefined : types.get(rid);
                        if (type === undefined) {
                            continue;
                        }
                        if (type.endsWith(`/${expected}`)) {
                            continue;
                        }
                        issues.push({
                            severity: "error",
                            message:
                                `<w:${localName}> r:id="${rid}" names Relationship '${rid}' of Type '${type || "(none)"}'; ` +
                                `a w:${localName} needs a Relationship whose Type ends in /${expected}. ` +
                                "Microsoft Word offers to recover the file. Point the r:id at such a Relationship, or retype that Relationship.",
                            path: root.part,
                            code: WORD_HEADER_FOOTER_RELATIONSHIP,
                        });
                    }
                }
            }
        }

        return issues;
    },
};
