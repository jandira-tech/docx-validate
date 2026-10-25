/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A `w:hyperlink` whose `r:id` does not name an external hyperlink
 * relationship makes Word offer to recover the file. One such element is
 * enough. Removing the `r:id`, or adding a Relationship with the Transitional
 * or the ISO Strict hyperlink Type and TargetMode `External`, opens clean. A
 * hyperlink relationship without TargetMode, and an external relationship of
 * another type, still warn. A hyperlink with no `r:id` is left alone.
 *
 * Word session 28 (#614; chunk s28a in
 * docs/plans/evidence/word-probes-2026-10-24.md) pinned the accepted Types to
 * exactly two, byte for byte: the Transitional and ISO Strict hyperlink URIs
 * below open clean, while `http://example.com/hyperlink`,
 * `…/relationships/custom/hyperlink` and `…/relationships/HYPERLINK` are
 * refused. The old `endsWith("/hyperlink")` test passed the first two of
 * those, so the check now requires exact membership in the two accepted URIs.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { equivalentPartLookup } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import type { ValidationContext } from "../../../lib/validation-context";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Session 28 (#614): the only relationship Types Word 16.115 accepts for a
 * `w:hyperlink`, byte for byte — the Transitional URI first, then its ISO
 * Strict counterpart. Every other Type is refused, `…/custom/hyperlink`,
 * `http://example.com/hyperlink` and `…/HYPERLINK` included.
 */
const ACCEPTED_HYPERLINK_RELATIONSHIP_TYPES: readonly [string, string] = [
    `${OFFICE_RELATIONSHIPS_NAMESPACE}/hyperlink`,
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink",
];

const isAcceptedHyperlinkRelationshipType = (type: string): boolean => ACCEPTED_HYPERLINK_RELATIONSHIP_TYPES.includes(type);

/** An absent sidecar leaves every id unresolved; an unparseable sidecar is diagnosed by xml-syntax. */
const externalHyperlinkRelationshipIds = (
    ctx: ValidationContext,
    xmlFile: string,
    findPart: (partName: string) => string | undefined,
): Map<string, { type: string; targetMode: string; externalHyperlink: boolean }> | null => {
    const relsFile = findPart(path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`));
    const parsed = relsFile === undefined ? undefined : ctx.effectiveXml(relsFile);
    if (parsed && !parsed.ok) {
        return null;
    }
    const out = new Map<string, { type: string; targetMode: string; externalHyperlink: boolean }>();
    if (!parsed) {
        return out;
    }
    for (const rel of getRelationshipChildren(parsed.document)) {
        const id = rel.getAttribute("Id");
        if (!id) {
            continue;
        }
        const type = rel.getAttribute("Type") ?? "";
        const targetMode = rel.getAttribute("TargetMode") ?? "";
        out.set(id, { type, targetMode, externalHyperlink: isAcceptedHyperlinkRelationshipType(type) && targetMode === "External" });
    }
    return out;
};

export const wordHyperlinkRelationshipCheck: CheckDefinition = {
    id: "word-hyperlink-relationship",
    description: "Hyperlink ids name external hyperlink relationships.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        const findPart = equivalentPartLookup(ctx.partNames);

        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = externalHyperlinkRelationshipIds(ctx, xmlFile, findPart);
            if (!rels) {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const links = getElementsByTagNameNSAll(dom, ns, "hyperlink");
                for (const link of links) {
                    let rid: string | null = null;
                    for (const relNs of relationshipNamespaces) {
                        const value = link.getAttributeNS(relNs, "id");
                        if (value) {
                            rid = value;
                            break;
                        }
                    }
                    if (!rid) {
                        continue;
                    }
                    const rel = rels.get(rid);
                    if (rel?.externalHyperlink) {
                        continue;
                    }
                    let why: string;
                    if (!rel) {
                        why = `no Relationship with Id '${rid}'`;
                    } else if (!isAcceptedHyperlinkRelationshipType(rel.type)) {
                        why = `Relationship '${rid}' has Type '${rel.type || "(none)"}'`;
                    } else {
                        why = `Relationship '${rid}' has TargetMode '${rel.targetMode || "(none)"}'`;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `<w:hyperlink> r:id="${rid}" does not reference an external hyperlink relationship (${why}). ` +
                            "Microsoft Word offers to recover the file. " +
                            `Use a Relationship of Type ${ACCEPTED_HYPERLINK_RELATIONSHIP_TYPES[0]} ` +
                            `(or the ISO Strict ${ACCEPTED_HYPERLINK_RELATIONSHIP_TYPES[1]}) with TargetMode External, or remove the r:id.`,
                        path: xmlFile,
                        code: "word-hyperlink-relationship",
                    });
                }
            }
        }

        return issues;
    },
};
