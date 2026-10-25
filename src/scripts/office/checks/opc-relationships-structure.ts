/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName, isRelationshipsPartName } from "../../../lib/content-types";
import { applyMarkupCompatibility, understoodNamespacesForPart } from "../../../lib/markup-compatibility";
import { NS, type ValidationIssue } from "../../../lib/types";
import { getRelationshipChildren } from "../../../lib/xml-helpers";

/**
 * The package relationships part, as OPC compares part names: `_RELS/.RELS`
 * and `_rels/.%72els` name it too, as every canonical-part comparison does
 * (#364).
 */
const isPackageRelationshipsPart = (part: string): boolean => canonicalPartName(part) === "_rels/.rels";

const unexpectedChildren = (document: Document): readonly Element[] => {
    const relationships = new Set(getRelationshipChildren(document));
    const unexpected: Element[] = [];
    for (let child = document.documentElement?.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 1 && !relationships.has(child as Element)) unexpected.push(child as Element);
    }
    return unexpected;
};

/** Attribute each replacement to its raw child, retaining the root's inherited MCE directives. */
const childLeavesOnlyRelationships = (document: Document, child: Element, part: string): boolean => {
    const isolated = document.cloneNode(false) as Document;
    const root = document.documentElement!.cloneNode(false);
    isolated.appendChild(root);
    root.appendChild(child.cloneNode(true));
    return unexpectedChildren(applyMarkupCompatibility(isolated, { understood: understoodNamespacesForPart(part) }).document).length === 0;
};

/**
 * OPC CT_Relationships allows only Relationship children, without xsd:any.
 * Word 16.115 first applies markup compatibility (session 33 s33d/s33e,
 * #632): selected Relationship children and removed Ignorable children open
 * cleanly. Surviving foreign children and wrong roots still offer recovery
 * (#433). It applies that processing to the package relationships part's
 * root too (#646): a root resolving to exactly one Relationships element
 * opened clean, whether an mc:AlternateContent Fallback (sessions 36 and 38,
 * s36c and s38a), a selected mc:Choice or an Ignorable mc:ProcessContent
 * wrapper (s38a) held it. A root resolving to two Relationship elements
 * (s36c) or to none, its Relationships in an unselected Choice beside an
 * empty Fallback (s38a), offered recovery; the processor returns the stored
 * view there (#640). So the package part's root is judged
 * on the processed view every relationships reader uses
 * (`relationshipElements`, main-part discovery); a root that resolves to
 * anything else keeps the finding. A root-wrapped sidecar is unobserved and
 * keeps the stored-root verdict, as lenient's XSD walk does for every part.
 * Keep removed/replaced raw children as strict-only errors through the
 * profile's existing demotion of non-Word-blocking codes.
 */
export const opcRelationshipsStructureCheck: CheckDefinition = {
    id: "opc-relationships-structure",
    description: "Relationships parts have a package Relationships root and only direct Relationship children.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.partNames) {
            if (!isRelationshipsPartName(part)) continue;
            const parsed = ctx.xml(part);
            if (!parsed?.ok) continue;
            const effective = ctx.effectiveXml(part);
            if (!effective?.ok) continue;
            // The stored root is required even if compatibility processing
            // removes it — except in the package relationships part, whose
            // root Word resolves through that processing (#646).
            const storedRoot = parsed.document.documentElement;
            const root = isPackageRelationshipsPart(part) ? effective.document.documentElement : storedRoot;
            if (root?.localName !== "Relationships" || root.namespaceURI !== NS.PR) {
                issues.push({
                    code: opcRelationshipsStructureCheck.id,
                    path: part,
                    severity: "error",
                    message: `Relationships part root must be {${NS.PR}}Relationships.`,
                });
                continue;
            }
            const surviving = unexpectedChildren(effective.document);
            for (const element of surviving) {
                issues.push({
                    code: opcRelationshipsStructureCheck.id,
                    path: part,
                    severity: "error",
                    message: `Unexpected child {${element.namespaceURI ?? ""}}${element.localName} in Relationships; expected {${NS.PR}}Relationship.`,
                });
            }
            if (storedRoot?.localName !== "Relationships" || storedRoot.namespaceURI !== NS.PR) {
                // A wrapper root's Choice and Fallback branches are not
                // children in Relationships; the root check above judged the
                // wrapper (#646).
                continue;
            }
            for (const element of unexpectedChildren(parsed.document)) {
                // When another sibling survives, process this child in isolation
                // to avoid attributing that sibling's defect to removed markup.
                // Use the same part capabilities as ctx.effectiveXml, including w14.
                if (surviving.length > 0 && !childLeavesOnlyRelationships(parsed.document, element, part)) continue;
                issues.push({
                    code: "opc-relationships-markup-compatibility",
                    path: part,
                    severity: "error",
                    message: `Stored child {${element.namespaceURI ?? ""}}${element.localName} in Relationships is removed or replaced by markup compatibility, leaving only Relationship children. Microsoft Word opened each tested shape of this markup clean; other shapes are untested. Strict OPC requires direct Relationship children in the stored XML.`,
                });
            }
        }
        return issues;
    },
};
