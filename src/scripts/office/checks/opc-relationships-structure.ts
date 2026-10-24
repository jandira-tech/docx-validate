/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { isRelationshipsPartName } from "../../../lib/content-types";
import { NS, type ValidationIssue } from "../../../lib/types";
import { getRelationshipChildren } from "../../../lib/xml-helpers";

/**
 * OPC CT_Relationships allows only Relationship children, without xsd:any.
 * Word 16.115 offers recovery for a foreign sibling, a wrapped main
 * relationship and a wrong root in _rels/.rels (#433).
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
            // Judge the stored XML; compatibility processing must not hide foreign children.
            const parsed = ctx.xml(part);
            if (!parsed?.ok) continue;
            const root = parsed.document.documentElement;
            if (root?.localName !== "Relationships" || root.namespaceURI !== NS.PR) {
                issues.push({
                    code: opcRelationshipsStructureCheck.id,
                    path: part,
                    severity: "error",
                    message: `Relationships part root must be {${NS.PR}}Relationships.`,
                });
                continue;
            }
            const relationships = new Set(getRelationshipChildren(parsed.document));
            for (let child = root.firstChild; child; child = child.nextSibling) {
                if (child.nodeType !== 1) continue;
                const element = child as Element;
                if (relationships.has(element)) continue;
                issues.push({
                    code: opcRelationshipsStructureCheck.id,
                    path: part,
                    severity: "error",
                    message: `Unexpected child {${element.namespaceURI ?? ""}}${element.localName} in Relationships; expected {${NS.PR}}Relationship.`,
                });
            }
        }
        return issues;
    },
};
