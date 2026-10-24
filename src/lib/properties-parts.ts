/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { canonicalPartName, findEquivalentPart, resolvePartName } from "./content-types";
import { NS } from "./types";
import type { ValidationContext } from "./validation-context";
import { getElementsByTagNameNSAll } from "./xml-helpers";
import { checkXmlSyntax } from "./xml-syntax";

const PROPERTIES_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties",
]);

/**
 * Canonical names of properties parts Word reads through package relationships.
 * Word 16.115 offers recovery for syntax, prefix and encoding defects in
 * referenced core/app/custom parts; the unreferenced SuperDoc custom.xml opens
 * clean (#409; docs/plans/evidence/word-probes-2026-10-24.md).
 * Session 3c settles the rule: unreferenced malformed core/app/custom parts
 * open clean, while a truncated core part stored as props/meta.xml and named by
 * its relationship makes Word offer recovery. Compute once per run and pass the
 * set explicitly to the shared Word-blocking predicate.
 */
export const propertiesParts = async (context: ValidationContext): Promise<ReadonlySet<string>> => {
    const parts = new Set<string>();
    const name = findEquivalentPart(context.partNames, "_rels/.rels");
    if (name === undefined) {
        return parts;
    }
    // Do not use relationships recovered by the DOM parser. An unreadable
    // package relationships part already blocks Word-valid on its own (#388).
    const issues = await checkXmlSyntax(context.bytes(name)!, name);
    if (issues.some((issue) => issue.severity === "error")) {
        return parts;
    }
    const parsed = context.xml(name);
    if (!parsed?.ok) {
        return parts;
    }
    for (const relationship of getElementsByTagNameNSAll(parsed.document, NS.PR, "Relationship")) {
        if (
            relationship.getAttribute("TargetMode") === "External" ||
            !PROPERTIES_RELATIONSHIP_TYPES.has(relationship.getAttribute("Type") ?? "")
        ) {
            continue;
        }
        // Package relationships resolve from the root, never from _rels/.
        const target = resolvePartName(context.partNames, relationship.getAttribute("Target") ?? "");
        if (target !== undefined) {
            parts.add(canonicalPartName(target));
        }
    }
    return parts;
};
