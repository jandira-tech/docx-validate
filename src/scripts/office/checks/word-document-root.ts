/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The required main document root, in Word's effective view.
 *
 * Word evidence (#317, Word 16.115, 2026-10-23): MINIMAL_PARTS from
 * tests/support/minimal-docx.ts with its main XML replaced by
 * `<x:foo xmlns:x="urn:x"/>` makes Word refuse to open the file
 * (sha256 31a09771…2011). An unknown x:document root with mc:Ignorable="x"
 * enclosing a w:body containing "Hello" opens without a prompt but shows
 * 0 characters: Word silently drops all content (sha256 0f5e75cb…cbc6;
 * issue correction). The unchanged w:document control opens cleanly.
 *
 * The package officeDocument relationship identifies the main part, as in
 * the schema walk's mce-main-root-ignored check (#287). Non-main roots are
 * outside this rule; their mce-root-ignored remains informational.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { findEquivalentPart, resolvePartName } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_TYPES = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
]);

export const wordDocumentRootCheck: CheckDefinition = {
    id: "word-document-root",
    description: "The main document part retains a WordprocessingML document root after markup compatibility processing.",
    groups: ["xml"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        // The package relationships part is found by OPC equivalence, so
        // `_RELS/.RELS` names it too, as the schema walk finds it (#364).
        const relationships = ctx.xml(findEquivalentPart(ctx.partNames, "_rels/.rels") ?? "_rels/.rels");
        if (!relationships?.ok) return [];

        const issues: ValidationIssue[] = [];
        for (const relationship of getElementsByTagNameNSAll(relationships.document, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship")) {
            if (
                relationship.getAttribute("TargetMode") === "External" ||
                !OFFICE_DOCUMENT_TYPES.has(relationship.getAttribute("Type") ?? "")
            ) {
                continue;
            }
            // OPC part names compare ASCII case-insensitively with unreserved
            // escapes decoded, so `WORD/document.xml` names word/document.xml (#344).
            const part = resolvePartName(ctx.partNames, relationship.getAttribute("Target") ?? "");
            if (part === undefined) continue;
            const parsed = ctx.effectiveXml(part);
            // XML syntax and missing targets belong to their existing checks.
            if (!parsed?.ok) continue;

            const root = parsed.document.documentElement;
            if (root?.localName === "document" && WORD_PARAGRAPH_NAMESPACES.some((namespace) => namespace === root.namespaceURI)) {
                continue;
            }
            issues.push({
                code: wordDocumentRootCheck.id,
                severity: "error",
                path: part,
                message: root
                    ? `Main document part ${part} has effective root {${root.namespaceURI ?? ""}}${root.localName}; expected w:document in the Transitional or Strict WordprocessingML namespace. Microsoft Word refuses to open the file.`
                    : `Main document part ${part} has no effective root after markup compatibility processing. Microsoft Word opens it empty and silently drops all content.`,
            });
        }
        return issues;
    },
};

/**
 * ECMA-376's Minimal WordprocessingML Document requires package relationships
 * pointing to a main document part (#380). No real-Word observation exists
 * yet, so this spec finding remains advisory under full Word-valid.
 */
export const wordMainDocumentPartCheck: CheckDefinition = {
    id: "word-main-document-part",
    description: "The package has an internal officeDocument relationship targeting an existing main document part.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const relationshipsPart = findEquivalentPart(ctx.partNames, "_rels/.rels");
        if (relationshipsPart === undefined) {
            return [
                {
                    code: wordMainDocumentPartCheck.id,
                    severity: "error",
                    message:
                        "Package relationships part _rels/.rels is missing; ECMA-376 requires the main document part and its package relationship.",
                },
            ];
        }
        const relationships = ctx.xml(relationshipsPart);
        if (!relationships?.ok) return [];

        const hasMainPart = getElementsByTagNameNSAll(relationships.document, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship").some(
            (relationship) => {
                if (
                    relationship.getAttribute("TargetMode") === "External" ||
                    !OFFICE_DOCUMENT_TYPES.has(relationship.getAttribute("Type") ?? "")
                ) {
                    return false;
                }
                const part = resolvePartName(ctx.partNames, relationship.getAttribute("Target") ?? "");
                return part !== undefined && ctx.partNames.includes(part);
            },
        );
        if (hasMainPart) return [];

        return [
            {
                code: wordMainDocumentPartCheck.id,
                severity: "error",
                path: relationshipsPart,
                message:
                    "No internal officeDocument relationship targets an existing part; ECMA-376 requires the main document part and its package relationship.",
            },
        ];
    },
};
