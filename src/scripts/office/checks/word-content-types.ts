/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word compatibility checks migrated without changing their issue details.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    asciiLowercase,
    canonicalPartName,
    CT_DUPLICATE_DECLARATION,
    duplicateContentTypeDeclarations,
    findContentTypesStream,
    isContentTypesStreamName,
    partExtension,
    parseContentTypes,
    resolveContentType,
} from "../../../lib/content-types";
import { wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** Package declarations are checked under every profile and format, including lenient. */
export const ctDuplicateDeclarationCheck: CheckDefinition = {
    id: CT_DUPLICATE_DECLARATION,
    description: "Content-type Default extensions and Override part names are unique under OPC equivalence.",
    groups: ["content-types"],
    formats: ["docx", "pptx", "xlsx"],
    requires: ["xml-syntax"],
    run: (ctx) => {
        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) return [];
        // Judged after Markup Compatibility processing (OPC [M1.19]), like
        // word-content-type-invalid: a duplicate in an ignored branch is not
        // judged; Word refuses the markup itself (#459).
        const parsed = ctx.effectiveXml(stream);
        return parsed?.ok ? duplicateContentTypeDeclarations(parsed.document, stream) : [];
    },
};

/**
 * Word 16.115.26100424 offers recovery for an empty ContentType on the main
 * document's Override (session 22 in
 * docs/plans/evidence/word-probes-2026-10-24.md, #581). Other parts, Defaults,
 * missing attributes and whitespace-only values are unobserved and stay out.
 * Resolve the effective stream through the shared table, like the generic
 * and inexact checks; the package officeDocument relationship names the main.
 */
const emptyMainContentTypeFindings = async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
    const stream = findContentTypesStream(ctx.partNames);
    if (stream === undefined) return [];
    const parsed = ctx.effectiveXml(stream);
    if (!parsed?.ok) return [];
    const table = parseContentTypes(parsed.document);
    const overrides = getElementsByTagNameNSAll(parsed.document, CONTENT_TYPES_NAMESPACE, "Override");
    const issues: ValidationIssue[] = [];
    for (const { part } of await wordTypedRelationshipTargets(ctx, new Set(["officeDocument"]), profile)) {
        const resolved = resolveContentType(table, part);
        if (resolved?.source !== "override" || resolved.contentType !== "") continue;
        // The table represents a missing ContentType as "" too. Only the
        // explicitly empty attribute on the winning declaration was probed.
        const declaration = overrides.find((elem) => canonicalPartName(elem.getAttribute("PartName") ?? "") === canonicalPartName(part));
        if (!declaration?.hasAttribute("ContentType")) continue;
        issues.push({
            severity: "error",
            code: "word-content-type-invalid",
            path: stream,
            message: `Microsoft Word refuses empty content type '' for ${part} (main document Override)`,
        });
    }
    return issues;
};

export const wordContentTypeInvalidCheck: CheckDefinition = {
    id: "word-content-type-invalid",
    description: "Word-rejected invalid-* content types and an empty main-document Override are diagnosed.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        const contentTypesFile = findContentTypesStream(ctx.partNames);
        if (contentTypesFile === undefined) return issues;
        const parsed = ctx.effectiveXml(contentTypesFile);
        if (!parsed?.ok) {
            return issues;
        }
        const dom = parsed.document;
        const check = (elem: Element, attrName: "PartName" | "Extension"): void => {
            const contentType = elem.getAttribute("ContentType");
            if (!contentType || !contentType.startsWith("invalid-")) {
                return;
            }
            const target = elem.getAttribute(attrName) ?? "";
            issues.push({
                severity: "error",
                message: `Microsoft Word refuses content type '${contentType}' for ${target}`,
                path: contentTypesFile,
                code: "word-content-type-invalid",
            });
        };

        const overrides = getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Override");
        // Word ignores declarations that do not apply to a stored part (#401).
        // The stream is a ZIP item, not a part covered by its own declarations.
        const storedParts = new Set(ctx.partNames.filter((name) => !isContentTypesStreamName(name)).map(canonicalPartName));
        const overriddenParts = new Set(overrides.map((elem) => canonicalPartName(elem.getAttribute("PartName") ?? "")));
        for (const elem of overrides) {
            if (storedParts.has(canonicalPartName(elem.getAttribute("PartName") ?? ""))) {
                check(elem, "PartName");
            }
        }
        const defaultExtensions = new Set([...storedParts].filter((name) => !overriddenParts.has(name)).map(partExtension));
        const defaults = getElementsByTagNameNSAll(dom, CONTENT_TYPES_NAMESPACE, "Default");
        for (const elem of defaults) {
            const extension = elem.getAttribute("Extension");
            if (extension && defaultExtensions.has(asciiLowercase(extension))) {
                check(elem, "Extension");
            }
        }

        return [...issues, ...(await emptyMainContentTypeFindings(ctx, profile))];
    },
};
