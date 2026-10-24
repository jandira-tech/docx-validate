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
    findContentTypesStream,
    isContentTypesStreamName,
    partExtension,
} from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

export const wordContentTypeInvalidCheck: CheckDefinition = {
    id: "word-content-type-invalid",
    description: "Content types using Word-rejected invalid-* values are diagnosed.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
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

        return issues;
    },
};
