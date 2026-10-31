/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { citesWordSession } from "../../../lib/conformance-class";
import {
    asciiLowercase,
    equivalentPartLookup,
    findContentTypesStream,
    findEquivalentPart,
    parseContentTypes,
    resolveContentType,
    resolveTarget,
} from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import { getRelationshipChildren } from "../../../lib/xml-helpers";
import { xmlPartNames } from "./xml";

const THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

/**
 * Word 16.115 offered recovery for PNG package thumbnails and opened JPEG
 * ones clean (sessions 89 s89a and 90 s90a, #1186; two runs per row, all
 * four readings agree). Other image types, parameterized types and reaches
 * from relationships parts other than the package's are unprobed.
 */
export const wordPackageThumbnailTypeCheck: CheckDefinition = {
    id: "word-package-thumbnail-type",
    description: "Internal package thumbnail relationships do not name stored parts typed image/png.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const stream = findContentTypesStream(ctx.partNames);
        const relsPart = findEquivalentPart(ctx.partNames, "_rels/.rels");
        if (stream === undefined || relsPart === undefined) return [];
        const contentTypes = ctx.effectiveXml(stream);
        const relationships = ctx.effectiveXml(relsPart);
        if (!contentTypes?.ok || !relationships?.ok) return [];

        const table = parseContentTypes(contentTypes.document);
        const lookup = equivalentPartLookup(ctx.partNames);
        const issues: ValidationIssue[] = [];
        for (const relationship of getRelationshipChildren(relationships.document)) {
            if (relationship.getAttribute("TargetMode") === "External" || relationship.getAttribute("Type") !== THUMBNAIL) continue;
            const part = resolveTarget(lookup, "", relationship.getAttribute("Target") ?? "");
            if (part === undefined) continue;
            const resolved = resolveContentType(table, part);
            // Compare type/subtype case-insensitively, but do not strip a
            // parameter or padding: those spellings were not probed.
            if (resolved === null || asciiLowercase(resolved.contentType) !== "image/png") continue;
            const recovery = citesWordSession(
                ctx,
                xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
            )
                ? " Word 16.115 offered recovery for a PNG package thumbnail and opened a JPEG one clean (Word probes sessions 89 and 90, #1186)."
                : "";
            issues.push({
                severity: "error",
                code: wordPackageThumbnailTypeCheck.id,
                path: part,
                message: `${part}, named by an internal package thumbnail relationship, is typed image/png.${recovery}`,
            });
        }
        return issues;
    },
};
