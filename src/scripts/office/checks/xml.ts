/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * XML syntax gate. Every DOM-reading check requires it, so a malformed part
 * stops those checks before they run instead of letting them crash or report
 * findings about a document the parser had to guess at.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";

const XML_PART = /\.(?:xml|rels)$/i;

export const xmlSyntaxCheck: CheckDefinition = {
    id: "xml-syntax",
    description: "Every XML part of the package is well-formed.",
    groups: ["xml"],
    formats: ["docx", "pptx", "xlsx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.partNames) {
            if (!XML_PART.test(part)) continue;
            const parsed = ctx.xml(part);
            if (parsed && !parsed.ok) issues.push({ severity: "error", code: "xml-syntax", message: parsed.message, path: part });
        }
        return issues;
    },
};
