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
import { checkXmlSyntax } from "../../../lib/xml-syntax";
import type { ValidationIssue } from "../../../lib/types";

const XML_PART = /\.(?:xml|rels)$/i;

export const xmlSyntaxCheck: CheckDefinition = {
    id: "xml-syntax",
    description: "Every XML part of the package is well-formed, parsed from its original bytes without recovery.",
    groups: ["xml"],
    formats: ["docx", "pptx", "xlsx"],
    run: async (ctx): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.partNames) {
            const bytes = XML_PART.test(part) ? ctx.bytes(part) : undefined;
            if (bytes !== undefined) {
                issues.push(...(await checkXmlSyntax(bytes, part)));
            }
        }
        return issues;
    },
};
