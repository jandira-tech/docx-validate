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

/**
 * The id of the XML well-formedness check. The validators' own XML pass
 * does the same work on the same bytes, so they report it under this id
 * when they disclose the checks a run executed (#319).
 */
export const XML_SYNTAX_CHECK_ID = "xml-syntax";

export const xmlSyntaxCheck: CheckDefinition = {
    id: XML_SYNTAX_CHECK_ID,
    description: "Every XML part of the package is well-formed, parsed from its original bytes without recovery.",
    groups: ["xml"],
    formats: ["docx", "pptx", "xlsx"],
    // A malformed part hides only itself from the checks that require this
    // one; they still judge every part that parsed (#307).
    gatesByPart: true,
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
