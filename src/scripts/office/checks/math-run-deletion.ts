/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";
import { MATH_NAMESPACES, WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

export const mathRunDeletionCheck: CheckDefinition = {
    id: "word-math-run-deletion",
    description: "Report tracked deletions nested directly inside math runs instead of wrapping the runs.",
    groups: ["math", "tracked-changes"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        // Math runs occur in stories, not in the styles content model. Both
        // conformance classes carry the defect: Transitional and Strict spell
        // the WML and math namespaces differently but share CT_R.
        for (const part of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const w of WORD_PARAGRAPH_NAMESPACES) {
                for (const deletion of getElementsByTagNameNSAll(parsed.document, w, "del")) {
                    const parent = deletion.parentNode as Element | null;
                    if (parent === null || !MATH_NAMESPACES.includes(parent.namespaceURI ?? "") || parent.localName !== "r") {
                        continue;
                    }
                    // shared-math.xsd CT_R excludes revisions; CT_OMath/CT_OMathArg
                    // admit CT_RunTrackChange via EG_PContentMath (wml.xsd).
                    issues.push({
                        severity: "warning",
                        code: "word-math-run-deletion",
                        path: part,
                        message:
                            "w:del is a direct child of math run m:r; schema CT_R requires run properties and run content, not revision children. " +
                            "Use w:del to wrap the whole m:r inside m:oMath or a math argument, or wrap the whole math zone at run level. " +
                            "Use m:ctrlPr/w:del only for control-property revisions.",
                    });
                }
            }
        }
        return issues;
    },
};
