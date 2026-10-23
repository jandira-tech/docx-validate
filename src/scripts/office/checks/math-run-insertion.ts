/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

export const mathRunInsertionCheck: CheckDefinition = {
    id: "math-run-insertion",
    description: "Report tracked insertions nested directly inside math runs instead of wrapping the runs.",
    groups: ["math", "tracked-changes"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        // Math runs occur in stories, not in the styles content model.
        for (const part of ctx.storyParts()) {
            const parsed = ctx.xml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const insertion of getElementsByTagNameNSAll(parsed.document, W, "ins")) {
                const parent = insertion.parentNode as Element | null;
                if (parent?.namespaceURI !== M || parent.localName !== "r") {
                    continue;
                }
                // shared-math.xsd CT_R excludes revisions; CT_OMath/CT_OMathArg
                // admit CT_RunTrackChange via EG_PContentMath (wml.xsd).
                issues.push({
                    severity: "warning",
                    code: "math-run-insertion",
                    path: part,
                    message:
                        "w:ins is a direct child of math run m:r; schema CT_R requires run properties and run content, not revision children. " +
                        "Use w:ins to wrap the whole m:r inside m:oMath or a math argument, or wrap the whole math zone at run level. " +
                        "Use m:ctrlPr/w:ins only for control-property revisions.",
                });
            }
        }
        return issues;
    },
};
