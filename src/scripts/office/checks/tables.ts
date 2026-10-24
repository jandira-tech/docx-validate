/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Table-structure checks over the Word story parts. */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/**
 * A `w:tr` with no `w:tc` makes Word report an open error (fixture 1993).
 * Removing `w:gridBefore` and leaving the row still errors. One cell opens
 * clean, so `gridBefore` is not the cause.
 */
export const wordRowWithoutCellCheck: CheckDefinition = {
    id: "word-row-without-cell",
    description: "Every table row in a Word story contains at least one w:tc cell (Word refuses to open the file).",
    groups: ["tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) continue;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const row of getElementsByTagNameNSAll(parsed.document, ns, "tr")) {
                    if (getElementsByTagNameNSAll(row, ns, "tc").length > 0) continue;
                    issues.push({
                        severity: "error",
                        message: "A table row contains no w:tc cell. Microsoft Word cannot open the file. Remove the row or add a cell.",
                        path: part,
                        code: "word-row-without-cell",
                    });
                }
            }
        }
        return issues;
    },
};
