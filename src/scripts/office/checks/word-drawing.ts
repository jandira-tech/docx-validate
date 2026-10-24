/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

/** A drawing container needs its own direct wp:extent child. */
export const wordDrawingExtentMissingCheck: CheckDefinition = {
    id: "word-drawing-extent-missing",
    description: "Word drawing inline and anchor containers have a direct extent child.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const local of ["inline", "anchor"]) {
                for (const container of getElementsByTagNameNSAll(parsed.document, WP_NAMESPACE, local)) {
                    const hasExtent = Array.from(container.childNodes).some(
                        (child) =>
                            child.nodeType === 1 &&
                            (child as Element).namespaceURI === WP_NAMESPACE &&
                            (child as Element).localName === "extent",
                    );
                    if (hasExtent) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        code: "word-drawing-extent-missing",
                        path: part,
                        message: `wp:extent is missing from parent wp:${local}; expected a direct wp:extent before a:graphic. Microsoft Word cannot open an incomplete drawing scaffold.`,
                    });
                }
            }
        }
        return issues;
    },
};
