/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

// The sequence prefix before extent in CT_Inline / CT_Anchor. The schema
// integration test pins both the parent set and these ordered prefixes.
export const DRAWING_EXTENT_PREDECESSORS: ReadonlyMap<string, readonly string[]> = new Map([
    ["inline", []],
    ["anchor", ["simplePos", "positionH", "positionV"]],
]);

export const wordDrawingExtentMisplacedCheck: CheckDefinition = {
    id: "word-drawing-extent-misplaced",
    description: "Drawing extents occupy their CT_Inline or CT_Anchor sequence position.",
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
            for (const extent of getElementsByTagNameNSAll(parsed.document, WP_NAMESPACE, "extent")) {
                const parent = extent.parentNode!;
                const container = parent.nodeType === 1 ? (parent as Element) : undefined;
                const predecessors =
                    container?.namespaceURI === WP_NAMESPACE ? DRAWING_EXTENT_PREDECESSORS.get(container.localName) : undefined;
                if (container && predecessors) {
                    const children = Array.from(container.childNodes).filter((child): child is Element => child.nodeType === 1);
                    const index = children.indexOf(extent);
                    const isPredecessor = (child: Element) => child.namespaceURI === WP_NAMESPACE && predecessors.includes(child.localName);
                    // Missing positioning children are a different defect: this
                    // rule only checks where an existing extent sits.
                    if (children.slice(0, index).every(isPredecessor) && !children.slice(index + 1).some(isPredecessor)) {
                        continue;
                    }
                }
                issues.push({
                    severity: "error",
                    code: "word-drawing-extent-misplaced",
                    path: part,
                    message: `wp:extent has actual parent ${parent.nodeName}; expected a single direct extent in wp:inline (first) or wp:anchor (after wp:simplePos, wp:positionH, wp:positionV), before a:graphic and the other drawing properties.`,
                });
            }
        }
        return issues;
    },
};
