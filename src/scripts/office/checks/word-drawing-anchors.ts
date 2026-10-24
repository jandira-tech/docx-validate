/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * An empty `w:drawing` (no element child) is not a drawing Word can load.
 * Inside `mc:Choice Requires="wps"` Word reports an open error. Inside a
 * `w:r` Word offers to recover the file. One such element is enough.
 * A drawing placed directly under `w:p` is a different case: Word drops
 * that misplaced run content and opens clean, so it is not flagged here.
 *
 * A wp:anchor missing locked, layoutInCell, or allowOverlap makes
 * Word refuse the file (fixture 1061). One missing attribute on
 * one anchor is enough. Adding that attribute opens clean.
 * simplePos, relativeHeight, and the wrap distances are not this
 * defect.
 *
 * A wp:positionH or wp:positionV that contains both wp:posOffset
 * and wp:align makes Word refuse the file (fixture 1061). One
 * such element is enough. Either child alone opens clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export const wordEmptyDrawingCheck: CheckDefinition = {
    id: "word-empty-drawing",
    description: "Drawings contain an element child unless Word discards their paragraph-level placement.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const drawings = getElementsByTagNameNSAll(dom, ns, "drawing");
                for (const drawing of drawings) {
                    let hasElementChild = false;
                    for (let child = drawing.firstChild; child; child = child.nextSibling) {
                        if (child.nodeType === 1) {
                            hasElementChild = true;
                            break;
                        }
                    }
                    if (hasElementChild) {
                        continue;
                    }
                    const parent = drawing.parentNode;
                    if (parent && parent.nodeType === 1) {
                        const parentElem = parent as Element;
                        const parentLocal = parentElem.localName || parentElem.tagName.split(":").pop() || "";
                        if (parentLocal === "p" && WORD_PARAGRAPH_NAMESPACES.includes(parentElem.namespaceURI ?? "")) {
                            continue;
                        }
                    }
                    issues.push({
                        severity: "error",
                        message:
                            "An empty <w:drawing> has no child element. Microsoft Word does not open the file cleanly. " +
                            "Remove the drawing, or give it a wp:anchor or wp:inline child.",
                        path: xmlFile,
                        code: "word-empty-drawing",
                    });
                }
            }
        }

        return issues;
    },
};

export const wordAnchorMissingAttributeCheck: CheckDefinition = {
    id: "word-anchor-missing-attribute",
    description: "Drawing anchors declare locked, layoutInCell and allowOverlap.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        const required = ["locked", "layoutInCell", "allowOverlap"] as const;
        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const anchors = getElementsByTagNameNSAll(dom, WP_NAMESPACE, "anchor");
            for (const anchor of anchors) {
                for (const attr of required) {
                    if (anchor.getAttribute(attr) !== null) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message: `A wp:anchor is missing ${attr}. ` + "Microsoft Word cannot open the file. Add that attribute.",
                        path: xmlFile,
                        code: "word-anchor-missing-attribute",
                    });
                }
            }
        }

        return issues;
    },
};

export const wordPositionOffsetAndAlignCheck: CheckDefinition = {
    id: "word-position-offset-and-align",
    description: "Drawing positions choose either an offset or alignment.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const localName of ["positionH", "positionV"] as const) {
                const positions = getElementsByTagNameNSAll(dom, WP_NAMESPACE, localName);
                for (const position of positions) {
                    let hasOffset = false;
                    let hasAlign = false;
                    for (const child of Array.from(position.childNodes)) {
                        if (!child || child.nodeType !== 1) {
                            continue;
                        }
                        const element = child as Element;
                        if (element.namespaceURI !== WP_NAMESPACE) {
                            continue;
                        }
                        const local = element.localName || element.tagName.split(":").pop() || "";
                        if (local === "posOffset") {
                            hasOffset = true;
                        }
                        if (local === "align") {
                            hasAlign = true;
                        }
                    }
                    if (!hasOffset || !hasAlign) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `A wp:${localName} contains both wp:posOffset and wp:align. ` +
                            "Microsoft Word cannot open the file. Keep only one of those children.",
                        path: xmlFile,
                        code: "word-position-offset-and-align",
                    });
                }
            }
        }

        return issues;
    },
};
