/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A wp:extent cx or cy that is a negative integer makes Word
 * report an open error (fixture 2213). One negative attribute
 * is enough. The same drawings with positive extents open clean.
 * Zero and positive extents are not this defect.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP14_DRAWING_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const WORD_DRAWING_SCALAR_TEXT_ELEMENTS: readonly { namespace: string; localName: string }[] = [
    { namespace: WP_NAMESPACE, localName: "align" },
    { namespace: WP_NAMESPACE, localName: "posOffset" },
    { namespace: WP14_DRAWING_NAMESPACE, localName: "pctWidth" },
    { namespace: WP14_DRAWING_NAMESPACE, localName: "pctHeight" },
];

const previewRepr = (text: string, max: number): string => {
    // Mirrors Python's `repr(text)` truncation: wraps in single quotes and
    // escapes embedded backslashes / quotes the same way repr() does.
    const repr = pythonRepr(text);
    return repr.length > max ? `${repr.slice(0, max)}...` : repr;
};
const pythonRepr = (s: string): string => {
    let out = "";
    for (const ch of s) {
        if (ch === "\\") {
            out += "\\\\";
        } else if (ch === "'") {
            out += "\\'";
        } else if (ch === "\n") {
            out += "\\n";
        } else if (ch === "\r") {
            out += "\\r";
        } else if (ch === "\t") {
            out += "\\t";
        } else {
            out += ch;
        }
    }
    return `'${out}'`;
};

export const wordDrawingScalarWhitespaceCheck: CheckDefinition = {
    id: "word-drawing-scalar-whitespace",
    description: "Drawing scalar text has no leading or trailing whitespace.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const { namespace, localName } of WORD_DRAWING_SCALAR_TEXT_ELEMENTS) {
                const elems = getElementsByTagNameNSAll(dom, namespace, localName);
                for (const elem of elems) {
                    for (let child = elem.firstChild; child; child = child.nextSibling) {
                        if (child.nodeType !== 3) {
                            continue;
                        }
                        const text = child as Text;
                        const trimmed = text.data.trim();
                        if (trimmed === text.data) {
                            continue;
                        }
                        issues.push({
                            severity: "error",
                            message:
                                `<${elem.tagName}> contains leading/trailing whitespace around scalar value ` +
                                `${previewRepr(text.data, 80)}; Microsoft Word refuses this shape in drawing anchors.`,
                            path: xmlFile,
                            code: "word-drawing-scalar-whitespace",
                        });
                    }
                }
            }
        }

        return issues;
    },
};

export const wordDrawingArcRadiusCheck: CheckDefinition = {
    id: "word-drawing-arc-radius",
    description: "Drawing arc radii use unsigned integer values.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const arcs = getElementsByTagNameNSAll(dom, DRAWINGML_NAMESPACE, "arcTo");
            for (const arc of arcs) {
                for (const attr of ["wR", "hR"] as const) {
                    const value = arc.getAttribute(attr);
                    if (value === null || /^[0-9]+$/.test(value)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `<a:arcTo> ${attr}="${value}" is not an integer. ` +
                            "Microsoft Word cannot open the file. Use an integer radius.",
                        path: xmlFile,
                        code: "word-drawing-arc-radius",
                    });
                }
            }
        }

        return issues;
    },
};

export const wordDrawingExtentNegativeCheck: CheckDefinition = {
    id: "word-drawing-extent-negative",
    description: "Drawing extents have no negative integer dimensions.",
    groups: ["drawings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const extents = getElementsByTagNameNSAll(dom, WP_NAMESPACE, "extent");
            for (const extent of extents) {
                for (const attr of ["cx", "cy"] as const) {
                    const value = extent.getAttribute(attr);
                    if (value === null || !/^-\d+$/.test(value)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `<wp:extent> ${attr}="${value}" is negative. ` + "Microsoft Word cannot open the file. Use a positive extent.",
                        path: xmlFile,
                        code: "word-drawing-extent-negative",
                    });
                }
            }
        }

        return issues;
    },
};
