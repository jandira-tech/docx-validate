/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * w:delText or w:delInstrText inside w:txbxContent makes Word
 * offer recovery when no w:del inside that text box wraps it.
 * That content sits in a VML text box (fixture 0011) or a
 * drawingML wps:txbx (fixture 0044). The same elements in the
 * document body are tolerated. An mc:Fallback text box is
 * tolerated (fixture 2054): Word uses the sibling mc:Choice.
 * A w:del inside the text box is tolerated (fixture 0948).
 * A w:del outside the text box does not cover the text-box story.
 *
 * w:delText inside w:moveFrom makes Word offer recovery (fixture
 * 0726). One such element is enough. w:t inside w:moveFrom opens
 * clean, and w:delText in the document body is tolerated.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const hasMoveFromAncestor = (node: Node): boolean => {
    let parent = node.parentNode;
    while (parent && parent.nodeType === 1) {
        const elem = parent as Element;
        const local = elem.localName || elem.tagName.split(":").pop() || "";
        if (local === "moveFrom" && WORD_PARAGRAPH_NAMESPACES.includes(elem.namespaceURI ?? "")) {
            return true;
        }
        parent = parent.parentNode;
    }
    return false;
};

const hasTextBoxContentAncestor = (node: Node): boolean => {
    let parent = node.parentNode;
    while (parent && parent.nodeType === 1) {
        const elem = parent as Element;
        const local = elem.localName || elem.tagName.split(":").pop() || "";
        if (local === "txbxContent" && WORD_PARAGRAPH_NAMESPACES.includes(elem.namespaceURI ?? "")) {
            return true;
        }
        parent = parent.parentNode;
    }
    return false;
};

const hasTextBoxDeletionAncestor = (node: Node): boolean => {
    let parent = node.parentNode;
    while (parent && parent.nodeType === 1) {
        const elem = parent as Element;
        const local = elem.localName || elem.tagName.split(":").pop() || "";
        const ns = elem.namespaceURI ?? "";
        if (local === "txbxContent" && WORD_PARAGRAPH_NAMESPACES.includes(ns)) {
            return false;
        }
        if (local === "del" && WORD_PARAGRAPH_NAMESPACES.includes(ns)) {
            return true;
        }
        parent = parent.parentNode;
    }
    return false;
};

const hasMarkupCompatibilityFallbackAncestor = (node: Node): boolean => {
    let parent = node.parentNode;
    while (parent && parent.nodeType === 1) {
        const elem = parent as Element;
        const local = elem.localName || elem.tagName.split(":").pop() || "";
        if (local === "Fallback" && elem.namespaceURI === MC_NAMESPACE) {
            return true;
        }
        parent = parent.parentNode;
    }
    return false;
};

export const wordTextboxDeletedTextCheck: CheckDefinition = {
    id: "word-textbox-deleted-text",
    description: "Deleted text in a text box is wrapped in a deletion within that story.",
    groups: ["tracked-changes"],
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
                for (const local of ["delText", "delInstrText"] as const) {
                    const nodes = getElementsByTagNameNSAll(dom, ns, local);
                    for (const node of nodes) {
                        if (!node || !hasTextBoxContentAncestor(node)) {
                            continue;
                        }
                        if (hasMarkupCompatibilityFallbackAncestor(node)) {
                            continue;
                        }
                        if (hasTextBoxDeletionAncestor(node)) {
                            continue;
                        }
                        issues.push({
                            severity: "error",
                            message:
                                "A text box contains deleted text. Microsoft Word offers to recover the file. " +
                                "Remove the w:delText or w:delInstrText from the text box.",
                            path: xmlFile,
                            code: "word-textbox-deleted-text",
                        });
                    }
                }
            }
        }

        return issues;
    },
};

export const wordMoveFromDeletedTextCheck: CheckDefinition = {
    id: "word-move-from-deleted-text",
    description: "Move sources use ordinary text instead of deleted text.",
    groups: ["tracked-changes"],
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
                const nodes = getElementsByTagNameNSAll(dom, ns, "delText");
                for (const node of nodes) {
                    if (!node || !hasMoveFromAncestor(node)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            "A move source contains deleted text. Microsoft Word offers to recover the file. " +
                            "Use w:t inside w:moveFrom, not w:delText.",
                        path: xmlFile,
                        code: "word-move-from-deleted-text",
                    });
                }
            }
        }

        return issues;
    },
};
