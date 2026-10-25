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

const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const localNameOf = (elem: Element): string => elem.localName || elem.tagName.split(":").pop() || "";

const isWordElement = (elem: Element, local: string): boolean =>
    localNameOf(elem) === local && WORD_PARAGRAPH_NAMESPACES.includes(elem.namespaceURI ?? "");

interface TextBoxFrame {
    node: Node;
    inTextBox: boolean;
    /** A w:del encloses the node inside its innermost text box. */
    deleted: boolean;
    inFallback: boolean;
}

/**
 * The w:delText and w:delInstrText elements inside a w:txbxContent that no
 * w:del within that text box wraps, outside any mc:Fallback. One walk
 * carries the enclosing state down the tree; walking every deleted-text
 * element's ancestors took 2.3-3.6x as long (bot PR #199).
 */
const textBoxDeletedText = (dom: Document): Element[] => {
    const found: Element[] = [];
    const stack: TextBoxFrame[] = [{ node: dom, inTextBox: false, deleted: false, inFallback: false }];
    while (stack.length > 0) {
        const frame = stack.pop();
        if (frame === undefined) break;
        let { inTextBox, deleted, inFallback } = frame;
        if (frame.node.nodeType === 1 /* ELEMENT_NODE */) {
            const elem = frame.node as Element;
            if (isWordElement(elem, "txbxContent")) {
                inTextBox = true;
                deleted = false;
            } else if (isWordElement(elem, "del")) {
                deleted = true;
            } else if (localNameOf(elem) === "Fallback" && elem.namespaceURI === MC_NAMESPACE) {
                inFallback = true;
            } else if (inTextBox && !deleted && !inFallback && (isWordElement(elem, "delText") || isWordElement(elem, "delInstrText"))) {
                found.push(elem);
            }
        }
        for (let child = frame.node.lastChild; child; child = child.previousSibling) {
            if (child.nodeType === 1 /* ELEMENT_NODE */) stack.push({ node: child, inTextBox, deleted, inFallback });
        }
    }
    return found;
};

/** The w:delText elements inside a w:moveFrom, found in one walk. */
const moveFromDeletedText = (dom: Document): Element[] => {
    const found: Element[] = [];
    const stack: [Node, boolean][] = [[dom, false]];
    while (stack.length > 0) {
        const frame = stack.pop();
        if (frame === undefined) break;
        const [node, enclosed] = frame;
        let inMoveFrom = enclosed;
        if (node.nodeType === 1 /* ELEMENT_NODE */) {
            const elem = node as Element;
            if (isWordElement(elem, "moveFrom")) {
                inMoveFrom = true;
            } else if (inMoveFrom && isWordElement(elem, "delText")) {
                found.push(elem);
            }
        }
        for (let child = node.lastChild; child; child = child.previousSibling) {
            if (child.nodeType === 1 /* ELEMENT_NODE */) stack.push([child, inMoveFrom]);
        }
    }
    return found;
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
            for (const _node of textBoxDeletedText(parsed.document)) {
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
            for (const _node of moveFromDeletedText(parsed.document)) {
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

        return issues;
    },
};
