/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Textbox-content namespace checks over the Word story parts. */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/**
 * The three textbox models pin the namespace of the content element:
 *
 * - `wp:txbx` (wordprocessingDrawing `wsp`) requires `wp:txbxContent` in the
 *   same namespace — Transitional or Strict `…drawingml/…wordprocessingDrawing`.
 * - `wps:txbx` (Word 2010 `wordprocessingShape`) requires the WML
 *   `w:txbxContent`.
 * - `v:textbox` (VML) likewise requires `w:txbxContent`.
 *
 * Word's Strict writer emits the DrawingML text body in its own extension
 * namespace (`wne:txbxContent`), which survives Strict→Transitional
 * conversion verbatim (jubarte lane k12, issue #244): the body no longer
 * satisfies `CT_TextboxInfo`, and any MCE-conforming consumer that honors
 * `wne` inside `mc:Ignorable` sees an empty textbox.
 */
const WORDPROCESSING_DRAWING_NAMESPACES: readonly string[] = [
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
];
const WORDPROCESSING_SHAPE_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const VML_NAMESPACE = "urn:schemas-microsoft-com:vml";

const elementChildren = (parent: Element, localName: string): Element[] => {
    const out: Element[] = [];
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 1 && (child as Element).localName === localName) {
            out.push(child as Element);
        }
    }
    return out;
};

const checkTextbox = (txbx: Element, requiredNamespaces: readonly string[], part: string, issues: ValidationIssue[]): void => {
    const parentName = txbx.tagName;
    for (const content of elementChildren(txbx, "txbxContent")) {
        const actual = content.namespaceURI;
        if (actual !== null && requiredNamespaces.includes(actual)) {
            continue;
        }
        issues.push({
            severity: "warning",
            code: "textbox-content-namespace",
            path: part,
            message:
                `${parentName} carries its text body as ${content.tagName} in namespace "${actual ?? ""}", ` +
                `but the schema requires txbxContent in "${requiredNamespaces[0]}". ` +
                `Rename the element into that namespace; the w:txbxContent spelling belongs to the ` +
                `wps:txbx and VML v:textbox models, and other spellings satisfy no model.`,
        });
    }
};

export const textboxContentNamespaceCheck: CheckDefinition = {
    id: "textbox-content-namespace",
    description: "A textbox's txbxContent sits in the namespace its textbox model declares.",
    groups: ["drawings"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.xml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const ns of WORDPROCESSING_DRAWING_NAMESPACES) {
                for (const txbx of getElementsByTagNameNSAll(parsed.document, ns, "txbx")) {
                    checkTextbox(txbx, [ns], part, issues);
                }
            }
            for (const txbx of getElementsByTagNameNSAll(parsed.document, WORDPROCESSING_SHAPE_NAMESPACE, "txbx")) {
                checkTextbox(txbx, WORD_PARAGRAPH_NAMESPACES, part, issues);
            }
            for (const textbox of getElementsByTagNameNSAll(parsed.document, VML_NAMESPACE, "textbox")) {
                checkTextbox(textbox, WORD_PARAGRAPH_NAMESPACES, part, issues);
            }
        }
        return issues;
    },
};
