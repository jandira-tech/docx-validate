/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unknown WordprocessingML elements and namespace-less `link` elements, as
 * Word sees them. Replaces the `}p2': This element is not expected` and
 * `Element 'link': This element is not expected` message matches in
 * `isWordBlockingXsdIssue` (validators/docx.ts) with DOM checks that need no
 * libxml2 diagnostic.
 *
 * Word evidence:
 * - `tests/fixtures/external/open-xml-sdk/UnknownElement.docx` (manifest
 *   `word: open-error`) carries `<w:p2>` — an element whose local name
 *   wml.xsd never declares — directly in the body; real Word answers
 *   OPEN_ERROR. The probed twins live in `word-invalid-fixtures/OPEN_ERROR/`
 *   (`tests__fixtures__unknown-element.docx`,
 *   `tests__fixtures__external__docx-validate__unknown-element.docx`).
 * - TOLERATED, deliberately NOT reported: unknown WordprocessingML-namespace
 *   names inside property containers. `annotations_import_2.docx` and
 *   `fields_attrs1.docx` (manifest `word: clean-open`) carry
 *   `<w:fieldType>`/`<w:fieldTypeShort>`/`<w:fieldColor>` in `w:sdtPr`, and
 *   `package.word-choice-requires-declared.docx` (pinned word-valid) carries
 *   `<w:textFlow>` in `w:pPr`; Word opens them all cleanly. So the check
 *   fires only when the unknown element occupies a BLOCK-LEVEL slot (body,
 *   table structure, structured-document or text-box content) — the shape
 *   Word cannot load — never inside property containers. Extension-namespace
 *   elements (w14:, w15:, m:, wp14:, ...) target their own namespaces and are
 *   not this defect at all.
 * - `tests/fixtures/external/superdoc/super-editor/hyperlink-exported.docx`
 *   (manifest `word: open-error`) emits `<link />` — an element with NO XML
 *   namespace — inside `w:rPr`; real Word answers OPEN_ERROR. The tolerated
 *   control is the namespaced `w:link` style child
 *   (`tests/fixtures/working/word-tolerated-misplaced-link.docx`, Word
 *   opens clean).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameAll } from "../../../lib/xml-helpers";
import { WORDPROCESSINGML_NAMESPACE as W, wmlPartDocuments } from "./wml-parts";
import { WML_ELEMENT_NAMES } from "./wml-vocabulary";

/**
 * Block-level slots: parents whose content model is story structure. An
 * unknown WordprocessingML element here is the shape Word refuses to load
 * (`w:p2` under `w:body`); inside property containers Word skips unknown
 * names (see the module evidence above).
 */
export const UNKNOWN_ELEMENT_BLOCK_PARENTS: ReadonlySet<string> = new Set(["body", "sdtContent", "tc", "tbl", "tr", "txbxContent"]);

const describeParent = (element: Element): string => {
    const parent = element.parentNode;
    if (parent?.nodeType !== 1) {
        return "the document root";
    }
    const parentElement = parent as Element;
    return `${parentElement.namespaceURI === W ? "w:" : ""}${parentElement.localName ?? ""}`;
};

const unknownElementIssues = (part: string, doc: Document): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const element of getElementsByTagNameAll(doc)) {
        if (element.namespaceURI !== W) {
            continue;
        }
        const local = element.localName ?? "";
        if (WML_ELEMENT_NAMES.has(local)) {
            continue;
        }
        const parent = element.parentNode;
        if (parent?.nodeType !== 1) {
            continue;
        }
        const parentElement = parent as Element;
        if (parentElement.namespaceURI !== W || !UNKNOWN_ELEMENT_BLOCK_PARENTS.has(parentElement.localName ?? "")) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordUnknownElementCheck.id,
            path: part,
            message:
                `w:${local} sits directly inside w:${parentElement.localName}; wml.xsd declares no w:${local} in the WordprocessingML namespace. ` +
                "Microsoft Word refuses to open an unknown WordprocessingML element in story structure.",
        });
    }
    return issues;
};

const linkWithoutNamespaceIssues = (part: string, doc: Document): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const element of getElementsByTagNameAll(doc)) {
        if (element.localName !== "link" || element.namespaceURI) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordLinkWithoutNamespaceCheck.id,
            path: part,
            message:
                `link appears without an XML namespace, directly inside ${describeParent(element)}; unprefixed elements are not part of any OOXML content model ` +
                "(a style link is w:link inside w:style; a hyperlink is w:hyperlink around runs). Microsoft Word refuses to open the file.",
        });
    }
    return issues;
};

export const wordUnknownElementCheck: CheckDefinition = {
    id: "word-unknown-element",
    description: "Every WordprocessingML element name is one wml.xsd declares; extension namespaces are not this defect.",
    groups: ["schema"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] =>
        wmlPartDocuments(ctx).flatMap(({ part, document }) => [
            ...unknownElementIssues(part, document),
            ...linkWithoutNamespaceIssues(part, document),
        ]),
};

export const wordLinkWithoutNamespaceCheck: CheckDefinition = {
    id: "word-link-without-namespace",
    description: "link elements carry a namespace; a namespace-less link is not part of any OOXML content model.",
    groups: ["schema"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] =>
        wmlPartDocuments(ctx).flatMap(({ part, document }) => [
            ...unknownElementIssues(part, document),
            ...linkWithoutNamespaceIssues(part, document),
        ]),
};
