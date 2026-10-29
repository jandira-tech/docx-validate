/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Content-model placement of paragraph-properties blocks, as Word sees it.
 *
 * Replaces the `}pPr': This element is not expected` + `"Expected is one of"`
 * message match in `isWordBlockingXsdIssue` (validators/docx.ts) with a
 * parent-aware DOM check that needs no libxml2 diagnostic.
 *
 * Word evidence:
 * - Rejected: `tests/fixtures/external/open-xml-sdk/5Errors.docx` (manifest
 *   `word: unreadable-content-warning`) nests `w:pPr` INSIDE `w:pPr`, where
 *   CT_PPrBase admits no paragraph-properties block; Word offers recovery.
 * - Tolerated control: `tests/fixtures/working/word-tolerated-duplicate-ppr.docx`
 *   — two `w:pPr` children of ONE `w:p`. The schema also rejects that shape,
 *   but real Word silently merges the duplicate and opens the file cleanly,
 *   so a duplicate under `w:p` must never fire; only a `w:pPr` whose parent
 *   cannot hold one does.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { type WmlPartDocument, WORDPROCESSINGML_NAMESPACE as W, wmlPartDocuments } from "./wml-parts";

/**
 * Local names of the parents that may hold a `w:pPr` directly (CT_P, CT_Lvl,
 * CT_Style, CT_TblStylePr, CT_PPrDefault, CT_PPrChange). Pinned against the
 * bundled wml.xsd by `tests/word-content-model.integration.test.ts`.
 */
export const PARAGRAPH_PROPERTIES_PARENTS: ReadonlySet<string> = new Set(["lvl", "p", "pPrChange", "pPrDefault", "style", "tblStylePr"]);

const paragraphPropertiesIssues = ({ part, document: doc }: WmlPartDocument): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const properties of getElementsByTagNameNSAll(doc, W, "pPr")) {
        const parent = properties.parentNode;
        if (parent?.nodeType !== 1) {
            continue;
        }
        const parentElement = parent as Element;
        if (parentElement.namespaceURI !== W || PARAGRAPH_PROPERTIES_PARENTS.has(parentElement.localName ?? "")) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordParagraphPropertiesMisplacedCheck.id,
            path: part,
            message:
                `w:pPr sits directly inside w:${parentElement.localName}; paragraph properties are admitted only inside ` +
                `${[...PARAGRAPH_PROPERTIES_PARENTS]
                    .sort()
                    .map((local) => `w:${local}`)
                    .join(", ")}. ` +
                "Microsoft Word offers recovery for paragraph properties outside their containers; a duplicate w:pPr in one w:p is tolerated and never reported.",
        });
    }
    return issues;
};

export const wordParagraphPropertiesMisplacedCheck: CheckDefinition = {
    id: "word-paragraph-properties-misplaced",
    description: "Paragraph-properties blocks sit in their schema parents; duplicates inside one paragraph are tolerated.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => wmlPartDocuments(ctx).flatMap(paragraphPropertiesIssues),
};
