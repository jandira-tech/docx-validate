/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Content-model placement of section properties, as Word sees it.
 *
 * Replaces the `sectPr` entry of `WORD_BLOCKING_MISPLACED_LOCALS`
 * (validators/docx.ts) with a parent-aware DOM check that needs no libxml2
 * diagnostic.
 *
 * Word evidence: the rejection class is documented at corpus level —
 * `word-invalid-fixtures/README.md` ("Misplaced content-model elements (most
 * OPEN_ERROR + table UNREADABLE): a `<w:r>` / `<w:rPr>` / `<w:sectPr>` ...
 * sitting where the schema forbids it") over the 1380-file labeled probe
 * (CLEAN-FP 0/1304, BAD-COVERED 61/70). No vendored fixture isolates a
 * misplaced `w:sectPr`, so the independence test below builds the defect
 * synthetically; the tolerated control is the legal `w:sectPr` placements —
 * the last child of `w:body` and inside `w:pPr` — that every working fixture
 * carries (e.g. word-tolerated-misplaced-pgsz.docx, whose sectPr children
 * Word accepts even out of order).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { type WmlPartDocument, WORDPROCESSINGML_NAMESPACE as W, wmlPartDocuments } from "./wml-parts";

/**
 * Local names of the parents that may hold a `w:sectPr` directly (CT_Body,
 * CT_DocPartBody, CT_PPr, CT_SectPrChange). Pinned against the bundled
 * wml.xsd by `tests/word-content-model.integration.test.ts`.
 */
export const SECTION_PROPERTIES_PARENTS: ReadonlySet<string> = new Set(["body", "docPartBody", "pPr", "sectPrChange"]);

const sectionPropertiesIssues = ({ part, document: doc }: WmlPartDocument): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const properties of getElementsByTagNameNSAll(doc, W, "sectPr")) {
        const parent = properties.parentNode;
        if (parent?.nodeType !== 1) {
            continue;
        }
        const parentElement = parent as Element;
        if (parentElement.namespaceURI !== W || SECTION_PROPERTIES_PARENTS.has(parentElement.localName ?? "")) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordSectionPropertiesMisplacedCheck.id,
            path: part,
            message:
                `w:sectPr sits directly inside w:${parentElement.localName}; section properties are admitted only inside ` +
                `${[...SECTION_PROPERTIES_PARENTS]
                    .sort()
                    .map((local) => `w:${local}`)
                    .join(", ")} — as the last block of the body or inside paragraph properties. ` +
                "Microsoft Word refuses to open section properties outside their containers.",
        });
    }
    return issues;
};

export const wordSectionPropertiesMisplacedCheck: CheckDefinition = {
    id: "word-section-properties-misplaced",
    description: "Section properties sit in the body, in paragraph properties, or in a section-properties revision.",
    groups: ["story"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => wmlPartDocuments(ctx).flatMap(sectionPropertiesIssues),
};
