/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Content-model placement of table-property elements, as Word sees it.
 *
 * Replaces the `tblBorders` and `tblCellSpacing` entries of
 * `WORD_BLOCKING_MISPLACED_LOCALS` (validators/docx.ts) with namespace- and
 * parent-aware DOM checks that need no libxml2 diagnostic.
 *
 * Word evidence:
 * - Rejected, out-of-order table properties in STORY parts:
 *   `tests/fixtures/external/docx-parser-converter/tagged_tests/table_tests_v2.docx`
 *   writes `w:tblBorders` after `w:tblLayout` (12 tables) and real Word
 *   answered unreadable content — the probe copy is vendored as
 *   `word-invalid-fixtures/UNREADABLE_CONTENT/tests__fixtures__external__
 *   docx-parser-converter__tagged_tests__table_tests_v2.docx`. The probed
 *   `table_advanced_demo.docx` (same directory tree) carries the
 *   `w:tblCellSpacing`-after-`w:tblBorders` variant, as does
 *   `tests/fixtures/broken/word-rejects-table-cellspacing.docx`.
 * - TOLERATED, deliberately NOT reported:
 *   - the jubarte 0.13.0 lane specimens `tests/lane/jubarte-0.13.0/k07_*`,
 *     `k08_*`, `k09_*` put `w:tblBorders` directly inside `w:tcPr` (within
 *     `w:tblStylePr` in word/styles.xml) and real Word opened and exported
 *     every one of them (lane manifest `word_opened_and_exported`).
 *   - `tests/fixtures/working/package.word-row-without-cell-repaired.docx`
 *     carries `w:tblBorders` after `w:tblCellMar` in word/styles.xml and is
 *     the recorded "one cell opens clean" control of fixture 1993
 *     (word-row-without-cell). Word's style handler tolerates table-property
 *     order quirks in styles parts, so the order arm fires in story parts
 *     only; wml.xsd admits neither quirk, so the tolerated set is
 *     intentionally WIDER than the schema.
 * - Rejected, `w:tblCellSpacing` under a parent that cannot hold it:
 *   the word-valid XSD walk test builds it under `w:sectPr`. Controls:
 *   `w:tblCellSpacing` ordered before `w:tblBorders` inside `w:tblPr`/
 *   `w:tblPrEx`, and anywhere inside `w:trPr` (a choice model).
 *
 * No corpus specimen puts either element under an outright illegal direct
 * parent; the parent arm below follows the corpus-probe rejection class
 * documented in `word-invalid-fixtures/README.md` ("a ... table-property
 * element ... sitting where the schema forbids it", CLEAN-FP 0/1304).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { type WmlPartDocument, WORDPROCESSINGML_NAMESPACE as W, wmlPartDocuments } from "./wml-parts";

/**
 * Parents that may hold a `w:tblBorders` directly: the schema containers
 * (`w:tblPr`, `w:tblPrEx` — CT_TblPrBase) plus the Word-tolerated `w:tcPr`
 * documented above. Pinned by `tests/word-content-model.integration.test.ts`.
 */
export const TABLE_BORDERS_PARENTS: ReadonlySet<string> = new Set(["tblPr", "tblPrEx", "tcPr"]);

/** Parents that may hold a `w:tblCellSpacing` directly (CT_TblPrBase, CT_TrPr). */
export const TABLE_CELL_SPACING_PARENTS: ReadonlySet<string> = new Set(["tblPr", "tblPrEx", "trPr"]);

/**
 * The CT_TblPrBase child sequence, in order. `w:tblCellSpacing` (position 8)
 * must precede every element that follows it, and `w:tblBorders` (position
 * 10) must precede `w:shd` and everything after it; pinned against wml.xsd by
 * `tests/word-content-model.integration.test.ts`.
 */
export const TBLPRBASE_ORDER: readonly string[] = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
];

/** CT_TblPrBase elements that must follow `element` in the sequence. */
const followers = (element: string): ReadonlySet<string> => new Set(TBLPRBASE_ORDER.slice(TBLPRBASE_ORDER.indexOf(element) + 1));

const AFTER_TABLE_BORDERS = followers("tblBorders");
const AFTER_TABLE_CELL_SPACING = followers("tblCellSpacing");

/** The closest preceding CT_TblPrBase element that must follow `element`, if any. */
const precedingFollower = (element: Element, after: ReadonlySet<string>): string | null => {
    for (let sibling = element.previousSibling; sibling; sibling = sibling.previousSibling) {
        if (sibling.nodeType !== 1 || (sibling as Element).namespaceURI !== W) {
            continue;
        }
        const local = (sibling as Element).localName ?? "";
        if (after.has(local)) {
            return local;
        }
    }
    return null;
};

const tablePropertyIssues = ({ part, document: doc }: WmlPartDocument, storyParts: ReadonlySet<string>): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    const borderIssues = getElementsByTagNameNSAll(doc, W, "tblBorders");
    const spacingIssues = getElementsByTagNameNSAll(doc, W, "tblCellSpacing");
    for (const [local, elements, parents, after, allowedOrders] of [
        ["w:tblBorders", borderIssues, TABLE_BORDERS_PARENTS, AFTER_TABLE_BORDERS, ["tblPr", "tblPrEx"]],
        ["w:tblCellSpacing", spacingIssues, TABLE_CELL_SPACING_PARENTS, AFTER_TABLE_CELL_SPACING, ["tblPr", "tblPrEx"]],
    ] as const) {
        for (const element of elements) {
            const parent = element.parentNode;
            if (parent?.nodeType !== 1) {
                continue;
            }
            const parentElement = parent as Element;
            if (parentElement.namespaceURI !== W) {
                continue;
            }
            const parentLocal = parentElement.localName ?? "";
            if (!parents.has(parentLocal)) {
                issues.push({
                    severity: "error",
                    code: wordTablePropertiesMisplacedCheck.id,
                    path: part,
                    message:
                        `${local} sits directly inside w:${parentLocal}; it is admitted only inside ` +
                        `${[...parents]
                            .sort()
                            .map((container) => `w:${container}`)
                            .join(", ")}${local === "w:tblBorders" ? " — with w:tcPr tolerated because Word opened it there" : ""}. ` +
                        "Microsoft Word refuses to open a table-property element outside its containers.",
                });
                continue;
            }
            if (!(allowedOrders as readonly string[]).includes(parentLocal) || !storyParts.has(part)) {
                continue;
            }
            const follower = precedingFollower(element, after);
            if (follower !== null) {
                issues.push({
                    severity: "error",
                    code: wordTablePropertiesMisplacedCheck.id,
                    path: part,
                    message:
                        `${local} follows w:${follower} inside w:${parentLocal}; CT_TblPrBase orders it ` +
                        `before ${[...after].map((later) => `w:${later}`).join(", ")}. ` +
                        "Microsoft Word refuses to open a table with reordered table properties.",
                });
            }
        }
    }
    return issues;
};

export const wordTablePropertiesMisplacedCheck: CheckDefinition = {
    id: "word-table-properties-misplaced",
    description:
        "Table borders and table cell spacing sit in their schema containers and their CT_TblPrBase order; borders under tcPr stay tolerated as Word-opened evidence shows.",
    groups: ["tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const storyParts = new Set(ctx.storyParts());
        return wmlPartDocuments(ctx).flatMap((partDocument) => tablePropertyIssues(partDocument, storyParts));
    },
};
