/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Required and ordered `w:tbl` children. `CT_Tbl` is an `xsd:sequence`, so a
 * table is legal only as range markup, then `w:tblPr`, then `w:tblGrid`, then
 * the rows. Word falls back to table defaults when `w:tblPr` is absent, so the
 * lane's header tables written as `w:tbl > w:tblGrid > w:tr` open fine; the
 * finding is an advisory and strict schema validation is what rejects them.
 *
 * Schema rule, quoted from `ISO-IEC29500-4_2016/wml.xsd:2532-2539`:
 *
 *     <xsd:complexType name="CT_Tbl">
 *       <xsd:sequence>
 *         <xsd:group ref="EG_RangeMarkupElements" minOccurs="0" maxOccurs="unbounded"/>
 *         <xsd:element name="tblPr" type="CT_TblPr"/>
 *         <xsd:element name="tblGrid" type="CT_TblGrid"/>
 *         <xsd:group ref="EG_ContentRowContent" minOccurs="0" maxOccurs="unbounded"/>
 *       </xsd:sequence>
 *     </xsd:complexType>
 *
 * `tblPr` and `tblGrid` carry no `minOccurs`, so both are required.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll, getLocalName } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/** `EG_RangeMarkupElements` (wml.xsd:1034-1053) — the only children allowed before `w:tblPr`. */
const RANGE_MARKUP: ReadonlySet<string> = new Set([
    "bookmarkStart",
    "bookmarkEnd",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "commentRangeStart",
    "commentRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
]);

const SEQUENCE_RULE =
    "CT_Tbl in ISO/IEC 29500-4 (wml.xsd:2532) is an xsd:sequence of optional range markup, a required w:tblPr, " +
    "a required w:tblGrid, then the rows.";
const ADD_TBL_PR = "Insert w:tblPr as the first child; an empty <w:tblPr/> satisfies CT_TblPr, whose every CT_TblPrBase child is optional.";
const ADD_TBL_GRID = "Add a w:tblGrid directly after w:tblPr holding one w:gridCol per table column.";

/** Direct element children, in document order. */
const childElements = (element: Element): readonly Element[] => {
    const children: Element[] = [];
    for (let child = element.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 1) {
            children.push(child as Element);
        }
    }
    return children;
};

/**
 * Walk one table's children. `w:tblPr` and `w:tblGrid` must both appear, each
 * once, in that order and before the first row. A required child that shows up
 * late is reported once, as misplaced, rather than also as missing.
 */
const scanTable = (tbl: Element, ns: string): readonly ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    let seenPr = false;
    let seenGrid = false;
    let rows = false;
    /** A required child that has been skipped over and may still turn up late. */
    let pendingPr = false;
    let pendingGrid = false;
    /** Last significant child name, so a misplaced child can name its predecessor. */
    let last = "";

    const report = (problem: string, guidance: string): void => {
        issues.push({
            severity: "warning",
            code: "word-table-properties-required",
            path: "",
            message: `w:tbl: ${problem}. ${SEQUENCE_RULE} ${guidance}`,
        });
    };

    for (const child of childElements(tbl)) {
        if (child.namespaceURI !== ns) {
            continue;
        }
        const name = getLocalName(child);
        if (RANGE_MARKUP.has(name)) {
            continue;
        }
        if (name === "tblPr") {
            if (seenPr) {
                report(
                    last === name ? "w:tblPr appears twice" : `w:tblPr appears after w:${last}`,
                    "Keep exactly one w:tblPr, as the first child.",
                );
                continue;
            }
            seenPr = true;
            if (pendingPr) {
                report(`w:tblPr appears after w:${last}`, "Keep exactly one w:tblPr, as the first child.");
                pendingPr = false;
            }
            last = name;
            continue;
        }
        if (name === "tblGrid") {
            if (seenGrid) {
                report(
                    last === name ? "w:tblGrid appears twice" : `w:tblGrid appears after w:${last}`,
                    "Keep exactly one w:tblGrid, directly after w:tblPr.",
                );
                continue;
            }
            seenGrid = true;
            if (pendingGrid) {
                report(`w:tblGrid appears after w:${last}`, "Keep exactly one w:tblGrid, directly after w:tblPr and before the rows.");
                pendingGrid = false;
            } else if (!seenPr) {
                // w:tblGrid landed before any w:tblPr; a later one is misplaced.
                pendingPr = true;
            }
            last = name;
            continue;
        }
        // Row content: nothing after it can still be in its required position.
        if (!rows) {
            rows = true;
            pendingPr = pendingPr || !seenPr;
            pendingGrid = pendingGrid || !seenGrid;
        }
        last = name;
    }

    if (!seenPr) {
        report(`w:tblPr is missing${seenGrid ? " — w:tblGrid appears before it" : ""}`, ADD_TBL_PR);
    }
    if (!seenGrid) {
        report("w:tblGrid is missing", ADD_TBL_GRID);
    }
    return issues;
};

export const wordTablePropertiesRequiredCheck: CheckDefinition = {
    id: "word-table-properties-required",
    description: "Every w:tbl carries its required w:tblPr and w:tblGrid in the CT_Tbl order: range markup, w:tblPr, w:tblGrid, rows.",
    groups: ["tables"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.xml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                if (parsed.document.documentElement?.namespaceURI !== ns) {
                    continue;
                }
                for (const tbl of getElementsByTagNameNSAll(parsed.document, ns, "tbl")) {
                    for (const issue of scanTable(tbl, ns)) {
                        issues.push({ ...issue, path: part });
                    }
                }
            }
        }
        return issues;
    },
};
