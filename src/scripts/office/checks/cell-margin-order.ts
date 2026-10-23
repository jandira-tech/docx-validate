/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Cell-margin child order. `w:tblCellMar` (`CT_TblCellMar`, wml.xsd:2458) and
 * `w:tcMar` (`CT_TcMar`, wml.xsd:2299) are `xsd:sequence` types that declare
 * the same six margins. A margin serialized after a later sibling makes the
 * part schema-invalid; Word reads margins by name, so it opens such a file and
 * the finding stays an advisory.
 *
 * Schema sequence, quoted from `ISO-IEC29500-4_2016/wml.xsd:2458-2467`:
 *
 *     <xsd:complexType name="CT_TblCellMar">
 *       <xsd:sequence>
 *         <xsd:element name="top" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *         <xsd:element name="start" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *         <xsd:element name="left" type="CT_TblWidth" minOccurs="0"/>
 *         <xsd:element name="bottom" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *         <xsd:element name="end" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *         <xsd:element name="right" type="CT_TblWidth" minOccurs="0"/>
 *       </xsd:sequence>
 *     </xsd:complexType>
 *
 * Every slot is optional, so omitting margins is legal; only a child that
 * follows a later sibling is reported.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { getElementsByTagNameNSAll, getLocalName } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/** The six sequence slots, in the order the bundled XSD declares them. */
const MARGIN_SEQUENCE = ["top", "start", "left", "bottom", "end", "right"] as const;
const SLOT_OF = new Map<string, number>(MARGIN_SEQUENCE.map((name, index) => [name, index]));
const EXPECTED_ORDER = MARGIN_SEQUENCE.join(", ");

/**
 * Margin element to the parents the schema allows it in: `CT_TblPrBase` /
 * `CT_TblPrExBase` carry `w:tblCellMar` (wml.xsd:2495, :2519) and
 * `CT_TcPrBase` carries `w:tcMar` (wml.xsd:2331).
 */
const MARGIN_PARENTS: Readonly<Record<string, readonly string[]>> = {
    tblCellMar: ["tblPr", "tblPrEx"],
    tcMar: ["tcPr"],
};

/**
 * Parts that can carry the construct: every Word story, plus `word/styles.xml`,
 * where `CT_Style/w:tblPr` (wml.xsd:3256) and `CT_TblStylePr/w:tblPr`
 * (wml.xsd:3218) hold the same margin elements.
 */
const STYLES_PART = "word/styles.xml";

const marginParts = (ctx: ValidationContext): readonly string[] =>
    [...ctx.storyParts(), STYLES_PART].filter((part, index, all) => ctx.has(part) && all.indexOf(part) === index);

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

/** Report every margin child that follows a sibling from a later slot. */
const outOfOrderChildren = (margin: Element, ns: string): readonly ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    const container = getLocalName(margin);
    let latestSlot = -1;
    let latestName = "";
    for (const child of childElements(margin)) {
        if (child.namespaceURI !== ns) {
            continue;
        }
        const name = getLocalName(child);
        const slot = SLOT_OF.get(name);
        if (slot === undefined) {
            continue;
        }
        if (slot < latestSlot) {
            issues.push({
                severity: "warning",
                code: "word-cell-margin-order",
                path: "",
                message:
                    `w:${name} follows w:${latestName} inside w:${container}, but the margins form an xsd:sequence: ` +
                    `${EXPECTED_ORDER}. Emit them in that order (Word's own transitional output is top, left, bottom, right) ` +
                    `or leave a margin out and keep the inherited default.`,
            });
            continue;
        }
        latestSlot = slot;
        latestName = name;
    }
    return issues;
};

export const wordCellMarginOrderCheck: CheckDefinition = {
    id: "word-cell-margin-order",
    description:
        "Every w:tblCellMar and w:tcMar lists its margins in the CT_TblCellMar / CT_TcMar sequence order top, start, left, bottom, end, right.",
    groups: ["tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of marginParts(ctx)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                if (parsed.document.documentElement?.namespaceURI !== ns) {
                    continue;
                }
                for (const [container, parents] of Object.entries(MARGIN_PARENTS)) {
                    for (const margin of getElementsByTagNameNSAll(parsed.document, ns, container)) {
                        const parent = margin.parentNode as Element | null;
                        if (parent?.namespaceURI !== ns || !parents.includes(getLocalName(parent))) {
                            continue;
                        }
                        for (const issue of outOfOrderChildren(margin, ns)) {
                            issues.push({ ...issue, path: part });
                        }
                    }
                }
            }
        }
        return issues;
    },
};
