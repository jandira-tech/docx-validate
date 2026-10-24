/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Grid-column attribute check (Task 8 family 8d; lane issues #209, #219,
 * #224, #226, #233, #240).
 *
 * `CT_TblGridCol` (wml.xsd lines 2268-2270) declares exactly one attribute:
 * `w:w`, typed `s:ST_TwipsMeasure`. The jubarte 0.13.0 lane writes
 * `w:type="dxa"` on every grid column — inherited from corpus documents whose
 * own validation never surfaced it — so strict schema validation rejects each
 * occurrence with "attribute 'type' is not allowed" while Word opens the file
 * regardless. A grid column's width is already implicitly twips, so no unit
 * attribute exists. This check reports every attribute the schema does not
 * declare, keeping the `w:w` value visible in the finding.
 *
 * `w:gridCol` appears in `w:tblGrid` (live table grids) and inside
 * `w:tblGridChange/w:tblGrid` (tracked table-grid revisions): both are scanned
 * across every story part. Attributes in foreign namespaces stay out of scope
 * — markup-compatibility and extension namespaces are handled by their own
 * processing — but an unprefixed attribute is reported too, because no
 * unprefixed attribute is declared on `CT_TblGridCol` either.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

export const wordGridColumnAttributesCheck: CheckDefinition = {
    id: "word-grid-column-attributes",
    description: "Every w:gridCol in a Word story carries only the w:w attribute that CT_TblGridCol declares.",
    groups: ["tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) continue;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const gridCol of getElementsByTagNameNSAll(parsed.document, ns, "gridCol")) {
                    issues.push(...gridColumnIssues(gridCol, ns, part));
                }
            }
        }
        return issues;
    },
};

/** One issue per attribute `CT_TblGridCol` does not declare. */
const gridColumnIssues = (gridCol: Element, ns: string, part: string): ValidationIssue[] => {
    const prefix = gridCol.prefix ? `${gridCol.prefix}:` : "";
    const width = gridCol.getAttributeNS(ns, "w");
    const widthNote =
        width === null || width === undefined
            ? "its only declared attribute is the optional w:w"
            : `the only declared attribute is w:w (here w:w="${width}")`;
    const issues: ValidationIssue[] = [];
    for (let i = 0; i < gridCol.attributes.length; i += 1) {
        const attr = gridCol.attributes[i]!;
        if (attr.nodeName.startsWith("xmlns")) continue;
        const isDeclaredWidth = attr.namespaceURI === ns && attr.localName === "w";
        const inWordprocessingNamespace = attr.namespaceURI === null || WORD_PARAGRAPH_NAMESPACES.includes(attr.namespaceURI);
        if (isDeclaredWidth || !inWordprocessingNamespace) continue;
        issues.push({
            severity: "warning",
            code: "word-grid-column-attributes",
            path: part,
            message:
                `${prefix}gridCol carries the attribute ${prefix}${attr.localName}="${attr.value}", which CT_TblGridCol ` +
                `does not allow; ${widthNote}. Drop ${prefix}${attr.localName}: a grid column's width is already ` +
                `measured in twips (ST_TwipsMeasure), so no unit attribute exists.`,
        });
    }
    return issues;
};
