/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Page-margin checks over the Word story parts. */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/**
 * Every attribute `CT_PageMar` declares `use="required"` (wml.xsd). Word
 * silently defaults the missing ones — header/footer to 720 twips, gutter
 * to 0 — so each missing attribute is an advisory, reported on its own.
 */
const REQUIRED_PAGE_MARGIN_ATTRIBUTES: readonly string[] = ["top", "right", "bottom", "left", "header", "footer", "gutter"];

/**
 * Where the `w:pgMar`'s `w:sectPr` sits, for the message. `w:sectPr` is
 * legal as the body's last child, inside `w:pPr` (a paragraph section
 * break), and inside `w:sectPrChange` (the recorded previous properties).
 */
const sectPrPlacement = (pgMar: Element): string => {
    const sectPr = pgMar.parentNode;
    const holder = sectPr?.parentNode;
    const holderName = holder?.nodeType === 1 ? (holder as Element).localName : undefined;
    if (holderName === "pPr") {
        return "the w:pPr/w:sectPr section break";
    }
    if (holderName === "sectPrChange") {
        return "the w:sectPr recorded by a w:sectPrChange";
    }
    if (holderName === "body") {
        return "the body-level w:sectPr";
    }
    return "a w:sectPr";
};

export const pageMarginAttributesCheck: CheckDefinition = {
    id: "word-page-margin-attributes",
    description: "Every w:pgMar in a w:sectPr carries the seven attributes CT_PageMar requires.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        for (const part of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const pgMar of getElementsByTagNameNSAll(parsed.document, ns, "pgMar")) {
                    for (const attribute of REQUIRED_PAGE_MARGIN_ATTRIBUTES) {
                        if (pgMar.getAttributeNS(ns, attribute) !== null) {
                            continue;
                        }
                        issues.push({
                            severity: "warning",
                            code: "word-page-margin-attributes",
                            path: part,
                            message:
                                `The w:pgMar of ${sectPrPlacement(pgMar)} is missing the required w:${attribute} attribute. ` +
                                `CT_PageMar requires it; write w:${attribute} explicitly — Word's own writers emit ` +
                                `w:header="720" w:footer="720" w:gutter="0" for the defaults.`,
                        });
                    }
                }
            }
        }
        return issues;
    },
};
