/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A w:sdtPr/w:id with w:val="" makes Word refuse the file (fixture
 * 1830). Setting that value to a number opens clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";

export const wordSdtIdEmptyCheck: CheckDefinition = {
    id: "word-sdt-id-empty",
    description: "Content control ids have no empty values.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const ids = getElementsByTagNameNSAll(dom, ns, "id");
                for (const id of ids) {
                    const parent = id.parentNode;
                    if (!parent || parent.nodeType !== 1) {
                        continue;
                    }
                    const parentElem = parent as Element;
                    const parentLocal = parentElem.localName || parentElem.tagName.split(":").pop() || "";
                    if (parentLocal !== "sdtPr" || !WORD_PARAGRAPH_NAMESPACES.includes(parentElem.namespaceURI ?? "")) {
                        continue;
                    }
                    if (id.getAttributeNS(ns, "val") !== "") {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message: 'A w:sdt has w:id w:val="". Microsoft Word cannot open the file. Set w:val to a number.',
                        path: xmlFile,
                        code: "word-sdt-id-empty",
                    });
                }
            }
        }

        return issues;
    },
};
