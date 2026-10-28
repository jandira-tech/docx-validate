/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word compatibility checks migrated without changing their issue details.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";

export const wordTextContainsParagraphCheck: CheckDefinition = {
    id: "word-text-contains-paragraph",
    description: "Text elements contain no nested paragraphs.",
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
                const texts = getElementsByTagNameNSAll(dom, ns, "t");
                for (const text of texts) {
                    const nested = getElementsByTagNameNSAll(text, ns, "p");
                    if (nested.length === 0) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `A <w:t> element contains ${nested.length} <w:p> element(s). ` +
                            "Microsoft Word cannot open the file. Move the paragraph out of the text run.",
                        path: xmlFile,
                        code: "word-text-contains-paragraph",
                    });
                }
            }
        }

        return issues;
    },
};
