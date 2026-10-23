/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A w:footnoteReference whose w:id is not a w:footnote w:id makes
 * Word offer recovery (fixture 1829). One such reference is enough.
 * Removing it opens clean once text-box deleted text is also gone.
 * A reference that names a present footnote opens clean. Duplicate
 * footnote ids are not this defect.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

export const wordFootnoteReferenceMissingCheck: CheckDefinition = {
    id: "word-footnote-reference-missing",
    description: "Footnote references name footnotes present in the package.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        const defined = new Set<string>();
        for (const xmlFile of ctx.partNames) {
            if (path.posix.basename(xmlFile) !== "footnotes.xml") {
                continue;
            }
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const notes = getElementsByTagNameNSAll(dom, ns, "footnote");
                for (const entry of notes) {
                    const id = entry?.getAttributeNS(ns, "id");
                    if (id) {
                        defined.add(id);
                    }
                }
            }
        }
        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const refs = getElementsByTagNameNSAll(dom, ns, "footnoteReference");
                for (const entry of refs) {
                    const id = entry?.getAttributeNS(ns, "id");
                    if (!id || defined.has(id)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `A footnote reference uses id ${id}, and this package has no w:footnote with that id. ` +
                            "Microsoft Word offers to recover the file. " +
                            "Remove the w:footnoteReference, or add the matching footnote.",
                        path: xmlFile,
                        code: "word-footnote-reference-missing",
                    });
                }
            }
        }

        return issues;
    },
};
