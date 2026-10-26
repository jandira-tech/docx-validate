/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A w:footnoteReference whose w:id is not a w:footnote w:id makes
 * Word offer recovery (fixture 1829). One such reference is enough.
 * Removing it opens clean once text-box deleted text is also gone.
 * A reference that names a present footnote opens clean. Duplicate
 * footnote ids are not this defect. The same mechanism judges endnotes:
 * Word session 58 (s58a) offers recovery for a reference to endnote 1 when
 * the part holds only endnote 5; the matching-id calibration opens clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const noteReferenceMissingCheck = (kind: "footnote" | "endnote"): CheckDefinition => ({
    id: `word-${kind}-reference-missing`,
    description: `${kind === "footnote" ? "Footnote" : "Endnote"} references name ${kind}s present in the package.`,
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        const defined = new Set<string>();
        for (const xmlFile of ctx.partNames) {
            if (path.posix.basename(xmlFile) !== `${kind}s.xml`) {
                continue;
            }
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const notes = getElementsByTagNameNSAll(dom, ns, kind);
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
            // Word does not follow endnote references from inside a footnote
            // (session 16's clean-open calibration); keep that source unjudged.
            if (kind === "endnote" && dom.documentElement?.localName === "footnotes") {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const refs = getElementsByTagNameNSAll(dom, ns, `${kind}Reference`);
                for (const entry of refs) {
                    const id = entry?.getAttributeNS(ns, "id");
                    if (!id || defined.has(id)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `${kind === "footnote" ? "A" : "An"} ${kind} reference uses id ${id}, and this package has no w:${kind} with that id. ` +
                            "Microsoft Word offers to recover the file. " +
                            `Remove the w:${kind}Reference, or add the matching ${kind}.`,
                        path: xmlFile,
                        code: `word-${kind}-reference-missing`,
                    });
                }
            }
        }

        return issues;
    },
});

export const wordFootnoteReferenceMissingCheck = noteReferenceMissingCheck("footnote");
export const wordEndnoteReferenceMissingCheck = noteReferenceMissingCheck("endnote");
