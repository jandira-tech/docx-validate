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
 *
 * The defined notes come only from the parts a `footnotes`
 * (respectively `endnotes`) relationship of the main document, or of a
 * reached glossary, names, whatever the part is called: Word session 60
 * (s60a) opened clean a used notes part stored as `word/notes-foot.xml`
 * and `word/notes-end.xml` and related by its relationship, which the
 * name-only reading rejected (#793), and session 61 (s61a, s61b) offered
 * recovery for a body reference whose note 1 sat only in a
 * `word/footnotes.xml` / `word/endnotes.xml` no relationship names — alone
 * or beside a related part holding note 5 — while the same unrelated part
 * with nothing referencing it opened clean: Word defines notes only from
 * the parts their relationships name, so the former basename reading is
 * gone (#797). The relationship walks are ungated by usage, so a
 * related-but-unused part defines its ids too — this set only adds
 * definitions, never a rejection. The union over the main document's and
 * reached glossaries' relationships is unprobed for a main-body reference
 * satisfied only by a glossary's notes part.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { glossaryTypedRelationshipTargets, wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const noteReferenceMissingCheck = (kind: "footnote" | "endnote"): CheckDefinition => ({
    id: `word-${kind}-reference-missing`,
    description: `${kind === "footnote" ? "Footnote" : "Endnote"} references name ${kind}s present in the package.`,
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        const defined = new Set<string>();
        const collect = (part: string): void => {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                return;
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
        };
        // A note is defined only by the part its relationship names,
        // whatever the part is called, from the main document and from a
        // reached glossary alike (session 60 s60a, #793); a part no
        // relationship names defines nothing, even under the standard
        // names (session 61 s61a, s61b, #797). The walks are ungated by
        // usage, so a related-but-unused part defines its ids too, as this
        // set only adds definitions, never a rejection.
        const segment = `${kind}s`;
        for (const target of await wordTypedRelationshipTargets(ctx, new Set([segment]), profile)) {
            collect(target.part);
        }
        for (const target of await glossaryTypedRelationshipTargets(ctx, new Set([segment]), profile)) {
            collect(target.part);
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
