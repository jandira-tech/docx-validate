/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A comments, footnotes or endnotes part that both the main document and
 * a reached glossary read makes Word 16.115 offer to recover the file
 * (#850; sessions 66 and 67, chunks s66a and s67a in
 * docs/plans/evidence/word-probes-2026-10-24.md, two runs each, all four
 * readings agree). Session 66 (s66a) offered recovery for a
 * `word/comments.xml` the main document relates and references while a
 * reached glossary relates it (`../comments.xml`) and references it too —
 * comment 0 holding nothing but text — so the shared reach, not the note
 * reference session 65's rows added, is what Word refused; the same part
 * read by the main document alone or by the glossary alone opens clean.
 * Session 67 (s67a) settled the scope: related by both roots but
 * referenced by only one of them, or by neither, opens clean, and the
 * shared footnotes part and its endnotes twin, referenced by both roots,
 * are refused too — the refusal needs both roots to use the part, and it
 * covers notes parts.
 *
 * The read model is the whole rule. `wordReadTypedTargets` lists these
 * lazily read segments only when the root uses them (#437), so a part
 * both roots relate but only one references — or neither — is absent and
 * stays valid, as those rows opened clean. One finding per (canonical
 * part name, segment) pair the model lists both for a main-document
 * entry (no `via`) and for a glossary entry (`via: "glossary"`, #593),
 * so the order of the relationships cannot decide the verdict; the
 * finding names the part through the main document's spelling, which the
 * traversal reaches first. Unprobed, and unjudged: a shared header or
 * footer part, a part two glossaries share, and a part one root reads
 * under one of these segments and the other under another. A package
 * whose view cannot be read yields nothing (#816); its own diagnostics
 * block the run.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordReadTypedTargets } from "../../../lib/word-read-parts";

/**
 * The lazily read story segments whose shared part Word refuses, each
 * with the session that showed it: comments (session 66, s66a), footnotes
 * and endnotes (session 67, s67a). Headers and footers are unprobed and
 * stay unjudged.
 */
const SHARED_STORY_SEGMENTS: ReadonlyMap<string, string> = new Map([
    ["comments", "session 66"],
    ["footnotes", "session 67"],
    ["endnotes", "session 67"],
]);

export const wordSharedStoryPartCheck: CheckDefinition = {
    id: "word-shared-story-part",
    description: "A comments, footnotes or endnotes part is read by one story root only.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        // The read model lists a lazily read part only when its root uses
        // it (#437), so the (canonical part, segment) pairs each root's
        // entries name are the whole rule: a pair a main-document entry
        // (no `via`) and a glossary entry (`via: "glossary"`, #593) both
        // name is a part both story roots read. Main-document entries come
        // first in traversal order, so the finding reports the part through
        // the spelling the main document's own relationship resolved to.
        const main = new Map<string, { readonly part: string; readonly segment: string; readonly session: string }>();
        const glossary = new Map<string, string>();
        for (const target of await wordReadTypedTargets(ctx, profile)) {
            const session = SHARED_STORY_SEGMENTS.get(target.segment);
            if (session === undefined) {
                continue;
            }
            const key = `${canonicalPartName(target.part)}\n${target.segment}`;
            if (target.via === "glossary") {
                if (target.root !== undefined && !glossary.has(key)) {
                    glossary.set(key, target.root);
                }
            } else if (!main.has(key)) {
                main.set(key, { part: target.part, segment: target.segment, session });
            }
        }
        for (const [key, entry] of main) {
            const root = glossary.get(key);
            if (root === undefined) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    `${entry.part} is the ${entry.segment} part of both the main document and the reached glossary ${root}. ` +
                    `Microsoft Word offers to recover the file (${entry.session}). ` +
                    `Give each story root its own ${entry.segment} part.`,
                path: entry.part,
                code: "word-shared-story-part",
            });
        }
        return issues;
    },
};
