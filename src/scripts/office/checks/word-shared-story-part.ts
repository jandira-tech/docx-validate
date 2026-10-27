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
 * `wordReadTypedTargets` lists these lazily read segments unless the
 * root is known not to use them (#437): a part both roots relate but
 * only one references — or neither — is absent and stays valid, as those
 * rows opened clean, but a root whose XML fails the byte-level checks or
 * whose effective view cannot be read keeps its targets with usage
 * unknown. The rule needs the use itself, so it also asks
 * `storyRootUses` of each root (#858): an unknown use reports nothing
 * here, and the root's own XML diagnostics block the run. One finding
 * per (canonical part name, segment) pair the model lists both for a
 * main-document entry (no `via`) and for a glossary entry (`via:
 * "glossary"`, #593) whose roots are known to use it, so the order of
 * the relationships cannot decide the verdict; the finding names the
 * part through the main document's spelling, which the traversal
 * reaches first. Unprobed, and unjudged: a shared header or
 * footer part, a part two glossaries share, and a part one root reads
 * under one of these segments and the other under another. A package
 * whose view cannot be read yields nothing (#816); its own diagnostics
 * block the run.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { storyRootUses, wordReadTypedTargets } from "../../../lib/word-read-parts";

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
        // A pair a main-document entry (no `via`) and a glossary entry
        // (`via: "glossary"`, #593) both name is a part both story roots
        // reach; the read model keeps a reach whose use it cannot decide
        // (#437), so each pair also needs a root on each side known to use
        // it (#858). Each root's own `officeDocument` entry comes first, so
        // the current main root is the last one seen. Main-document entries
        // come first in traversal order, so the finding reports the part
        // through the spelling the main document's own relationship
        // resolved to.
        const main = new Map<
            string,
            { readonly part: string; readonly segment: string; readonly session: string; readonly roots: string[] }
        >();
        const glossary = new Map<string, string[]>();
        let mainRoot: string | undefined;
        for (const target of await wordReadTypedTargets(ctx, profile)) {
            if (target.via === undefined && target.segment === "officeDocument") {
                mainRoot = target.part;
                continue;
            }
            const session = SHARED_STORY_SEGMENTS.get(target.segment);
            if (session === undefined) {
                continue;
            }
            const key = `${canonicalPartName(target.part)}\n${target.segment}`;
            if (target.via === "glossary") {
                if (target.root !== undefined) {
                    glossary.set(key, [...(glossary.get(key) ?? []), target.root]);
                }
            } else if (mainRoot !== undefined) {
                const entry = main.get(key) ?? { part: target.part, segment: target.segment, session, roots: [] };
                entry.roots.push(mainRoot);
                main.set(key, entry);
            }
        }
        const uses = async (roots: readonly string[], segment: string): Promise<string | undefined> => {
            for (const root of roots) {
                if (await storyRootUses(ctx, root, segment)) {
                    return root;
                }
            }
            return undefined;
        };
        for (const [key, entry] of main) {
            const roots = glossary.get(key);
            if (roots === undefined || (await uses(entry.roots, entry.segment)) === undefined) {
                continue;
            }
            const root = await uses(roots, entry.segment);
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
