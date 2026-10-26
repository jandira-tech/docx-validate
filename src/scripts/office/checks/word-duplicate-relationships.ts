/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Two theme relationships in a story root's relationships part make
 * Word offer recovery. Either theme alone opens clean; rewriting an
 * absolute target still warns. Session 52 (F1/F2, #728) pinned the count
 * to exact Types the theme rule always used: a
 * `…/relationships/custom/theme` relationship beside the real theme opens
 * clean, two real theme relationships get the recovery offer, and the
 * one-theme control opens clean.
 * Fixture: tests/fixtures/broken/package.word-duplicate-theme.docx.
 *
 * Two `fontTable` relationships in a story root's relationships part
 * make Word offer to recover the file. A second relationship that
 * targets the same `fontTable.xml` as the first still warns. Rewriting
 * an absolute `/word/...` target to a relative path still warns. One
 * font table relationship opens clean. The extra part and its content
 * type are not required for the warning.
 *
 * Two `webSettings` relationships in a story root's relationships part
 * make Word offer to recover the file. A second relationship that
 * targets the same `webSettings.xml` as the first still warns. One web
 * settings relationship opens clean. The extra part is not required for
 * the warning.
 *
 * Two settings relationships in a story root's relationships part make
 * Word offer recovery (fixture 0689). A second relationship that targets
 * the same settings.xml still warns. One settings relationship
 * opens clean.
 *
 * Two `styles` relationships in the main document's relationships part
 * make Word offer to recover the file, to two parts in either order —
 * listed in Id order or against it — and to the same part (#694; session
 * 45); the real styles relationship beside one or two External styles
 * relationships is refused too (session 49). One styles relationship
 * opens clean. A type counts only
 * as the namespace plus the one segment — Word reads Types
 * exactly (session 27) — so `…/relationships/custom/styles` is
 * no styles type (#519); session 47 opened a suffix-only Type
 * clean beside a real one, for settings, webSettings and
 * fontTable too, so every duplicate rule here counts exact
 * Types. Two `numbering` relationships make Word offer to
 * recover the file the same way; one opens clean.
 * Two `footnotes` relationships open clean, so footnotes keeps
 * no check of its own.
 *
 * Scope (#717; sessions 49, 50 and 54, Word 16.115): two of one kind in
 * a reached glossary's own relationships part — to two parts in its
 * folder — is refused for styles, numbering, settings and fontTable
 * (sessions 49 and 50) and for webSettings and theme (session 54,
 * s54d), while the same pair in `_rels/.rels` opens clean for styles,
 * numbering, settings and theme (sessions 49 and 50) and for fontTable
 * and webSettings (s54d), and in a referenced header's sidecar for all
 * six kinds. Every rule here therefore judges only the relationships
 * part of each story root — the main document and every glossary
 * document a main-document `glossaryDocument` relationship reaches,
 * found by relationship as `wordStoryRootParts` finds them — and a
 * relationships part no story root owns is not judged. Lenient output
 * is unchanged.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { officeRelationshipTypeSegment, sidecarOf, wordStoryRootParts } from "../../../lib/word-read-parts";
import { getRelationshipChildren } from "../../../lib/xml-helpers";

/**
 * The stored spellings of the relationships parts these rules judge: the
 * sidecar of every story root — each main document and each reached
 * glossary document — found by relationship (#717), once each. A package
 * whose package relationships cannot be read yields nothing, as every
 * other reach-based rule yields nothing (#388).
 */
const storyRootRelsParts = async (ctx: ValidationContext, profile?: Profile): Promise<readonly string[]> => {
    const parts: string[] = [];
    const seen = new Set<string>();
    for (const root of await wordStoryRootParts(ctx, profile)) {
        const sidecar = sidecarOf(ctx.partNames, root);
        if (sidecar !== undefined && !seen.has(sidecar)) {
            seen.add(sidecar);
            parts.push(sidecar);
        }
    }
    return parts;
};

/** One finding per judged relationships part holding two or more `segment` relationships of exact Type. */
const duplicateFindings = (
    ctx: ValidationContext,
    relsParts: readonly string[],
    segment: string,
    label: string,
    code: string,
): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const relsFile of relsParts) {
        const parsed = ctx.effectiveXml(relsFile);
        if (!parsed?.ok) {
            continue;
        }
        const rels = getRelationshipChildren(parsed.document);
        const ids: string[] = [];
        for (const rel of rels) {
            const type = rel.getAttribute("Type") ?? "";
            if (officeRelationshipTypeSegment(type) !== segment) {
                continue;
            }
            ids.push(rel.getAttribute("Id") ?? "");
        }
        if (ids.length < 2) {
            continue;
        }
        issues.push({
            severity: "error",
            message:
                `${relsFile} contains ${ids.length} ${label} relationships ` +
                `(${ids.join(", ")}). Microsoft Word offers to recover the file. Keep one ${label} relationship.`,
            path: relsFile,
            code,
        });
    }
    return issues;
};

export const wordDuplicateThemeRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-theme-relationship",
    description: "Relationship parts contain at most one theme relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(ctx, await storyRootRelsParts(ctx, profile), "theme", "theme", "word-duplicate-theme-relationship"),
};

export const wordDuplicateFontTableRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-font-table-relationship",
    description: "Relationship parts contain at most one font table relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(ctx, await storyRootRelsParts(ctx, profile), "fontTable", "font table", "word-duplicate-font-table-relationship"),
};

export const wordDuplicateWebSettingsRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-web-settings-relationship",
    description: "Relationship parts contain at most one web settings relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(
            ctx,
            await storyRootRelsParts(ctx, profile),
            "webSettings",
            "web settings",
            "word-duplicate-web-settings-relationship",
        ),
};

export const wordDuplicateSettingsRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-settings-relationship",
    description: "Relationship parts contain at most one settings relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(ctx, await storyRootRelsParts(ctx, profile), "settings", "settings", "word-duplicate-settings-relationship"),
};

export const wordDuplicateStylesRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-styles-relationship",
    description: "Relationship parts contain at most one styles relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(ctx, await storyRootRelsParts(ctx, profile), "styles", "styles", "word-duplicate-styles-relationship"),
};

export const wordDuplicateNumberingRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-numbering-relationship",
    description: "Relationship parts contain at most one numbering relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        duplicateFindings(ctx, await storyRootRelsParts(ctx, profile), "numbering", "numbering", "word-duplicate-numbering-relationship"),
};
