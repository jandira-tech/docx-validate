/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A related part declared with a different mapped part's media type (#600).
 * Session 24 observed Word 16.115 recover main-document header/footer swaps
 * and a reached glossary footer swap, including a parameterized swap.
 * Session 26 observed the reached glossary header swap. Sessions 49 and 50
 * (2026-10-26) observed recovery for the main styles part typed with the
 * settings type and for the main document and a related settings or
 * numbering part typed with another mapped Word type (the styles type on
 * the main document, the settings type on a related part). Session 53
 * (s53a–s53h) observed the remainder and the co-reach: every other mapped
 * Word type on those reaches, and an eagerly read part — the main document,
 * a related numbering part or the styles part — that a read footer or a
 * second eager relationship of another kind also reaches, whatever the
 * Override says. Word-valid reports those observed shapes — the swaps Word
 * reads through the swapping relationship, the eagerly parsed main document
 * and styles, settings and numbering reaches, which need no usage lookup,
 * and the co-reach of a read footer typed with the main document, numbering
 * or styles type; fontTable, webSettings, theme, the stories and every
 * glossary reach were not shown a mapped type and stay silent until probed.
 * Strict judges every mapped relationship.
 * Clause 11.3 of ECMA-376 Part 1 (4th edition) assigns the WML types; theme
 * is DrawingML and its assignment is in clause 14.2.7 instead.
 */
import type { CheckDefinition } from "../../../lib/check-registry";
import {
    asciiLowercase,
    canonicalPartName,
    contentTypeMediaType,
    findContentTypesStream,
    parseContentTypes,
    resolveContentType,
} from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordReadTypedTargets } from "../../../lib/word-read-parts";
import { MAIN_DOCUMENT_CONTENT_TYPES, REPORTING_SEGMENT_ORDER } from "./word-part-content-type-generic";
import {
    expectedContentTypesFor,
    LAZY_OBSERVED_SEGMENTS,
    SPEC_SEGMENTS,
    wordContentTypeTargets,
    type WordPartContentTypeFinding,
} from "./word-part-content-type-inexact";

export const WORD_PART_CONTENT_TYPE_MISMATCH = "word-part-content-type-mismatch";
// Evidence: word-probes-2026-10-24.md, sessions 24, 26, 49 and 50. The
// header/footer keys are the swaps Word was shown, judged read; sessions 49
// and 50 add the eagerly parsed reaches — the main document and the main
// document's own styles, settings and numbering relationships — typed with
// another mapped Word type.
const OBSERVED_SESSIONS: ReadonlyMap<string, number> = new Map([
    ["main\nfooter", 24],
    ["main\nheader", 24],
    ["glossary\nfooter", 24],
    ["glossary\nheader", 26],
    ["main\nofficeDocument", 50],
    ["main\nstyles", 49],
    ["main\nsettings", 50],
    ["main\nnumbering", 50],
]);
/**
 * Session 53's co-reach (s53a): a footer relationship the body references
 * whose target is typed with the main document, numbering or styles type —
 * Word refused every co-reach shape whatever the Override says, so a
 * matching reach on the same part exempts none of it. An exact spelling
 * only; an inexact one is the inexact check's finding.
 */
const FOOTER_COREACH_TYPES: ReadonlySet<string> = new Set(
    [MAIN_DOCUMENT_CONTENT_TYPES, expectedContentTypesFor("numbering"), expectedContentTypesFor("styles")].flat().map(asciiLowercase),
);
const FOOTER_COREACH_SESSION = 53;
export const MAPPED_MEDIA_TYPES: ReadonlySet<string> = new Set([...SPEC_SEGMENTS].flatMap(expectedContentTypesFor).map(asciiLowercase));

export const wordPartContentTypeMismatchFindings = async (
    ctx: ValidationContext,
    scope: "observed" | "spec" = "observed",
    profile?: Profile,
): Promise<readonly WordPartContentTypeFinding[]> => {
    const stream = findContentTypesStream(ctx.partNames);
    if (stream === undefined) return [];
    const parsed = ctx.effectiveXml(stream);
    if (!parsed?.ok) return [];
    const table = parseContentTypes(parsed.document);
    // Unread parts keep every reaching type for the spec scope; parts
    // read through story references are judged only through those segments
    // (#607). Eager relationships keep their existing spec scope.
    const targets = await wordContentTypeTargets(ctx, "spec", profile);
    const byPart = new Map<string, typeof targets>();
    for (const target of targets) {
        const canonical = canonicalPartName(target.part);
        const entries = byPart.get(canonical) ?? [];
        entries.push(target);
        byPart.set(canonical, entries);
    }
    let readKeys: ReadonlySet<string> | undefined;
    const findings: WordPartContentTypeFinding[] = [];
    for (const [canonical, candidates] of byPart) {
        const entries = [...candidates].sort(
            (a, b) => REPORTING_SEGMENT_ORDER.indexOf(a.segment) - REPORTING_SEGMENT_ORDER.indexOf(b.segment),
        );
        const { part } = entries[0]!;
        const resolved = resolveContentType(table, part);
        if (resolved === null) continue;
        const declared = resolved.contentType;
        const mediaType = contentTypeMediaType(declared);
        // Generic XML and unknown types belong outside this check.
        if (!MAPPED_MEDIA_TYPES.has(mediaType)) continue;
        const matches = (segment: string): boolean => expectedContentTypesFor(segment).some((type) => asciiLowercase(type) === mediaType);
        const swaps = entries.filter(
            ({ segment }) =>
                (segment === "header" || segment === "footer") &&
                mediaType === asciiLowercase(expectedContentTypesFor(segment === "header" ? "footer" : "header")[0]!),
        );
        // Read story XML for a possible swap, the session 53 footer
        // co-reach, or competing relationship types. Unrelated parts need no
        // usage lookup; an eager read also interferes with staged-directory
        // write-back failures after repair (#412).
        if (swaps.length > 0 || entries.length > 1 || entries.some(({ segment }) => segment === "footer")) {
            readKeys ??= new Set(
                (await wordReadTypedTargets(ctx, profile)).map(
                    (target) => `${target.via ?? "main"}\n${canonicalPartName(target.part)}\n${target.segment}`,
                ),
            );
        }
        const readEntries = entries.filter(
            ({ segment, via }) => LAZY_OBSERVED_SEGMENTS.has(segment) && readKeys?.has(`${via ?? "main"}\n${canonical}\n${segment}`),
        );
        const judged = readEntries.length > 0 ? readEntries : entries;
        // A matching used type exempts the part's story reaches, including
        // inexact spellings owned by the inexact check; an unused match
        // cannot hide a read swap (#607). Session 53 ended the exemption's
        // reach over the whole part: a type that matches one reach no longer
        // exempts a mismatch on another observed reach, so every observed
        // reach is judged over all of the part's entries, not over `judged`
        // (#739).
        const storyExempt = readEntries.some(({ segment }) => matches(segment));
        // Session 53's footer co-reach: a read footer typed with the main
        // document, numbering or styles type, which no matching reach exempts.
        const isFooterCoreach = (entry: (typeof entries)[number]): boolean =>
            entry.segment === "footer" &&
            entry.via !== "glossary" &&
            FOOTER_COREACH_TYPES.has(mediaType) &&
            asciiLowercase(declared) === mediaType &&
            readKeys?.has(`${entry.via ?? "main"}\n${canonical}\nfooter`) === true;
        // The observed set covers the read header/footer swaps, the eagerly
        // parsed reaches sessions 49 and 50 probed — the main document and
        // the main document's own styles, settings and numbering
        // relationships, which need no usage lookup — and session 53's
        // footer co-reach. Header and footer keep the swap shape `swaps`
        // pins, read through the swapping segment.
        const observed =
            (storyExempt
                ? undefined
                : swaps.find(
                      ({ segment, via }) =>
                          readKeys?.has(`${via ?? "main"}\n${canonical}\n${segment}`) &&
                          OBSERVED_SESSIONS.has(`${via ?? "main"}\n${segment}`),
                  )) ??
            entries.find(
                ({ segment, via }) =>
                    segment !== "header" &&
                    segment !== "footer" &&
                    !matches(segment) &&
                    OBSERVED_SESSIONS.has(`${via ?? "main"}\n${segment}`),
            ) ??
            entries.find(isFooterCoreach);
        const chosen = observed ?? (scope === "spec" && !judged.some(({ segment }) => matches(segment)) ? judged[0] : undefined);
        if (chosen === undefined) continue;
        const { segment, via } = chosen;
        const by =
            segment === "officeDocument"
                ? "the package officeDocument relationship"
                : `the ${via === "glossary" ? "glossary" : "main document"}'s ${segment} relationship`;
        const expected = expectedContentTypesFor(segment)
            .map((type) => `"${type}"`)
            .join(" or ");
        const partTitle = segment === "officeDocument" ? "Main Document" : segment;
        const clauseNumber = segment === "officeDocument" ? "11.3.10" : segment === "theme" ? "14.2.7" : "11.3";
        const clause =
            scope === "spec"
                ? `, the content type ECMA-376 Part 1 (4th edition), clause ${clauseNumber} assigns the ${partTitle} part`
                : "";
        const session =
            observed === undefined
                ? undefined
                : isFooterCoreach(observed)
                  ? FOOTER_COREACH_SESSION
                  : OBSERVED_SESSIONS.get(`${observed.via ?? "main"}\n${observed.segment}`);
        const recovery = session === undefined ? "" : ` Word 16.115 offers to recover the package (session ${session}).`;
        findings.push({
            segment,
            segments: [...new Set(entries.map((entry) => entry.segment))],
            source: resolved.source,
            issue: {
                code: WORD_PART_CONTENT_TYPE_MISMATCH,
                severity: "error",
                path: part,
                message: `${part}, named by ${by}, is typed "${declared}" by ${resolved.source === "override" ? "an Override" : "a Default"}; it expects ${expected}${clause}.${recovery}`,
            },
        });
    }
    return findings;
};

export const wordPartContentTypeMismatchCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_MISMATCH,
    description: "Related Word parts carry their relationship's content type, not another part's type.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeMismatchFindings(ctx, "observed", profile)).map((finding) => finding.issue),
};
