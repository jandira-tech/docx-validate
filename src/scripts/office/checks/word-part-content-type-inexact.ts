/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-inexact` check (#550, renamed from
 * `word-part-content-type-parameter` for #540): a part Word reads, typed by
 * a spelling of its own Word content type that is not exactly that type
 * ignoring ASCII case.
 *
 * Word 16.115 offers to recover the package for every probed shape: any
 * MIME parameter, with or without a space and whether `charset` or not —
 * even though RFC 7303 §9.1 registers `charset` as an optional parameter of
 * `application/xml` — and a trailing space. An upper-case spelling opens
 * clean, so Word folds ASCII case and nothing else (probe sessions 16 and
 * 17, 2026-10-24; the table lives with issue #550).
 *
 * Scope follows the observation. Word-valid judges the parts sessions 17
 * and 20 probed: the main document — named with all four Word
 * main-document content types, since the check judges the package, not the
 * file name (#505) — styles, settings, fontTable, webSettings, a header the
 * main document references through `w:headerReference`, a related numbering
 * or theme part (both eagerly parsed, so related is enough), and a footer,
 * footnotes, endnotes or comments part the document uses, gated as the
 * header is through `wordReadParts` — a related-but-unused lazy part opens
 * clean, as #516 observed for the header and session 21 for a footer,
 * footnotes, endnotes, comments or header part carrying a parameter
 * (#570). Session 21 also observed a trailing space refused on the six
 * parts session 20 added, as session 17 did on styles. A parameterized
 * glossaryDocument type opens clean (session 20), so it stays silent under
 * word-valid. Sessions 20–22 are in
 * docs/plans/evidence/word-probes-2026-10-24.md.
 *
 * A reached glossary document is a second story root (#549), and word-valid
 * judges the parts its own relationships part names too — but only the
 * segments the probes observed: its styles part, eagerly parsed like the
 * main document's (#562, session 19), so a relationship is enough (session
 * 21, #579), and a header its `w:headerReference` uses, gated by
 * `wordReadParts` exactly as the main document's header is (session 22).
 * The glossary's settings, fontTable, webSettings, numbering and theme
 * relationships were not probed for an inexact type and stay silent under
 * word-valid; strict judges every Part 1 WordprocessingML relationship —
 * and theme — on the glossary too, citing clause 11.3 for the unobserved
 * WordprocessingML segments (the theme, a DrawingML part, cites no clause),
 * which keeps strict a superset (#579). The finding names the
 * relationship that reached the part — the glossary's — and a part the
 * main document and the glossary both relate by the same type is reported
 * once, naming the main document's relationship unless only the glossary
 * reads the part through it. A part reached by two
 * relationship types is judged against each (#593): the segment whose
 * expected content type's media type matches the declared value reports
 * the finding — the types are pairwise distinct, so at most one can match.
 * A lazily parsed segment the story root does not read neither reports nor
 * hides a read segment of the same part, because the `wordReadParts` gate
 * decides per (part, segment) pair, not per part — and a read story
 * segment hides no eager reach either: session 56 (s56a) observed Word
 * 16.115 recover the styles part typed `…styles+xml; charset=utf-8` with
 * a read footer relationship also targeting it (#760), so the spelling is
 * judged over every reach the story roots read, as session 53 judged the
 * mismatch over all of the part's entries (#739).
 *
 * Strict justifies the same defect by the spec — ECMA-376 Part 1 (4th
 * edition), clause 11.3, gives every WordprocessingML part its own content
 * type — so a full strict run also judges the Part 1 part types a
 * main-document relationship names but no probe has observed, citing the
 * content type Part 1 assigns them (#506), like the generic check's spec
 * mode; the header joined the probed set in session 17 and numbering,
 * footer, endnotes, footnotes and comments in session 20 (#570), so only
 * glossaryDocument keeps the clause-11.3 wording among the eager segments.
 * Theme is a DrawingML part — its root is the
 * `a:theme` element dml-main.xsd declares — so it is no clause-11.3 part
 * either; session 20 observed Word reject its parameterized type, which
 * puts it in the probed set strict shares, keeping strict a superset of
 * word-valid (#570). A lazy segment strict judges on a part Word does not
 * read keeps the clause-11.3 wording too: session 21 (s21b) observed a
 * related-but-unused footer, footnotes, endnotes, comments or header part
 * typed `…+xml; charset=utf-8` open clean, so the recovery claim stays on
 * the parts Word reads — eager and used lazy alike (#589).
 *
 * A part whose media type is itself generic XML is the generic check's
 * finding: `word-part-content-type-generic` judges the media type, with
 * parameters stripped, so the two checks can never report one part twice —
 * this one fires only when the media type is exactly the part's expected
 * Word type while the declared value is not (#540).
 *
 * Resolution follows ECMA-376 Part 2 §10.1.2.4 through the shared
 * `src/lib/content-types.ts` table and judges the stream after Markup
 * Compatibility processing, as every content-type judgment does (#459,
 * #489). The finding is an error, and Word-blocking, under word-valid for
 * DOCX/DOCM, an error under strict, and a warning under lenient. A full
 * strict run drops a Default-typed main finding only when
 * `ct-undeclared-part` errors at the same path (#535).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    asciiLowercase,
    canonicalPartName,
    contentTypeMediaType,
    findContentTypesStream,
    parseContentTypes,
    partExtension,
    resolveContentType,
} from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { glossaryTypedRelationshipTargets, wordReadTypedTargets, wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import {
    EXPECTED_CONTENT_TYPES,
    MAIN_DOCUMENT_CONTENT_TYPES,
    OFFICE_DOCUMENT_SEGMENT,
    SPEC_ONLY_CONTENT_TYPES,
} from "./word-part-content-type-generic";

/**
 * The theme part's content type — the DrawingML part whose root element
 * `a:theme` the bundled dml-main.xsd declares. It sits outside clause
 * 11.3's WordprocessingML list, so it has no `SPEC_ONLY_CONTENT_TYPES`
 * entry; session 20 probed it, so it joins the observed set directly (#570).
 */
const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";

export const WORD_PART_CONTENT_TYPE_INEXACT = "word-part-content-type-inexact";

/**
 * The relationship segments the profile-independent probes observed Word
 * refuse an inexact content type on: the main document, styles, settings,
 * fontTable, webSettings, a referenced header (#550, session 17), a related
 * numbering or theme part, and a footer, footnotes, endnotes or comments
 * part the document uses (#570, session 20). `wordReadParts` decides
 * whether the document uses a lazily parsed part, exactly as it does for
 * the XML diagnostics (#516); numbering and theme are eagerly parsed, so
 * related is enough.
 */
const OBSERVED_SEGMENTS: ReadonlySet<string> = new Set([
    OFFICE_DOCUMENT_SEGMENT,
    ...EXPECTED_CONTENT_TYPES.keys(),
    "header",
    "numbering",
    "theme",
    "footer",
    "footnotes",
    "endnotes",
    "comments",
]);

/** The lazily parsed observed segments: judged under the observed scope only when the document uses them (#516, #570). */
export const LAZY_OBSERVED_SEGMENTS: ReadonlySet<string> = new Set(["header", "footer", "comments", "footnotes", "endnotes"]);

/** The segments a strict run judges: the probed set plus every other Part 1 WordprocessingML part type (#506). */
export const SPEC_SEGMENTS: ReadonlySet<string> = new Set([...OBSERVED_SEGMENTS, ...SPEC_ONLY_CONTENT_TYPES.keys()]);

/**
 * The segments a reached glossary's relationships were probed on (#579):
 * its styles part — eagerly parsed (#562, session 19), so a relationship
 * is enough (session 21) — a header it references (session 22), and footer,
 * footnotes, endnotes and comments it references (session 23, #596), gated
 * by `wordReadParts` as lazy segments are. The other glossary-eager
 * segments were not probed for an inexact type, so they stay silent under
 * word-valid.
 */
const GLOSSARY_OBSERVED_SEGMENTS: ReadonlySet<string> = new Set(["styles", "header", "footer", "footnotes", "endnotes", "comments"]);

/** The glossary segments a strict run judges by the spec: every Part 1 WordprocessingML part type plus theme (#579). */
const GLOSSARY_SPEC_SEGMENTS: ReadonlySet<string> = new Set([...SPEC_SEGMENTS].filter((segment) => segment !== OFFICE_DOCUMENT_SEGMENT));

/** The content type a segment's part needs, `undefined` for the multi-typed main document. */
const expectedContentTypeFor = (segment: string): string | undefined =>
    segment === "theme" ? THEME_CONTENT_TYPE : (EXPECTED_CONTENT_TYPES.get(segment) ?? SPEC_ONLY_CONTENT_TYPES.get(segment));

/** What the part needs, named for the finding's message; `observed` names whether the segment+source pair was probed (#579). */
const expectedText = (segment: string, source: "override" | "default", observed: boolean): string => {
    if (segment === OFFICE_DOCUMENT_SEGMENT) {
        const list = MAIN_DOCUMENT_CONTENT_TYPES.map((type) => `"${type}"`);
        const types = `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
        return source === "default"
            ? `it needs its own Override carrying exactly one of the Word main-document content types — ${types}`
            : `its content type must be exactly one of the Word main-document content types — ${types}`;
    }
    const expected = expectedContentTypeFor(segment)!;
    if (observed) {
        return source === "default" ? `it needs its own Override "${expected}" exactly` : `its Override should carry "${expected}" exactly`;
    }
    // A spec-only segment has no Word observation, so its finding cites
    // clause 11.3, as the generic check's do (#506). A reached glossary's
    // theme is spec-only too (#579), but the theme is a DrawingML part, not
    // a WordprocessingML one, so it names its own content type and no clause.
    const clause =
        segment === "theme"
            ? ", the theme part's own content type"
            : `, the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the ${segment} part`;
    return source === "default" ? `it needs its own Override "${expected}"${clause}` : `its Override should carry "${expected}"${clause}`;
};

/** Shared related-part finding, retaining resolution details for strict de-duplication (#535, #600). */
export type WordPartContentTypeFinding = {
    readonly issue: ValidationIssue;
    /** Which relationship segment reached the part — the one the message names. */
    readonly segment: string;
    /** Every in-scope relationship segment that reaches the part (#593). */
    readonly segments: readonly string[];
    /** Which declaration produced the content type. */
    readonly source: "override" | "default";
};

/** Existing inexact finding name, retained for callers. */
export type WordInexactContentTypeFinding = WordPartContentTypeFinding;

/** Shared typed-target walk: every reaching (part, segment, story root) stays available (#593, #600). */
export const wordContentTypeTargets = async (ctx: ValidationContext, scope: "observed" | "spec", profile?: Profile) => {
    const targets = await wordTypedRelationshipTargets(ctx, scope === "spec" ? SPEC_SEGMENTS : OBSERVED_SEGMENTS, profile);
    const glossaryTargets = await glossaryTypedRelationshipTargets(
        ctx,
        scope === "spec" ? GLOSSARY_SPEC_SEGMENTS : GLOSSARY_OBSERVED_SEGMENTS,
        profile,
    );
    return [...targets, ...glossaryTargets];
};

/** The shared expected-type map, including all four main-document types (#505, #600). */
export const expectedContentTypesFor = (segment: string): readonly string[] => {
    if (segment === OFFICE_DOCUMENT_SEGMENT) return MAIN_DOCUMENT_CONTENT_TYPES;
    const expected = expectedContentTypeFor(segment);
    return expected === undefined ? [] : [expected];
};

/**
 * The parts {@link wordPartContentTypeInexactCheck} reports, one finding
 * per part: the media type — parameters stripped, trimmed — is the part's
 * expected Word content type while the declared value is not exactly that
 * type ignoring ASCII case (#550). `scope` is `"observed"` for the profiles
 * that predict Word — the probed parts only — and `"spec"` for strict,
 * which also judges the other Part 1 WordprocessingML part types (#506).
 */
export const wordPartContentTypeInexactFindings = async (
    ctx: ValidationContext,
    scope: "observed" | "spec" = "observed",
    profile?: Profile,
): Promise<readonly WordInexactContentTypeFinding[]> => {
    // A missing stream is `ct-missing`; one that does not parse is the
    // XML syntax gate's finding. Neither is judged here.
    const stream = findContentTypesStream(ctx.partNames);
    if (stream === undefined) {
        return [];
    }
    // Resolution reads the stream after Markup Compatibility processing
    // (OPC [M1.19]), as every content-type judgment does (#459, #489).
    const parsed = ctx.effectiveXml(stream);
    if (!parsed?.ok) {
        return [];
    }
    const table = parseContentTypes(parsed.document);

    const merged = await wordContentTypeTargets(ctx, scope, profile);
    // Headers, footers, comments, footnotes and endnotes are parsed lazily
    // — only when the document uses them (#516); session 20's specimens
    // reference theirs (#570) — so the typed reads `wordReadTypedTargets`
    // computes gate the lazy candidates the reachability walk named, the
    // same usage decision the XML diagnostics make, and the glossary's
    // targets too: its styles part is eagerly parsed (always read) while
    // its referenced stories are lazy like the main document's (#579,
    // #596). The set is
    // computed under the spec scope as well: strict still judges every
    // relationship, but its verdict follows the observation, and session 21
    // (s21b) observed a related-but-unused lazy part open clean — so under
    // strict an unread lazy candidate is reported with the spec text and no
    // Word sentence instead of being skipped (#589). The gate keys on
    // (part, segment): a part read through one type but related through an
    // unused lazy one is read through the first alone — the lazy segment
    // neither hides the read one nor adds a finding (#593).
    const readKeys = merged.some((target) => target.via === "glossary" || LAZY_OBSERVED_SEGMENTS.has(target.segment))
        ? new Set(
              (await wordReadTypedTargets(ctx, profile)).map(
                  (target) => `${target.via ?? "main"}\n${canonicalPartName(target.part)}\n${target.segment}`,
              ),
          )
        : undefined;
    // One finding per part. The types are pairwise distinct, so the
    // declared media type matches at most one segment's expected type; that
    // segment's entries — the main document's and a glossary's, when both
    // roots relate the part by it — are the candidates to report it.
    const byPart = new Map<string, typeof merged>();
    for (const target of merged) {
        const canonical = canonicalPartName(target.part);
        const entries = byPart.get(canonical) ?? [];
        entries.push(target);
        byPart.set(canonical, entries);
    }
    const findings: WordInexactContentTypeFinding[] = [];
    for (const [canonical, entries] of byPart) {
        const { part } = entries[0]!;
        const resolved = resolveContentType(table, part);
        if (resolved === null) {
            continue;
        }
        const declared = resolved.contentType;
        // A media type outside every reaching segment's own types is the
        // generic check's class or no known defect at all; this check never
        // double-reports.
        const mediaType = contentTypeMediaType(declared);
        // A lazily parsed candidate its story root does not read through
        // that segment is not Word's finding, and cannot hide a read one
        // (#593); `readKeys === undefined` (usage undecidable, #536) counts
        // as read.
        const isRead = ({ segment, via }: (typeof entries)[number]): boolean =>
            !(
                (via === "glossary" || LAZY_OBSERVED_SEGMENTS.has(segment)) &&
                readKeys !== undefined &&
                !readKeys.has(`${via ?? "main"}\n${canonical}\n${segment}`)
            );
        // A part read through another story-reference segment is a mismatch,
        // even when an unused relationship matches this inexact media type
        // (#607), so the judgment covers the reaches the story roots read —
        // a read lazy segment and every eager one — and an unused lazy
        // relationship neither reports nor hides. A read story segment
        // hides no eager reach either: session 56 (s56a) observed Word
        // recover the styles part typed `…styles+xml; charset=utf-8` with a
        // read footer relationship also targeting it (#760), as session 53
        // ended the matching reach's exemption in the mismatch check
        // (#739). Eager relationships keep their existing spec scope.
        const readEntries = entries.filter((entry) => LAZY_OBSERVED_SEGMENTS.has(entry.segment) && isRead(entry));
        const judged =
            readEntries.length > 0 ? entries.filter((entry) => !LAZY_OBSERVED_SEGMENTS.has(entry.segment) || isRead(entry)) : entries;
        const matching = judged.filter(({ segment }) =>
            expectedContentTypesFor(segment).some((type) => mediaType === asciiLowercase(type)),
        );
        if (matching.length === 0) {
            continue;
        }
        // Word folds ASCII case and nothing else: the declared value must be
        // the expected type to the letter, so any parameter or padding
        // reports while an upper-case spelling stays silent (session 17).
        // Both sides lowercase — `webSettings` and `fontTable` are camelCase.
        if (
            matching.some(({ segment }) =>
                expectedContentTypesFor(segment).some((type) => asciiLowercase(declared) === asciiLowercase(type)),
            )
        ) {
            continue;
        }
        // The observed pairs are the main-document probes plus the
        // glossary's probed styles and referenced stories (#579, #596); every
        // other glossary segment is spec-only.
        const isProbed = ({ segment, via }: (typeof matching)[number]): boolean =>
            via === "glossary" ? GLOSSARY_OBSERVED_SEGMENTS.has(segment) : OBSERVED_SEGMENTS.has(segment);
        // The finding names a read, probed candidate when there is one, then
        // a read one. Under the observed scope an unread part is not Word's
        // finding at all; strict still reports it, by the spec (#589).
        const chosen =
            matching.find((entry) => isRead(entry) && isProbed(entry)) ??
            matching.find(isRead) ??
            (scope === "spec" ? matching[0] : undefined);
        if (chosen === undefined) {
            continue;
        }
        const { segment, via } = chosen;
        // The message names the relationship that reached the part — the
        // glossary's for a part its own relationships part names (#579).
        const by =
            via === "glossary"
                ? `the glossary's ${segment} relationship`
                : segment === OFFICE_DOCUMENT_SEGMENT
                  ? "the package officeDocument relationship"
                  : `the main document's ${segment} relationship`;
        // Word recovery was observed only for the probed parts Word reads. A
        // lazy segment its story root does not read through opened clean
        // (session 21, s21b), so its finding cites the spec alone.
        const observed = isProbed(chosen) && isRead(chosen);
        const resolvedText =
            resolved.source === "default"
                ? `resolves only through the Default for extension "${partExtension(part)}" to "${declared}"`
                : `is typed "${declared}" by an Override`;
        // Word recovery was observed only for the probed parts (#506).
        const verdict = observed ? " Word 16.115 offers to recover the package." : "";
        findings.push({
            segment,
            segments: [...new Set(entries.map((entry) => entry.segment))],
            source: resolved.source,
            issue: {
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_INEXACT,
                path: part,
                message: `${part}, named by ${by}, ${resolvedText}; ${expectedText(segment, resolved.source, observed)}.${verdict}`,
            },
        });
    }
    return findings;
};

export const wordPartContentTypeInexactCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_INEXACT,
    description: "Parts Word reads carry exactly their Word content type, ignoring ASCII case — no MIME parameter or padding.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeInexactFindings(ctx, "observed", profile)).map((finding) => finding.issue),
};
