/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-unrecognized` check (#707): a part Word reads,
 * typed by a media type that is neither its own Word content type, nor a
 * generic XML type, nor another mapped Word part's type — an unrecognised
 * type such as `application/octet-stream`.
 *
 * Session 46 (chunk s46e of docs/plans/evidence/word-probes-2026-10-24.md)
 * observed Word 16.115 offer recovery for the main document's styles
 * relationship naming `word/styles.bin` typed `application/octet-stream` by a
 * Default — whether the part held non-XML bytes or well-formed styles XML —
 * while the same XML typed as styles opened clean. Word refuses a part it
 * reads whose content type is neither its own type nor generic XML. Only the
 * main document's `styles` reach is observed; the Override spelling and
 * every other segment — a reached glossary's styles included — are inference
 * from the same resolution, so their findings cite the spec instead of the
 * recovery claim.
 *
 * The media type — parameters stripped, trimmed, ASCII-lowercased — decides:
 * a generic XML type is `word-part-content-type-generic`'s finding and
 * another mapped Word part's type is `word-part-content-type-mismatch`'s, so
 * the checks stay disjoint (#540, #600). An empty or undeclared type stays
 * with `word-content-type-invalid` and `ct-undeclared-part`.
 *
 * Every mapped reach is judged — the same targets the generic, inexact and
 * mismatch checks walk — so an unobserved segment warns under word-valid
 * rather than staying silent, errors under strict (ECMA-376 Part 1, 4th
 * edition, clause 11.3 gives every WordprocessingML part its own content
 * type, clause 14.2.7 the theme's, and clause 11.3.10 the Main Document's),
 * and warns under lenient. Word-valid blocks only the observed reach: the
 * main document's `styles` relationship. That verdict is structural —
 * `wordObservedContentTypeParts` resolves the reach once per run for
 * `isWordBlockingIssue` — because the message embeds the declared content
 * type and the stored part name verbatim, so a package could carry the
 * recovery sentence itself (#718). One finding per part; when several
 * relationship types reach it, the observed one names the finding whenever
 * it exists, and the shared reporting order ranks the rest (#593).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    canonicalPartName,
    contentTypeMediaType,
    findContentTypesStream,
    parseContentTypes,
    resolveContentType,
} from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordTypedRelationshipTargets, type WordTypedTarget } from "../../../lib/word-read-parts";
import { GENERIC_CONTENT_TYPES, REPORTING_SEGMENT_ORDER } from "./word-part-content-type-generic";
import { expectedContentTypesFor, wordContentTypeTargets, type WordPartContentTypeFinding } from "./word-part-content-type-inexact";
import { MAPPED_MEDIA_TYPES } from "./word-part-content-type-mismatch";

export const WORD_PART_CONTENT_TYPE_UNRECOGNIZED = "word-part-content-type-unrecognized";

/**
 * The reach session 46 observed (s46e): the main document's `styles`
 * relationship. One source for the check's own `isObserved` and for
 * `wordObservedContentTypeParts`, which carries the fact to
 * `isWordBlockingIssue` (#718).
 */
export const UNRECOGNIZED_OBSERVED_SEGMENTS: ReadonlySet<string> = new Set(["styles"]);

/** The reach {@link UNRECOGNIZED_OBSERVED_SEGMENTS} names; a glossary's own relationship never reaches it (#579). */
const isObserved = (target: WordTypedTarget): boolean => target.via !== "glossary" && UNRECOGNIZED_OBSERVED_SEGMENTS.has(target.segment);

/**
 * Canonical names of the parts the observed reach names, computed once per
 * run and passed to `isWordBlockingIssue` as `observedContentTypeParts`:
 * the structural fact that decides a `word-part-content-type-unrecognized`
 * finding's word-valid severity (#718), the same shape of per-run fact the
 * properties, Word-read and story-root sets are. The message cannot carry
 * it — it embeds the declared content type and the stored part name
 * verbatim, so a package controls every substring of it.
 */
export const wordObservedContentTypeParts = async (ctx: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> =>
    new Set(
        (await wordTypedRelationshipTargets(ctx, UNRECOGNIZED_OBSERVED_SEGMENTS, profile)).map((target) => canonicalPartName(target.part)),
    );

/**
 * The parts {@link wordPartContentTypeUnrecognizedCheck} reports, one finding
 * per part. Every mapped reach reports under every profile — the profile's
 * severity policy turns the unobserved ones into warnings under word-valid —
 * citing the spec instead of the recovery claim.
 */
export const wordPartContentTypeUnrecognizedFindings = async (
    ctx: ValidationContext,
    profile?: Profile,
): Promise<readonly WordPartContentTypeFinding[]> => {
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

    const targets = await wordContentTypeTargets(ctx, "spec", profile);
    // One finding per part (#593): the observed reach names it whenever it
    // reaches the part, and the shared reporting order ranks the rest, so the
    // relationships part's own order decides neither the verdict nor the
    // message.
    const byPart = new Map<string, typeof targets>();
    for (const target of targets) {
        const canonical = canonicalPartName(target.part);
        const candidates = byPart.get(canonical) ?? [];
        candidates.push(target);
        byPart.set(canonical, candidates);
    }
    const findings: WordPartContentTypeFinding[] = [];
    for (const candidates of byPart.values()) {
        const entries = [...candidates].sort(
            (a, b) => REPORTING_SEGMENT_ORDER.indexOf(a.segment) - REPORTING_SEGMENT_ORDER.indexOf(b.segment),
        );
        const { part } = entries[0]!;
        const resolved = resolveContentType(table, part);
        if (resolved === null) {
            continue;
        }
        const declared = resolved.contentType;
        const mediaType = contentTypeMediaType(declared);
        // Generic XML is the generic check's class, another mapped Word type
        // the mismatch check's, and an empty declared type
        // `word-content-type-invalid`'s class; only the unrecognised remainder
        // lands here.
        if (mediaType === "" || GENERIC_CONTENT_TYPES.has(mediaType) || MAPPED_MEDIA_TYPES.has(mediaType)) {
            continue;
        }
        const observed = entries.find(isObserved);
        const { segment, via } = observed ?? entries[0]!;
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
            observed === undefined
                ? `, the content type ECMA-376 Part 1 (4th edition), clause ${clauseNumber} assigns the ${partTitle} part`
                : "";
        const recovery = observed === undefined ? "" : " Word 16.115 offers to recover the package (session 46).";
        findings.push({
            segment,
            segments: [...new Set(entries.map((entry) => entry.segment))],
            source: resolved.source,
            issue: {
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_UNRECOGNIZED,
                path: part,
                message: `${part}, named by ${by}, is typed "${declared}" by ${resolved.source === "override" ? "an Override" : "a Default"}; it expects ${expected}${clause}.${recovery}`,
            },
        });
    }
    return findings;
};

export const wordPartContentTypeUnrecognizedCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_UNRECOGNIZED,
    description:
        "Parts Word reads carry a recognized content type — their own, generic XML, or another Word part's — never an unrecognised one.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeUnrecognizedFindings(ctx, profile)).map((finding) => finding.issue),
};
