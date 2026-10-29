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
 * Sessions 46, 49 and 50 (chunks s46e, s49a/s49b and s50c/s50d of
 * docs/plans/evidence/word-probes-2026-10-24.md, every row run twice and
 * read twice) observed Word 16.115 offer recovery whenever a part it reads
 * at open carries an unrecognised type: the main document itself — its
 * Override typed `application/octet-stream` (session 49), `APPLICATION/OCTET-
 * STREAM`, `application/octet-stream; charset=utf-8`, `image/png` or
 * `application/foo+xml` (session 50) — the main document's `styles` part —
 * `word/styles.bin` typed `application/octet-stream` by a Default, holding
 * non-XML bytes or well-formed styles XML alike, while the same XML typed
 * as styles opened clean (session 46), and `word/styles.xml` typed by its
 * Override with the same five spellings (session 49) — and the `settings`,
 * `fontTable` and `numbering` parts the main document's own relationships
 * part relates, typed `application/octet-stream` by their Override (session
 * 50) and typed the four other s49a/s50c spellings — `APPLICATION/OCTET-
 * STREAM`, `application/octet-stream; charset=utf-8`, `image/png` or
 * `application/foo+xml` — (session 53, s53b–s53d). Every other reach —
 * through a glossary, a header or footer, or another segment — is inference
 * from the same resolution, so its finding cites the spec instead of the
 * recovery claim.
 *
 * The media type — parameters stripped, trimmed, ASCII-lowercased — decides:
 * a generic XML type is `word-part-content-type-generic`'s finding and
 * another mapped Word part's type is `word-part-content-type-mismatch`'s, so
 * the checks stay disjoint (#540, #600). An empty or undeclared type stays
 * with `word-content-type-invalid` and `ct-undeclared-part`, and so does an
 * `invalid-*` declared value — that check owns the class, reading the same
 * raw spelling on the declaration itself (#719).
 *
 * Every mapped reach is judged: this check walks the spec set
 * (`wordContentTypeTargets(ctx, "spec")` — every main-document relationship
 * of a Part 1 WordprocessingML segment plus theme, and the same set minus
 * `officeDocument` through a reached glossary's own relationships), while
 * the generic check walks its own observed segments through the main
 * document's relationships plus a glossary's `theme` and `styles`, the
 * inexact check the observed set under word-valid and the spec set under
 * strict, and the mismatch check the spec set. So an unobserved segment
 * warns under word-valid rather than staying silent, errors under strict
 * (ECMA-376 Part 1, 4th edition, clause 11.3 gives every WordprocessingML
 * part its own content type, clause 14.2.7 the theme's; the main document is
 * an observed reach since #721, so its findings always cite a session and
 * never a clause), and warns under lenient. Word-valid blocks only the
 * observed reach: the main document and the parts its own relationships
 * part relates by `styles`, `settings`, `fontTable` and `numbering`. That
 * verdict is structural — `wordObservedContentTypeParts` resolves the reach
 * once per run for `isWordBlockingIssue` — because the message embeds the
 * declared content type and the stored part name verbatim, so a package
 * could carry the recovery sentence itself (#718). One finding per part;
 * when several relationship types reach it, the observed one names the
 * finding whenever it exists, and the shared reporting order ranks the rest
 * (#593).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { citesWordSession } from "../../../lib/conformance-class";
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
import { xmlPartNames } from "./xml";

export const WORD_PART_CONTENT_TYPE_UNRECOGNIZED = "word-part-content-type-unrecognized";

/**
 * The sessions that observed each observed reach refuse an unrecognised
 * type, cited by the one recovery-sentence shape: styles was refused typed
 * `application/octet-stream` by a Default (46, s46e) and by an Override with
 * every s49a spelling (49), the main document by Override (49 s49b, 50
 * s50c), and settings, fontTable and numbering typed `application/octet-stream`
 * by their Override (50, s50d) and typed the four other s49a/s50c spellings
 * (53, s53b–s53d). Exported because {@link UNRECOGNIZED_OBSERVED_SEGMENTS}
 * is derived from its keys (#741).
 */
export const RECOVERY_SESSIONS: ReadonlyMap<string, string> = new Map([
    ["officeDocument", "sessions 49 and 50"],
    ["styles", "sessions 46 and 49"],
    ["settings", "sessions 50 and 53"],
    ["fontTable", "sessions 50 and 53"],
    ["numbering", "sessions 50 and 53"],
]);

/**
 * The reach the sessions observed: the main document (`officeDocument`,
 * sessions 49 and 50) and the parts the main document's own relationships
 * part relates by `styles` (46 and 49), `settings`, `fontTable` and
 * `numbering` (50, and 53 for the four other spellings). Derived from
 * {@link RECOVERY_SESSIONS}'s keys, so a segment without a session citation
 * cannot be added (#741); one source for the check's own `isObserved` and
 * for `wordObservedContentTypeParts`, which carries the fact to
 * `isWordBlockingIssue` (#718).
 */
export const UNRECOGNIZED_OBSERVED_SEGMENTS: ReadonlySet<string> = new Set(RECOVERY_SESSIONS.keys());

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
        // An `invalid-*` declared value is `word-content-type-invalid`'s class
        // too (#719): that check reports the declaration itself — the same
        // raw spelling, case-sensitively — so one defect draws one finding.
        if (declared.startsWith("invalid-")) {
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
        const clauseNumber = segment === "theme" ? "14.2.7" : "11.3";
        const clause =
            observed === undefined
                ? `, the content type ECMA-376 Part 1 (4th edition), clause ${clauseNumber} assigns the ${segment} part`
                : "";
        // `isObserved` guarantees the segment is one of RECOVERY_SESSIONS's
        // keys — the observed set is derived from them (#741) — so the
        // citation is defined.
        // Those sessions stored Transitional packages: a package with a
        // Strict part keeps the finding without them (#1016).
        const recovery =
            observed === undefined ||
            !citesWordSession(
                ctx,
                xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
            )
                ? ""
                : ` Word 16.115 offers to recover the package (${RECOVERY_SESSIONS.get(observed.segment)!}).`;
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
