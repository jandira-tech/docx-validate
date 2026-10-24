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
 * Scope follows the observation. Word-valid judges the parts session 17
 * probed: the main document — named with all four Word main-document
 * content types, since the check judges the package, not the file name
 * (#505) — styles, settings, fontTable, webSettings, and a header the main
 * document references through `w:headerReference` (a related-but-unused
 * header opens clean, as #516 observed). Footer, comments, footnotes,
 * endnotes, numbering, theme and glossaryDocument stay unprobed for an
 * inexact type and silent under word-valid.
 *
 * Strict justifies the same defect by the spec — ECMA-376 Part 1 (4th
 * edition), clause 11.3, gives every WordprocessingML part its own content
 * type — so a full strict run also judges the other Part 1 part types a
 * main-document relationship names: numbering, footer, endnotes, footnotes,
 * comments and glossaryDocument, each citing the content type Part 1
 * assigns it (#506), like the generic check's spec mode.
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
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordReadParts, wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import {
    EXPECTED_CONTENT_TYPES,
    MAIN_DOCUMENT_CONTENT_TYPES,
    OFFICE_DOCUMENT_SEGMENT,
    SPEC_ONLY_CONTENT_TYPES,
} from "./word-part-content-type-generic";

export const WORD_PART_CONTENT_TYPE_INEXACT = "word-part-content-type-inexact";

/**
 * The relationship segments the profile-independent probes observed Word
 * refuse an inexact content type on: the main document, styles, settings,
 * fontTable, webSettings, and a header the document references (#550,
 * session 17). `wordReadParts` decides whether the document uses a lazily
 * parsed header, exactly as it does for the XML diagnostics (#516).
 */
const OBSERVED_SEGMENTS: ReadonlySet<string> = new Set([OFFICE_DOCUMENT_SEGMENT, ...EXPECTED_CONTENT_TYPES.keys(), "header"]);

/** The segments a strict run judges: the probed set plus every other Part 1 WordprocessingML part type (#506). */
const SPEC_SEGMENTS: ReadonlySet<string> = new Set([...OBSERVED_SEGMENTS, ...SPEC_ONLY_CONTENT_TYPES.keys()]);

/** The content type a segment's part needs, `undefined` for the multi-typed main document. */
const expectedContentTypeFor = (segment: string): string | undefined =>
    EXPECTED_CONTENT_TYPES.get(segment) ?? SPEC_ONLY_CONTENT_TYPES.get(segment);

/** What the part needs, named for the finding's message. */
const expectedText = (segment: string, source: "override" | "default"): string => {
    if (segment === OFFICE_DOCUMENT_SEGMENT) {
        const list = MAIN_DOCUMENT_CONTENT_TYPES.map((type) => `"${type}"`);
        const types = `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
        return source === "default"
            ? `it needs its own Override carrying exactly one of the Word main-document content types — ${types}`
            : `its content type must be exactly one of the Word main-document content types — ${types}`;
    }
    const expected = expectedContentTypeFor(segment)!;
    if (OBSERVED_SEGMENTS.has(segment)) {
        return source === "default" ? `it needs its own Override "${expected}" exactly` : `its Override should carry "${expected}" exactly`;
    }
    // A spec-only segment has no Word observation, so its finding cites
    // clause 11.3, as the generic check's do (#506).
    const clause = `, the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the ${segment} part`;
    return source === "default" ? `it needs its own Override "${expected}"${clause}` : `its Override should carry "${expected}"${clause}`;
};

/** A finding, with the resolution detail the full strict run's de-duplication needs (#535). */
export type WordInexactContentTypeFinding = {
    readonly issue: ValidationIssue;
    /** Which relationship segment reached the part. */
    readonly segment: string;
    /** Which declaration produced the inexact type. */
    readonly source: "override" | "default";
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

    const targets = await wordTypedRelationshipTargets(ctx, scope === "spec" ? SPEC_SEGMENTS : OBSERVED_SEGMENTS);
    // A header is parsed lazily — only when the document uses it (#516), and
    // session 17's specimen references its header — so under the observed
    // scope `wordReadParts` gates the header targets the reachability walk
    // named, the same usage decision the XML diagnostics make. Strict judges
    // every header relationship by the spec, referenced or not.
    const read = scope === "observed" && targets.some((target) => target.segment === "header") ? await wordReadParts(ctx) : undefined;
    const findings: WordInexactContentTypeFinding[] = [];
    for (const { part, segment } of targets) {
        if (read !== undefined && segment === "header" && !read.has(canonicalPartName(part))) {
            continue;
        }
        const resolved = resolveContentType(table, part);
        if (resolved === null) {
            continue;
        }
        const declared = resolved.contentType;
        const expected = segment === OFFICE_DOCUMENT_SEGMENT ? MAIN_DOCUMENT_CONTENT_TYPES : [expectedContentTypeFor(segment)!];
        // Word folds ASCII case and nothing else: the declared value must be
        // the expected type to the letter, so any parameter or padding
        // reports while an upper-case spelling stays silent (session 17).
        // Both sides lowercase — `webSettings` and `fontTable` are camelCase.
        if (expected.some((type) => asciiLowercase(declared) === asciiLowercase(type))) {
            continue;
        }
        // A media type outside the part's own types is the generic check's
        // class or no known defect at all; this check never double-reports.
        const mediaType = contentTypeMediaType(declared);
        if (!expected.some((type) => mediaType === asciiLowercase(type))) {
            continue;
        }
        const by =
            segment === OFFICE_DOCUMENT_SEGMENT ? "the package officeDocument relationship" : `the main document's ${segment} relationship`;
        const resolvedText =
            resolved.source === "default"
                ? `resolves only through the Default for extension "${partExtension(part)}" to "${declared}"`
                : `is typed "${declared}" by an Override`;
        // Word recovery was observed only for the probed parts (#506).
        const verdict = OBSERVED_SEGMENTS.has(segment) ? " Word 16.115 offers to recover the package." : "";
        findings.push({
            segment,
            source: resolved.source,
            issue: {
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_INEXACT,
                path: part,
                message: `${part}, named by ${by}, ${resolvedText}; ${expectedText(segment, resolved.source)}.${verdict}`,
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
    run: async (ctx: ValidationContext): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeInexactFindings(ctx)).map((finding) => finding.issue),
};
