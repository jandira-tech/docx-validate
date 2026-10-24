/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-parameter` check (#540): a part Word reads,
 * typed by its own Word content type with a MIME parameter appended.
 *
 * Word 16.115 offers to recover the package for the main document and for
 * styles when their content type carries `; charset=utf-8` — even though the
 * media type is the part's own Word type, and RFC 7303 §9.1 registers
 * `charset` as an optional parameter of `application/xml`. Word compares the
 * content-type string exactly (session 16, 2026-10-24; the table lives with
 * issue #540).
 *
 * Scope follows the observation: the main document — named with all four
 * Word main-document content types, since the check judges the package, not
 * the file name (#505) — and styles. Settings, fontTable, webSettings and
 * the other Part 1 WordprocessingML part types are unprobed for this defect
 * and stay silent until a probe observes them; the generic check is the one
 * that widened its comparison for them (#540).
 *
 * A part whose media type is itself generic XML is the generic check's
 * finding: `word-part-content-type-generic` judges the media type, with
 * parameters stripped, so the two checks can never report one part twice —
 * this one fires only when the media type is exactly the part's expected
 * Word type and a `;` parameter section follows.
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
    contentTypeMediaType,
    findContentTypesStream,
    hasContentTypeParameter,
    parseContentTypes,
    partExtension,
    resolveContentType,
} from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import { EXPECTED_CONTENT_TYPES, MAIN_DOCUMENT_CONTENT_TYPES, OFFICE_DOCUMENT_SEGMENT } from "./word-part-content-type-generic";

export const WORD_PART_CONTENT_TYPE_PARAMETER = "word-part-content-type-parameter";

/** The relationship segments session 16 observed Word refuse a parameterized content type on. */
const PARAMETER_SEGMENTS: ReadonlySet<string> = new Set([OFFICE_DOCUMENT_SEGMENT, "styles"]);

/** What the part needs, named for the finding's message. */
const expectedText = (segment: string, source: "override" | "default"): string => {
    if (segment === OFFICE_DOCUMENT_SEGMENT) {
        const list = MAIN_DOCUMENT_CONTENT_TYPES.map((type) => `"${type}"`);
        const types = `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
        return source === "default"
            ? `it needs its own Override carrying exactly one of the Word main-document content types — ${types}, with no MIME parameter`
            : `its content type must be exactly one of the Word main-document content types — ${types}, with no MIME parameter`;
    }
    const expected = EXPECTED_CONTENT_TYPES.get(segment)!;
    return source === "default"
        ? `it needs its own Override "${expected}", with no MIME parameter`
        : `its Override should carry "${expected}" exactly, with no MIME parameter`;
};

/** A finding, with the resolution detail the full strict run's de-duplication needs (#535). */
export type WordParameterContentTypeFinding = {
    readonly issue: ValidationIssue;
    /** Which relationship segment reached the part. */
    readonly segment: string;
    /** Which declaration produced the parameterized type. */
    readonly source: "override" | "default";
};

/**
 * The parts {@link wordPartContentTypeParameterCheck} reports, one finding
 * per part: the media type is the part's expected Word content type and the
 * declared value carries a MIME parameter (#540).
 */
export const wordPartContentTypeParameterFindings = async (ctx: ValidationContext): Promise<readonly WordParameterContentTypeFinding[]> => {
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

    const findings: WordParameterContentTypeFinding[] = [];
    const targets = await wordTypedRelationshipTargets(ctx, PARAMETER_SEGMENTS);
    for (const { part, segment } of targets) {
        const resolved = resolveContentType(table, part);
        if (resolved === null || !hasContentTypeParameter(resolved.contentType)) {
            continue;
        }
        const mediaType = contentTypeMediaType(resolved.contentType);
        const expected = segment === OFFICE_DOCUMENT_SEGMENT ? MAIN_DOCUMENT_CONTENT_TYPES : [EXPECTED_CONTENT_TYPES.get(segment)!];
        if (!expected.some((type) => mediaType === asciiLowercase(type))) {
            continue;
        }
        const by =
            segment === OFFICE_DOCUMENT_SEGMENT ? "the package officeDocument relationship" : `the main document's ${segment} relationship`;
        const declared = asciiLowercase(resolved.contentType);
        const resolvedText =
            resolved.source === "default"
                ? `resolves only through the Default for extension "${partExtension(part)}" to "${declared}"`
                : `is typed "${declared}" by an Override`;
        findings.push({
            segment,
            source: resolved.source,
            issue: {
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_PARAMETER,
                path: part,
                message: `${part}, named by ${by}, ${resolvedText}; ${expectedText(segment, resolved.source)}. Word 16.115 offers to recover the package.`,
            },
        });
    }
    return findings;
};

export const wordPartContentTypeParameterCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_PARAMETER,
    description: "Parts Word reads carry their exact Word content type, not one with a MIME parameter.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeParameterFindings(ctx)).map((finding) => finding.issue),
};
