/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-generic` check (#432): a part Word reads, typed
 * by a generic XML content type.
 *
 * Word 16.115 offers to recover the package for every part probed whose
 * content type resolves only through a Default to `application/xml`: the
 * main document and styles with their Override removed, and settings,
 * fontTable and webSettings added with a relationship but no Override. Every
 * control with the part's proper Override opens clean (session 8a in
 * docs/plans/evidence/word-probes-2026-10-24.md). Which Default supplied the
 * type does not change what Word receives, so any extension's Default maps
 * to this check; the finding names the one the part's own extension used
 * (#504).
 *
 * Session 15 widened the same defect class: the main document and styles
 * were probed with `text/xml` — by Default or Override — and with
 * `application/xml` by Override, and Word offers to recover every spelling.
 * Settings, fontTable and webSettings were probed only with the
 * `application/xml` Default; reporting the Override and `text/xml` spellings
 * for them is inference from the same resolution (the message names the
 * declaration that produced the type), so the doc comment says so (#517).
 * Session 20 adds related numbering and theme parts typed by the
 * `application/xml` Default: both raised repair warnings, and both exact
 * Override controls opened clean (Word 16.115.26100424, chunks s20c and
 * s20d of session 20 in docs/plans/evidence/word-probes-2026-10-24.md,
 * #571). Other generic spellings for these two parts are inference from
 * the same resolution; customXml stays outside this check.
 * Session 23 observes the same Default-typed theme defect when only a
 * reached glossary relates the theme (#595), with an exact-Override clean
 * control. Only theme is in the glossary's generic-type observed scope;
 * other generic spellings there remain inference from resolution.
 *
 * Session 16 widened the comparison itself: the check judges the media
 * type — everything before the first `;`, ASCII-lowercased — so a
 * parameterized spelling such as `application/xml; charset=utf-8` counts
 * as generic and the finding names the declared value (#540). Word 16.115
 * offers to recover a styles part whose Override carries exactly that
 * value; the parameterized spellings on settings, fontTable and
 * webSettings are inference from the same resolution, like #517's.
 *
 * The main document's finding names every Word main-document content type —
 * document, macro-enabled, template, macro-enabled template — because the
 * check judges the package, not the file name, and naming only the
 * macro-free type would turn a `.docm` into a macro-free document (#505).
 *
 * Resolution follows ECMA-376 Part 2 §10.1.2.4 through the shared
 * `src/lib/content-types.ts` table: an exact Override wins, so a part with a
 * specific Override — right or wrong — is out of scope unless that Override
 * itself carries a generic XML type. The finding is an error, and
 * Word-blocking, under word-valid for DOCX/DOCM, an error under strict
 * (Part 1 gives each of these parts its own content type), and a warning
 * under lenient. A full strict run does not report a main document resolved
 * through a Default, which `ct-undeclared-part` already rejects; a main
 * document carrying a generic Override is reported once, here. A
 * Default-typed theme is likewise suppressed only when that same path has
 * a `ct-undeclared-part` error (#571, #535).
 *
 * Scope follows the justification. Word-valid predicts Word, so it judges
 * only the probed parts above. Strict justifies the same defect by the spec
 * — ECMA-376 Part 1 (4th edition), clause 11.3, gives every WordprocessingML
 * part its own content type — so a full strict run also reports a part a
 * main-document relationship of the other Part 1 types names:
 * header, footer, endnotes, footnotes, comments and glossaryDocument, each
 * citing the content type Part 1 assigns it (#506). `customXml` stays out at
 * every profile — `application/xml` is the content type the spec gives the
 * custom XML data storage part itself — and types outside this list stay
 * unjudged.
 *
 * A part reached by two relationship types is judged against every type
 * that reaches it (#593): the walk yields one entry per (part, segment)
 * pair, so a Default-typed theme also named by an unreferenced footer
 * relationship keeps its theme segment — the reporting segment is the one
 * `REPORTING_SEGMENT_ORDER` ranks first, and the order of the
 * relationships in the relationships part decides neither the verdict nor
 * the message.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    asciiLowercase,
    canonicalPartName,
    findContentTypesStream,
    parseContentTypes,
    partExtension,
    resolveContentType,
    contentTypeMediaType,
} from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { WordTypedTarget } from "../../../lib/word-read-parts";
import { glossaryTypedRelationshipTargets, wordTypedRelationshipTargets } from "../../../lib/word-read-parts";

export const WORD_PART_CONTENT_TYPE_GENERIC = "word-part-content-type-generic";

/** The generic XML content types this check judges, compared ASCII case-insensitively (#517). */
const GENERIC_CONTENT_TYPES: ReadonlySet<string> = new Set(["application/xml", "text/xml"]);

/** The relationship segments Word was observed to reject on a generic XML type, and the content type each part needs. */
export const EXPECTED_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
    ["styles", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"],
    ["settings", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"],
    ["fontTable", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"],
    ["webSettings", "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"],
]);

/**
 * Additional generic-type observations (#571), kept separate because the
 * inexact-type check shares the older maps and has its own observed scope.
 */
const ADDITIONAL_OBSERVED_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
    ["numbering", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"],
    ["theme", "application/vnd.openxmlformats-officedocument.theme+xml"],
]);

/** The main document, reached through the package `officeDocument` relationship rather than a main-document one. */
export const OFFICE_DOCUMENT_SEGMENT = "officeDocument";

/**
 * The Word main-document content types (#505). The check judges the package,
 * not the file name, so an officeDocument finding cannot name one of them —
 * advising the macro-free type would turn a macro-enabled package into a
 * macro-free document.
 */
export const MAIN_DOCUMENT_CONTENT_TYPES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
] as const;

/**
 * The additional Part 1 types used by the shared inexact-type check, with
 * the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns each
 * (#506). A main-document relationship of one of these types is judged under
 * strict only; generic numbering is now observed separately (#571).
 */
export const SPEC_ONLY_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
    ["numbering", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"],
    ["header", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"],
    ["footer", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"],
    ["endnotes", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"],
    ["footnotes", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"],
    ["comments", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"],
    ["glossaryDocument", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"],
]);

/** The relationship segments this check reaches for under word-valid and lenient: the probed parts plus the main document. */
const OBSERVED_SEGMENTS: ReadonlySet<string> = new Set([
    OFFICE_DOCUMENT_SEGMENT,
    ...EXPECTED_CONTENT_TYPES.keys(),
    ...ADDITIONAL_OBSERVED_CONTENT_TYPES.keys(),
]);

/** The segments a strict run judges: the probed set plus every other Part 1 WordprocessingML part type (#506). */
const SPEC_SEGMENTS: ReadonlySet<string> = new Set([...OBSERVED_SEGMENTS, ...SPEC_ONLY_CONTENT_TYPES.keys()]);

/**
 * Which of a part's reaching segments reports the finding when more than one
 * would (#593): the observed segments first — the probed ones before the
 * additional ones — then the Part 1 spec-only types, so the message names a
 * relationship Word was observed reading whenever one exists, and the
 * relationships part's own order decides neither the verdict nor the
 * message. Segments outside the scope requested cannot appear; an
 * unlisted segment ranks last.
 */
const REPORTING_SEGMENT_ORDER: readonly string[] = [...SPEC_SEGMENTS];

/** What the part needs, named for the finding's message. */
const expectedContentTypeText = (segment: string, source: "override" | "default"): string => {
    if (segment === OFFICE_DOCUMENT_SEGMENT) {
        const list = MAIN_DOCUMENT_CONTENT_TYPES.map((type) => `"${type}"`);
        const types = `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
        return source === "default"
            ? `it needs its own Override carrying one of the Word main-document content types — ${types}`
            : `its Override should carry one of the Word main-document content types — ${types}`;
    }
    const observed = ADDITIONAL_OBSERVED_CONTENT_TYPES.get(segment);
    if (observed !== undefined) {
        return source === "default" ? `it needs its own Override "${observed}"` : `its Override should carry "${observed}"`;
    }
    const specOnly = SPEC_ONLY_CONTENT_TYPES.get(segment);
    if (specOnly !== undefined) {
        return source === "default"
            ? `it needs its own Override "${specOnly}", the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the ${segment} part`
            : `its Override should carry "${specOnly}", the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the ${segment} part`;
    }
    const expected = EXPECTED_CONTENT_TYPES.get(segment)!;
    return source === "default" ? `it needs its own Override "${expected}"` : `its Override should carry "${expected}"`;
};

/** A finding, with the resolution detail the full strict run's de-duplication needs (#517). */
export type WordGenericContentTypeFinding = {
    readonly issue: ValidationIssue;
    /** Which relationship segment reached the part — the one the message names. */
    readonly segment: string;
    /** Every in-scope relationship segment that reaches the part (#593). */
    readonly segments: readonly string[];
    /** Which declaration produced the generic type. */
    readonly source: "override" | "default";
};

/**
 * The parts {@link wordPartContentTypeGenericCheck} reports, one finding per
 * part. `scope` is `"observed"` for the profiles that predict Word — the
 * probed parts only — and `"spec"` for strict, which also judges the other
 * Part 1 WordprocessingML part types (#506). Findings for the spec-only
 * types cite ECMA-376 Part 1 instead of Word recovery, which no probe has
 * observed for them.
 */
export const wordPartContentTypeGenericFindings = async (
    ctx: ValidationContext,
    scope: "observed" | "spec" = "observed",
): Promise<readonly WordGenericContentTypeFinding[]> => {
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

    const findings: WordGenericContentTypeFinding[] = [];
    const targets = await wordTypedRelationshipTargets(ctx, scope === "spec" ? SPEC_SEGMENTS : OBSERVED_SEGMENTS);
    // Sessions 23 and 24 observed recovery for the reached glossary's
    // generic-Default theme and styles (#595, #601); other segments stay unprobed.
    const glossaryTargets = await glossaryTypedRelationshipTargets(ctx, new Set(["theme", "styles"]));
    // One finding per part. A part reached through two relationship types
    // keeps every segment (#593); the message names the segment
    // `REPORTING_SEGMENT_ORDER` ranks first, deterministically, so the
    // relationships part's own order decides nothing.
    const byPart = new Map<string, WordTypedTarget[]>();
    for (const target of [...targets, ...glossaryTargets]) {
        const canonical = canonicalPartName(target.part);
        const candidates = byPart.get(canonical) ?? [];
        candidates.push(target);
        byPart.set(canonical, candidates);
    }
    const rank = (segment: string): number => {
        const index = REPORTING_SEGMENT_ORDER.indexOf(segment);
        return index === -1 ? REPORTING_SEGMENT_ORDER.length : index;
    };
    for (const candidates of byPart.values()) {
        const { part } = candidates[0]!;
        const resolved = resolveContentType(table, part);
        if (resolved === null || !GENERIC_CONTENT_TYPES.has(contentTypeMediaType(resolved.contentType))) {
            continue;
        }
        const { segment, via } = candidates.reduce((best, candidate) => (rank(candidate.segment) < rank(best.segment) ? candidate : best));
        const by =
            segment === OFFICE_DOCUMENT_SEGMENT
                ? "the package officeDocument relationship"
                : `the ${via === "glossary" ? "glossary" : "main document"}'s ${segment} relationship`;
        const contentType = asciiLowercase(resolved.contentType);
        const resolvedText =
            resolved.source === "default"
                ? `resolves only through the Default for extension "${partExtension(part)}" to "${contentType}"`
                : `is typed "${contentType}" by an Override`;
        const needed = expectedContentTypeText(segment, resolved.source);
        // Word recovery was observed only for the probed parts (#506).
        // New segments have only the Default application/xml observation;
        // avoid claiming an observation for other inferred generic spellings.
        const verdict =
            ADDITIONAL_OBSERVED_CONTENT_TYPES.has(segment) || SPEC_ONLY_CONTENT_TYPES.has(segment)
                ? ""
                : " Word 16.115 offers to recover the package.";
        findings.push({
            segment,
            segments: candidates.map((candidate) => candidate.segment),
            source: resolved.source,
            issue: {
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_GENERIC,
                path: part,
                message: `${part}, named by ${by}, ${resolvedText}; ${needed}.${verdict}`,
            },
        });
    }
    return findings;
};

export const wordPartContentTypeGenericCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_GENERIC,
    description: "Parts Word reads carry their own content type, not a generic XML type.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext): Promise<ValidationIssue[]> =>
        (await wordPartContentTypeGenericFindings(ctx)).map((finding) => finding.issue),
};
