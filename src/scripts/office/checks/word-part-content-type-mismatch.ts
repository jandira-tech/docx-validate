/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A related part declared with a different mapped part's media type (#600).
 * Session 24 observed Word 16.115 recover main-document header/footer swaps
 * and a reached glossary footer swap, including a parameterized swap.
 * Session 26 observed the reached glossary header swap. On the eagerly
 * parsed reaches — the main document and the main document's own styles,
 * settings and numbering relationships — every recovery sentence names the
 * session that showed its pair: session 49 the styles part typed with the
 * settings type (s49b); session 50 the main document and a related settings
 * part typed with the styles type and a related numbering part typed with
 * the settings type (s50d); session 53 the other mapped types it ran — on
 * the main document (s53d–s53f) and on the settings, numbering and styles
 * parts (s53f–s53h); session 56 the styles part typed with the footer type
 * (s56a); and session 57 the thirteen pairs no earlier session had shown —
 * the main document typed webSettings, the settings part typed webSettings,
 * glossary, footer, footnotes or endnotes, the numbering part typed
 * webSettings, glossary, theme or footer, and the styles part typed
 * webSettings, glossary or footnotes; and session 59 (s59a) the remaining
 * three pairs — the styles part typed endnotes and the numbering part
 * typed footnotes or endnotes — so all 48 eager pairs are observed. Session 53
 * (s53a/s53b) also observed the co-reach: an eagerly read part — the main
 * document, a related numbering part or the styles part — that a read
 * footer or a second eager relationship of another kind also reaches,
 * whatever the Override says, and session 57 (s57c–s57e) the header and
 * footer parts themselves: a footer the body references typed with the
 * main document, numbering, styles or settings type and a header the body
 * references typed with the main document or styles type, including the
 * exact-typed styles part a read header relationship targets and a
 * reached glossary's styles part a main footer relationship also reaches,
 * while typings no session showed on a header or footer keep their
 * verdict. Session 58 (s58b) showed the other read stories: a used
 * comments, endnotes or footnotes part — its own root, the body's
 * reference matching — typed with the styles type gets the recovery
 * offer, while the exact-typed styles part a comments, endnotes or
 * footnotes relationship nothing references also targets opens clean
 * (s58a/s58b), so the observed shape is the read story typed styles,
 * not the styles part co-reached. Word-valid reports those observed
 * shapes — the swaps Word reads through the swapping relationship, the
 * eagerly parsed reaches, which need no usage lookup, and the read
 * story co-reach of the observed pairs; fontTable, webSettings, theme
 * and every glossary reach were not shown a mapped type, the comments,
 * footnotes and endnotes reaches only the styles type, exact (session 58)
 * or with a parameter (session 62 s62c, #802), and a story part typed
 * with a type no session showed stays silent until probed. Strict judges
 * every mapped relationship.
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
import { NS, type Profile, type ValidationIssue } from "../../../lib/types";
import type { ValidationContext, XmlPart } from "../../../lib/validation-context";
import { wordReadTypedTargets } from "../../../lib/word-read-parts";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { REPORTING_SEGMENT_ORDER } from "./word-part-content-type-generic";
import {
    expectedContentTypesFor,
    LAZY_OBSERVED_SEGMENTS,
    SPEC_SEGMENTS,
    wordContentTypeTargets,
    type WordPartContentTypeFinding,
} from "./word-part-content-type-inexact";

export const WORD_PART_CONTENT_TYPE_MISMATCH = "word-part-content-type-mismatch";
// Evidence: word-probes-2026-10-24.md, sessions 24, 26, 49, 50, 53, 56, 57
// and 59. The header/footer keys are the swaps Word was shown, judged
// read. The three-part keys name the eager pairs — the reach plus the
// declared mapped type's own segment — and the session that showed it:
// sessions 49 and 50 one pair per reach (s49b, s50d), session 53 the other
// mapped types it ran on the main document and the settings, numbering and
// styles reaches (s53d–s53h), session 56 the footer type on the styles
// reach (s56a), session 57 the thirteen pairs no earlier session had shown
// (s57a–s57c), and session 59 the remaining three (s59a). All 48 eager
// pairs have an entry; the lookup stays, so a pair without one would keep
// its verdict with no recovery sentence.
const OBSERVED_SESSIONS: ReadonlyMap<string, string> = new Map([
    ["main\nfooter", "session 24"],
    ["main\nheader", "session 24"],
    ["glossary\nfooter", "session 24"],
    ["glossary\nheader", "session 26"],
    ["main\nofficeDocument\nstyles", "session 50"],
    ["main\nofficeDocument\nsettings", "session 53"],
    ["main\nofficeDocument\nnumbering", "session 53"],
    ["main\nofficeDocument\nfontTable", "session 53"],
    ["main\nofficeDocument\nwebSettings", "session 57"],
    ["main\nofficeDocument\ntheme", "session 53"],
    ["main\nofficeDocument\nheader", "session 53"],
    ["main\nofficeDocument\nfooter", "session 53"],
    ["main\nofficeDocument\nfootnotes", "session 53"],
    ["main\nofficeDocument\nendnotes", "session 53"],
    ["main\nofficeDocument\ncomments", "session 53"],
    ["main\nofficeDocument\nglossaryDocument", "session 53"],
    ["main\nstyles\nofficeDocument", "session 53"],
    ["main\nstyles\nsettings", "session 49"],
    ["main\nstyles\nnumbering", "session 53"],
    ["main\nstyles\nfontTable", "session 53"],
    ["main\nstyles\nwebSettings", "session 57"],
    ["main\nstyles\ntheme", "session 53"],
    ["main\nstyles\nheader", "session 53"],
    ["main\nstyles\nfooter", "session 56"],
    ["main\nstyles\nfootnotes", "session 57"],
    ["main\nstyles\nendnotes", "session 59"],
    ["main\nstyles\ncomments", "session 53"],
    ["main\nstyles\nglossaryDocument", "session 57"],
    ["main\nsettings\nofficeDocument", "session 53"],
    ["main\nsettings\nstyles", "session 50"],
    ["main\nsettings\nnumbering", "session 53"],
    ["main\nsettings\nfontTable", "session 53"],
    ["main\nsettings\nwebSettings", "session 57"],
    ["main\nsettings\ntheme", "session 53"],
    ["main\nsettings\nheader", "session 53"],
    ["main\nsettings\nfooter", "session 57"],
    ["main\nsettings\nfootnotes", "session 57"],
    ["main\nsettings\nendnotes", "session 57"],
    ["main\nsettings\ncomments", "session 53"],
    ["main\nsettings\nglossaryDocument", "session 57"],
    ["main\nnumbering\nofficeDocument", "session 53"],
    ["main\nnumbering\nstyles", "session 53"],
    ["main\nnumbering\nsettings", "session 50"],
    ["main\nnumbering\nfontTable", "session 53"],
    ["main\nnumbering\nwebSettings", "session 57"],
    ["main\nnumbering\ntheme", "session 57"],
    ["main\nnumbering\nheader", "session 53"],
    ["main\nnumbering\nfooter", "session 57"],
    ["main\nnumbering\nfootnotes", "session 59"],
    ["main\nnumbering\nendnotes", "session 59"],
    ["main\nnumbering\ncomments", "session 53"],
    ["main\nnumbering\nglossaryDocument", "session 57"],
]);
/**
 * The eagerly parsed reaches Word was shown refusing another mapped type
 * (#722): the main document and the main document's own styles, settings
 * and numbering relationships. Session 59 (s59a) observed the last three
 * pairs, so every pair on these reaches has an `OBSERVED_SESSIONS` entry;
 * a pair without one would keep its verdict without the recovery sentence.
 */
const OBSERVED_REACHES: ReadonlySet<string> = new Set(["main\nofficeDocument", "main\nstyles", "main\nsettings", "main\nnumbering"]);
/** The relationship segment a mapped media type belongs to, e.g. "styles" for the styles type. */
const segmentForMediaType = (mediaType: string): string | undefined =>
    [...SPEC_SEGMENTS].find((segment) => expectedContentTypesFor(segment).some((type) => asciiLowercase(type) === mediaType));
/**
 * The read story co-reach, keyed by the segment and the declared mapped
 * type's own segment (#739, #778, #785). Session 53 (s53a) showed the
 * footer pairs on the main document, numbering and styles types — the
 * calibrations on eagerly read parts — and session 57 (s57c–s57e) each
 * pair's own shape: a footer part typed styles, main document, numbering
 * or settings (s57c), a header part typed main document (s57c) or
 * styles (s57d), a reached glossary's styles part a main footer
 * relationship also reaches (s57d), and the exact-typed styles part a
 * read header relationship targets (s57e). Session 58 (s58b) showed the
 * comments, endnotes and footnotes stories typed with the styles type,
 * while the exact-typed styles part a comments, endnotes or footnotes
 * relationship nothing references also targets opens clean (s58a/s58b)
 * — the observed shape is the read story typed styles, not the styles
 * part co-reached. Session 58's rows were own-root parts, so its pairs
 * co-reach only a part whose effective root is the story's own (#803).
 * Session 63 (s63b) also refused a comments part whose stored root is an
 * mc:AlternateContent unwrapping to the comments root, typed with the
 * styles type exactly; that shape cites session 63, and the endnotes and
 * footnotes parts in it, or any inexact spelling of it, are unprobed and
 * keep the verdict they had before #785 (#818). Word refused
 * every shape whatever the Override says, so a matching reach on the
 * same part exempts none of it, and typings
 * no session showed keep their verdict. Main-document reaches only: a
 * glossary's own relationship co-reaches nothing here. The session-58
 * pairs also take the styles media type with a parameter, which session
 * 62 (s62c) refused; their finding cites that session instead of
 * session 58 (#802); session 62's rows stored the story's own root. A
 * padded spelling with no parameter, and any inexact spelling on the
 * header and footer pairs, stay unprobed and keep their verdict.
 */
const READ_STORY_COREACH_SESSIONS: ReadonlyMap<string, string> = new Map([
    ["footer\nofficeDocument", "sessions 53 and 57"],
    ["footer\nnumbering", "sessions 53 and 57"],
    ["footer\nstyles", "sessions 53 and 57"],
    ["footer\nsettings", "session 57"],
    ["header\nofficeDocument", "session 57"],
    ["header\nstyles", "session 57"],
    ["comments\nstyles", "session 58"],
    ["endnotes\nstyles", "session 58"],
    ["footnotes\nstyles", "session 58"],
]);
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
        // Whether the declared value spells its media type exactly, with no
        // parameter or padding; the session-58 co-reach pairs below also
        // take one with a parameter (#802).
        const exactSpelling = asciiLowercase(declared) === mediaType;
        // Generic XML and unknown types belong outside this check.
        if (!MAPPED_MEDIA_TYPES.has(mediaType)) continue;
        const matches = (segment: string): boolean => expectedContentTypesFor(segment).some((type) => asciiLowercase(type) === mediaType);
        const swaps = entries.filter(
            ({ segment }) =>
                (segment === "header" || segment === "footer") &&
                mediaType === asciiLowercase(expectedContentTypesFor(segment === "header" ? "footer" : "header")[0]!),
        );
        // Read story XML for a possible swap, the read story co-reach, or
        // competing relationship types. Unrelated parts need no usage
        // lookup; an eager read also interferes with staged-directory
        // write-back failures after repair (#412).
        if (swaps.length > 0 || entries.length > 1 || entries.some(({ segment }) => LAZY_OBSERVED_SEGMENTS.has(segment))) {
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
        // The read story co-reach (#739, #778, #785): a header, footer,
        // comments, footnotes or endnotes relationship the body references
        // whose part is typed with one of the observed pairs' mapped types,
        // which no matching reach exempts — the main document's reaches
        // only, never a glossary's own relationship. Session 58 showed
        // own-root parts, so its pairs co-reach only a part whose effective
        // root is the story's own (#803).
        const isOwnRoot = (element: Element | null, segment: string): boolean =>
            element?.localName === segment && WORD_PARAGRAPH_NAMESPACES.includes(element.namespaceURI ?? "");
        const rootIsOwn = (parsed: XmlPart | undefined, segment: string): boolean =>
            isOwnRoot(parsed?.ok === true ? parsed.document.documentElement : null, segment);
        const childElements = (element: Element): Element[] =>
            Array.from(element.childNodes).filter((child): child is Element => child.nodeType === 1);
        // The session a read story co-reach cites, or undefined when `entry`
        // is none.
        const storyCoreachSession = (entry: (typeof entries)[number]): string | undefined => {
            const session = READ_STORY_COREACH_SESSIONS.get(`${entry.segment}\n${segmentForMediaType(mediaType)}`);
            if (
                session === undefined ||
                !LAZY_OBSERVED_SEGMENTS.has(entry.segment) ||
                entry.via === "glossary" ||
                readKeys?.has(`${entry.via ?? "main"}\n${canonical}\n${entry.segment}`) !== true
            ) {
                return undefined;
            }
            if (session !== "session 58") return exactSpelling ? session : undefined;
            if (!rootIsOwn(ctx.effectiveXml(part), entry.segment)) return undefined;
            // Sessions 58 and 62 stored the story's own root: the exact
            // spelling cites session 58, and the styles media type with a
            // parameter session 62 (s62c), while a padded spelling with no
            // parameter is unprobed (#802).
            const stored = ctx.xml(part);
            if (rootIsOwn(stored, entry.segment)) {
                return exactSpelling ? session : declared.includes(";") ? "session 62" : undefined;
            }
            // A stored mc:AlternateContent root whose branches hold nothing
            // but the story's root: session 63 (s63b) refused the comments
            // part typed styles exactly. The other stories and spellings,
            // and other stored roots the effective view unwraps to the
            // story's root, such as an ignorable wrapper that processes its
            // content, alone or inside a branch, are unprobed (#818, #831).
            const storedRoot = stored?.ok === true ? stored.document.documentElement : null;
            const branchesHoldOwnRoot =
                storedRoot?.localName === "AlternateContent" &&
                storedRoot.namespaceURI === NS.MC &&
                childElements(storedRoot).every(
                    (branch) =>
                        branch.namespaceURI === NS.MC &&
                        (branch.localName === "Choice" || branch.localName === "Fallback") &&
                        childElements(branch).every((child) => isOwnRoot(child, entry.segment)),
                );
            return entry.segment === "comments" && exactSpelling && branchesHoldOwnRoot ? "session 63" : undefined;
        };
        const isReadStoryCoreach = (entry: (typeof entries)[number]): boolean => storyCoreachSession(entry) !== undefined;
        // The observed set covers the read header/footer swaps, the eagerly
        // parsed reaches — the main document and the main document's own
        // styles, settings and numbering relationships, which need no usage
        // lookup — and the read story co-reach. Header and footer
        // keep the swap shape `swaps` pins, read through the swapping
        // segment, and the citation names the session that showed the
        // finding's pair.
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
                    OBSERVED_REACHES.has(`${via ?? "main"}\n${segment}`),
            ) ??
            entries.find(isReadStoryCoreach);
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
                : isReadStoryCoreach(observed)
                  ? storyCoreachSession(observed)
                  : swaps.includes(observed)
                    ? OBSERVED_SESSIONS.get(`${observed.via ?? "main"}\n${observed.segment}`)
                    : OBSERVED_SESSIONS.get(`${observed.via ?? "main"}\n${observed.segment}\n${segmentForMediaType(mediaType)}`);
        const recovery = session === undefined ? "" : ` Word 16.115 offers to recover the package (${session}).`;
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
