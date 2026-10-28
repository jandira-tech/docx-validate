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
 * comments, endnotes or footnotes part — its own Transitional root, the body's
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
import { citesWordSession } from "../../../lib/conformance-class";
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
import { type WordTypedTarget, wordReadTypedTargets } from "../../../lib/word-read-parts";
import { WORD_2006_NAMESPACE, WORD_PARAGRAPH_NAMESPACES, WORD_STRICT_NAMESPACE } from "../word-namespaces";
import { REPORTING_SEGMENT_ORDER } from "./word-part-content-type-generic";
import { xmlPartNames } from "./xml";
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
 * part co-reached. Sessions 58 and 62 stored Transitional story roots
 * (#803, #802); session 62 (s62c) showed the parameterized styles spelling.
 * Session 63 (s63b) showed a stored mc:AlternateContent with one w14 Choice
 * holding only the Transitional comments root, then an empty Fallback,
 * typed styles exactly (#818). Session 65 (s65b–s65d) changed one thing at
 * a time: text before that root in the Choice, text in the Fallback,
 * Requires naming the Transitional Word namespace, or a Strict comments
 * root in that Choice or stored directly in a Transitional package (#847).
 * Those shapes cite session 65; text directly between branches supports
 * no type citation because Word refused its comments-typed twin too (#851).
 * Other Requires namespaces, a Strict root in the Fallback, other text
 * positions within the Choice or Fallback or combinations of these edits,
 * CDATA anywhere in the wrapper (#857), the endnotes and footnotes twins
 * and the new shapes' other spellings are unprobed and keep their verdicts
 * without a citation. Every session cited here stored a package whose main
 * document is a Transitional w:document, so any other package keeps the
 * co-reach's verdict without a citation (#857, #861). Other effective-root wrappers remain unprobed for this type
 * mismatch (#831, #837, #843). Main-document reaches only: a glossary's own
 * relationship co-reaches nothing here. A padded styles spelling with no
 * parameter, and the header and footer pairs' inexact spellings, stay
 * unprobed and keep their verdict.
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
    let readTargets: readonly WordTypedTarget[] | undefined;
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
            readTargets ??= await wordReadTypedTargets(ctx, profile);
            readKeys ??= new Set(
                readTargets.map((target) => `${target.via ?? "main"}\n${canonicalPartName(target.part)}\n${target.segment}`),
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
        // Every session behind a co-reach citation stored a Transitional
        // package, each main document's stored root a Transitional
        // w:document (#857, #861).
        const transitionalMain = (): boolean => {
            const mains = (readTargets ?? []).filter(({ segment, via }) => segment === "officeDocument" && via === undefined);
            return (
                mains.length > 0 &&
                mains.every(({ part: main }) => {
                    const stored = ctx.xml(main);
                    const root = stored?.ok === true ? stored.document.documentElement : null;
                    return root?.localName === "document" && root.namespaceURI === WORD_2006_NAMESPACE;
                })
            );
        };
        // Keep the verdict of an unprobed co-reach without borrowing a
        // session's citation (#847).
        const storyCoreach = (entry: (typeof entries)[number]): { session?: string } | undefined => {
            const session = READ_STORY_COREACH_SESSIONS.get(`${entry.segment}\n${segmentForMediaType(mediaType)}`);
            if (
                session === undefined ||
                !LAZY_OBSERVED_SEGMENTS.has(entry.segment) ||
                entry.via === "glossary" ||
                readKeys?.has(`${entry.via ?? "main"}\n${canonical}\n${entry.segment}`) !== true
            ) {
                return undefined;
            }
            if (session !== "session 58") return exactSpelling ? { session } : undefined;
            if (!rootIsOwn(ctx.effectiveXml(part), entry.segment)) return undefined;
            // Sessions 58 and 62 stored Transitional story roots. Session
            // 65 (s65d) refused the exact-typed comments part with a Strict
            // root directly stored in a Transitional package (#847, #857).
            const stored = ctx.xml(part);
            const storedRoot = stored?.ok === true ? stored.document.documentElement : null;
            if (rootIsOwn(stored, entry.segment)) {
                if (!exactSpelling && !declared.includes(";")) return undefined;
                return {
                    session:
                        storedRoot?.namespaceURI === WORD_2006_NAMESPACE
                            ? exactSpelling
                                ? session
                                : "session 62"
                            : entry.segment === "comments" && exactSpelling && storedRoot?.namespaceURI === WORD_STRICT_NAMESPACE
                              ? "session 65"
                              : undefined,
                };
            }
            // Session 63 (s63b) stored one w14 Choice holding only the
            // Transitional comments root, then an empty Fallback. Session
            // 65 (s65b/s65c) changed one thing at a time: text before that
            // root, text in the Fallback, Requires naming w, or a Strict
            // root in the Choice of a Transitional package. Text between
            // branches is refused even typed comments (#851), so it
            // supports no type citation. No session stored CDATA, so
            // CDATA anywhere in the wrapper keeps the verdict without a
            // citation (#857).
            const [choice, fallback, ...rest] = storedRoot ? childElements(storedRoot) : [];
            if (
                entry.segment !== "comments" ||
                !exactSpelling ||
                storedRoot?.localName !== "AlternateContent" ||
                storedRoot.namespaceURI !== NS.MC ||
                rest.length !== 0 ||
                choice?.namespaceURI !== NS.MC ||
                choice.localName !== "Choice" ||
                !childElements(choice).every((child) => isOwnRoot(child, entry.segment)) ||
                fallback?.namespaceURI !== NS.MC ||
                fallback.localName !== "Fallback" ||
                childElements(fallback).length !== 0
            ) {
                return undefined;
            }
            const [storyRoot, ...otherRoots] = childElements(choice);
            const isText = (node: Node): boolean => node.nodeType === 3 && /[^ \t\r\n]/.test(node.nodeValue ?? "");
            if (Array.from(storedRoot.childNodes).some(isText)) return undefined;
            if ([storedRoot, choice, fallback].some((element) => Array.from(element.childNodes).some((node) => node.nodeType === 4)))
                return {};
            if (storyRoot === undefined || otherRoots.length !== 0) return {};
            const required = (choice.getAttribute("Requires") ?? "").trim().split(/[ \t\r\n]+/);
            if (required.length !== 1) return {};
            const requiredNamespace = choice.lookupNamespaceURI(required[0]!);
            const choiceText = Array.from(choice.childNodes).some(isText);
            const fallbackText = Array.from(fallback.childNodes).some(isText);
            const transitionalRoot = storyRoot?.namespaceURI === WORD_2006_NAMESPACE;
            if (requiredNamespace === NS.W14) {
                if (!choiceText && !fallbackText) {
                    return { session: transitionalRoot ? "session 63" : "session 65" };
                }
                const trailingText = Array.from(choice.childNodes)
                    .slice(Array.from(choice.childNodes).indexOf(storyRoot) + 1)
                    .some(isText);
                return { session: transitionalRoot && !trailingText && choiceText !== fallbackText ? "session 65" : undefined };
            }
            return {
                session:
                    requiredNamespace === WORD_2006_NAMESPACE && transitionalRoot && !choiceText && !fallbackText
                        ? "session 65"
                        : undefined,
            };
        };
        // A package whose main document is not Transitional keeps the
        // co-reach and its verdict, but no session showed it (#861), and so
        // does one with a Strict part outside the shape (#1016). Sessions 53
        // and 57 stored only Transitional parts. Sessions 58, 62, 63 and 65
        // stored the story part in the shape, and s65c and s65d stored its
        // root Strict, so the story part's own class is the shape's.
        const readStoryCoreach = (entry: (typeof entries)[number]): { session?: string } | undefined => {
            const coreach = storyCoreach(entry);
            if (coreach?.session === undefined) {
                return coreach;
            }
            const shapeParts = entry.segment === "header" || entry.segment === "footer" ? [] : [part];
            return transitionalMain() &&
                citesWordSession(
                    ctx,
                    xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
                    shapeParts,
                )
                ? coreach
                : {};
        };
        const storyCoreachSession = (entry: (typeof entries)[number]): string | undefined => readStoryCoreach(entry)?.session;
        const isReadStoryCoreach = (entry: (typeof entries)[number]): boolean => readStoryCoreach(entry) !== undefined;
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
        // The swap and eager-pair sessions stored Transitional packages: a
        // package with a Strict part keeps the finding without them (#1016).
        const cited = (key: string): string | undefined =>
            citesWordSession(
                ctx,
                xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
            )
                ? OBSERVED_SESSIONS.get(key)
                : undefined;
        const session =
            observed === undefined
                ? undefined
                : isReadStoryCoreach(observed)
                  ? storyCoreachSession(observed)
                  : swaps.includes(observed)
                    ? cited(`${observed.via ?? "main"}\n${observed.segment}`)
                    : cited(`${observed.via ?? "main"}\n${observed.segment}\n${segmentForMediaType(mediaType)}`);
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
