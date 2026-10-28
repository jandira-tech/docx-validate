/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The parts Microsoft Word reads at open, through relationships (#437), and
 * which relationship type reaches each of them (#432).
 *
 * Word 16.115 does not read a part because it lies under `word/`: an
 * unreferenced truncated `word/xtra.xml` opens clean even with its own
 * Override, and so does one named only by a relationship of an unknown type.
 * What Word reads is what its relationships reach:
 *
 *  - the main document, wherever it is stored, its own relationships part
 *    (found by OPC part-name equivalence, as `sidecarOf` finds it — Word
 *    reads it to reach styles, settings, headers and footers, and offers
 *    recovery when it is zero bytes or missing its root end tag, #514,
 *    session 15), and the parts a main-document relationship of an EAGER
 *    type names (styles, settings, fontTable, webSettings, numbering,
 *    customXml, theme, glossaryDocument — the last two eager since #516,
 *    session 15) — parsed at open, used or not;
 *  - the parts a main-document relationship of a LAZY type names, when the
 *    main document uses them (footer through a `w:footerReference r:id`,
 *    header through a `w:headerReference r:id`, comments through any
 *    `w:commentReference`, footnotes through any `w:footnoteReference`,
 *    endnotes through any `w:endnoteReference`) — related-but-unused opens
 *    clean, even under `word/` (#516, session 15);
 *  - a glossary document, reached through a main-document
 *    `glossaryDocument` relationship, is a second story root: its own
 *    relationships part is read even when the glossary references nothing
 *    (#549, session 17 — unclosed, zero-byte and unreferenced sidecars all
 *    get repair), and its `header` / `footer` relationships are used when
 *    its effective XML carries a `w:headerReference` / `w:footerReference`
 *    with that `r:id`, its `comments` / `footnotes` / `endnotes`
 *    relationships when it carries any `w:commentReference` /
 *    `w:footnoteReference` / `w:endnoteReference` — every referenced story
 *    gets repair when malformed, every well-formed control opens clean
 *    (#539 session 16, header; #549 session 17, the rest). Its `styles`,
 *    `settings`, `fontTable`, `webSettings`, `numbering` and `theme`
 *    relationships are eagerly parsed as a main document's are, even when
 *    the glossary references no story — every malformed part gets repair,
 *    every well-formed control opens clean (#562, session 19; the theme
 *    #569, session 20); `customXml` and a nested `glossaryDocument` from a
 *    glossary are observed, not read (session 20: malformed, both open
 *    clean) and stay unfollowed. A header it relates but does not reference
 *    is not read (clean open), and references from inside a footnote are
 *    not followed (session 16);
 *  - a part under `word/` that any relationship of a KNOWN Office type names
 *    with none of the observed types above (people, image, …). These
 *    are unobserved, so they keep the old `word/`-folder behaviour —
 *    blocking — and open no new false negative. A type has a segment only
 *    when it is exactly an Office relationship-type namespace plus one
 *    segment, so `…/relationships/custom/styles` is an unobserved type too,
 *    not the `styles` type (#519). The `chart` segment stayed in this set —
 *    `wordReadParts` is shared by the XML-syntax consumers — but
 *    `rels-missing-sidecar` alone now judges a chart-only reach by whether a
 *    Word-read part references the relationship (`wordReferenceLoadedSources`,
 *    #645, session 36), and a reach through the `image`, `diagramData`,
 *    `oleObject` and `hdphoto` types the same way (#647, session 44).
 *
 * A part two relationship types name yields one entry per type in the
 * typed walks and one typed read per type in `wordReadTypedTargets`, so the
 * order of the relationships in the relationships part cannot decide which
 * type a consumer judges or reads the part through (#593).
 *
 * Compute once per run, as `propertiesParts` (#409) is, and pass the set to
 * `isWordBlockingIssue`, so full and partial runs agree (#381). Evidence:
 * docs/plans/evidence/word-probes-2026-10-24.md, sessions 3c, 4, 5, 8b, 9a;
 * session 15 (the main document's own relationships part, #514; header and
 * endnotes lazy, theme and glossaryDocument eager, #516); session 16 (the
 * glossary's header reference, #539); session 17 (the glossary's
 * relationships part and its comments, footnotes and endnotes, #549);
 * session 19 (the glossary's eager styles, settings, fontTable, webSettings
 * and numbering, #562); session 20 (the glossary's theme is read too,
 * while its customXml and a nested glossaryDocument are observed, not
 * read, #569).
 *
 * The relationships these walks read are the ones markup compatibility
 * selects (#634, session 34 s34a/s34b): Word follows a relationship inside
 * an `mc:Fallback` — or an understood `mc:Choice` — exactly as it does a
 * direct child, so a truncated main document behind a Fallback
 * `officeDocument` relationship gets the open error its direct twin gets,
 * and so do the parts it and the package relationships reach. Under
 * `word-valid` and `strict` each relationships part is read through
 * `context.effectiveXml` (`understoodNamespacesForPart` keeps the
 * relationships vocabulary understood, #493); `getRelationshipChildren`
 * keeps its direct-children-only rule on that processed DOM, so a
 * relationship nested inside a foreign wrapper that survives processing
 * stays unfollowed (#434). `lenient` keeps the stored view, so it reports
 * what it always reported; a caller that passes no profile reads what Word
 * reads.
 */

import path from "node:path";

import {
    canonicalPartName,
    equivalentPartLookup,
    findEquivalentPart,
    hasUriScheme,
    isRelationshipsFolderName,
    isRelationshipsPartName,
    resolveTarget,
} from "./content-types";
import { NS, type Profile } from "./types";
import type { ValidationContext, XmlPart } from "./validation-context";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "./xml-helpers";
import { checkXmlSyntax } from "./xml-syntax";

/** Relationship-type namespaces, Transitional and Strict-conformance. */
const RELATIONSHIP_TYPES_TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const RELATIONSHIP_TYPES_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

/** The namespaces an `r:id` (or its Strict twin) attribute lives in. */
const RELATIONSHIP_ID_NAMESPACES: readonly string[] = [
    RELATIONSHIP_TYPES_TRANSITIONAL.slice(0, -1),
    RELATIONSHIP_TYPES_STRICT.slice(0, -1),
];

/** The WordprocessingML main namespace, Transitional and Strict — as in scripts/office/word-namespaces.ts. */
const WORD_MAIN_NAMESPACES: readonly string[] = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
];

const OFFICE_DOCUMENT_TYPES: ReadonlySet<string> = new Set([
    `${RELATIONSHIP_TYPES_TRANSITIONAL}officeDocument`,
    `${RELATIONSHIP_TYPES_STRICT}officeDocument`,
]);

/** Word parses a part of these types at open, whether the document uses it. */
const EAGER_TYPES: ReadonlySet<string> = new Set([
    "styles",
    "settings",
    "fontTable",
    "webSettings",
    "numbering",
    "customXml",
    "theme",
    "glossaryDocument",
]);

/** Word parses a part of these types only when the document uses it. */
const LAZY_TYPES: ReadonlySet<string> = new Set(["footer", "header", "comments", "footnotes", "endnotes"]);

/**
 * The eagerly parsed types a reached glossary's relationships are observed
 * to follow (#562, session 19; #569, session 20): real Word glossaries carry
 * their own styles, settings, fontTable and webSettings parts, and a
 * malformed one — numbering and theme included — gets repair. `customXml`
 * and a nested `glossaryDocument` from a glossary are observed, not read
 * (session 20: malformed, both open clean), so they stay unfollowed.
 */
const GLOSSARY_EAGER_TYPES: ReadonlySet<string> = new Set(["styles", "settings", "fontTable", "webSettings", "numbering", "theme"]);

/** Every relationship type whose reachability a probe has observed. */
const OBSERVED_TYPES: ReadonlySet<string> = new Set([...EAGER_TYPES, ...LAZY_TYPES]);

/** A relationship type in a namespace Office defines, whatever its segment. */
const isKnownOfficeNamespace = (type: string): boolean =>
    type.startsWith("http://schemas.openxmlformats.org/") ||
    type.startsWith("http://purl.oclc.org/ooxml/") ||
    type.startsWith("http://schemas.microsoft.com/office/");

/**
 * The segment of a relationship type that is exactly an Office
 * relationship-type namespace plus one segment: `…/relationships/styles` →
 * `styles`, Transitional or Strict. `undefined` for every other type — a
 * type with an extra path segment such as `…/relationships/custom/styles`
 * is an unknown type inside the namespace, not the `styles` type (#519).
 */
export const officeRelationshipTypeSegment = (type: string): string | undefined => {
    const prefix = type.startsWith(RELATIONSHIP_TYPES_TRANSITIONAL)
        ? RELATIONSHIP_TYPES_TRANSITIONAL
        : type.startsWith(RELATIONSHIP_TYPES_STRICT)
          ? RELATIONSHIP_TYPES_STRICT
          : undefined;
    if (prefix === undefined) {
        return undefined;
    }
    const segment = type.slice(prefix.length);
    return segment.length > 0 && !segment.includes("/") ? segment : undefined;
};
/**
 * The folder a relationships part's targets resolve against: its source
 * part's folder, the `_rels` segment compared canonically — the same
 * equivalence the sidecar lookup uses — so an escaped spelling resolves
 * as `_rels` does (#772).
 */
const relationshipsBaseDir = (relsPart: string): string => {
    const dir = path.posix.dirname(relsPart);
    return isRelationshipsFolderName(path.posix.basename(dir)) ? path.posix.dirname(dir) : dir;
};

/** The relationships part of a source part, found by OPC part-name equivalence. */
export const sidecarOf = (names: readonly string[], part: string): string | undefined =>
    findEquivalentPart(names, path.posix.join(path.posix.dirname(part), "_rels", `${path.posix.basename(part)}.rels`));

/**
 * The relationships part as the run's profile reads it: the stored DOM
 * under `lenient`, the markup-compatibility view `effectiveXml` builds
 * otherwise (#632, #634) — these helpers model what Word reads, so a
 * caller with no profile to give gets Word's view, and only lenient opts
 * into the stored one.
 */
const relationshipsXml = (context: ValidationContext, part: string, profile?: Profile): XmlPart | undefined =>
    profile === "lenient" ? context.xml(part) : context.effectiveXml(part);

/** Each relationship a part's relationships part holds, with the DOM the profile reads, or `undefined`. */
const relationshipsOf = (context: ValidationContext, relsPart: string | undefined, profile?: Profile): readonly Element[] => {
    if (relsPart === undefined) {
        return [];
    }
    const parsed = relationshipsXml(context, relsPart, profile);
    return parsed?.ok ? getRelationshipChildren(parsed.document) : [];
};

/** The relationship ids the main document's `w:<localName>` elements name. */
const referenceIds = (document: Document, localName: string): ReadonlySet<string> => {
    const ids = new Set<string>();
    for (const namespace of WORD_MAIN_NAMESPACES) {
        for (const elem of getElementsByTagNameNSAll(document, namespace, localName)) {
            for (const relNamespace of RELATIONSHIP_ID_NAMESPACES) {
                const id = elem.getAttributeNS(relNamespace, "id");
                if (id) {
                    ids.add(id);
                    break;
                }
            }
            // A literal `r:id` attribute is how a document without the
            // namespace declaration spells it; read it too, as
            // word-attached-toolbar does.
            const literal = elem.getAttribute("r:id");
            if (literal) {
                ids.add(literal);
            }
        }
    }
    return ids;
};

/** Whether the main document holds a `w:<localName>` reference, in either W namespace. */
const hasReference = (document: Document, localName: string): boolean =>
    WORD_MAIN_NAMESPACES.some((namespace) => getElementsByTagNameNSAll(document, namespace, localName).length > 0);

/**
 * A story root's effective XML when its bytes pass the byte-level XML
 * checks; `undefined` when they fail or the effective view cannot be read,
 * which leaves the root's usage undecided (#437, #536).
 */
const readableStory = async (context: ValidationContext, root: string): Promise<Document | undefined> => {
    const bytes = context.bytes(root);
    const readable = bytes !== undefined && !(await checkXmlSyntax(bytes, root)).some((issue) => issue.severity === "error");
    const story = readable ? context.effectiveXml(root) : undefined;
    return story?.ok ? story.document : undefined;
};

/** The reference element that uses each comments and notes segment. */
const STORY_REFERENCES: ReadonlyMap<string, string> = new Map([
    ["comments", "commentReference"],
    ["footnotes", "footnoteReference"],
    ["endnotes", "endnoteReference"],
]);

/**
 * Whether the story root `root` is known to use its `segment` part — a
 * comments, footnotes or endnotes reference in its readable effective XML.
 * The read model keeps a lazy target whose usage it cannot decide (#437),
 * so a rule that needs the use itself asks here (#858); `false` for an
 * unreadable root and for any other segment.
 */
export const storyRootUses = async (context: ValidationContext, root: string, segment: string): Promise<boolean> => {
    const reference = STORY_REFERENCES.get(segment);
    if (reference === undefined) {
        return false;
    }
    const story = await readableStory(context, root);
    return story !== undefined && hasReference(story, reference);
};

/** The package traversal the reachability views share: the names, the lookup, the main documents. */
type PackageReach = {
    readonly names: readonly string[];
    readonly lookup: (partName: string) => string | undefined;
    /** Stored spellings of the main document parts. */
    readonly mainParts: readonly string[];
};

/**
 * Read the package relationships part and find the main documents. As
 * `propertiesParts` does (#409, #388), a package relationships part that is
 * missing or does not parse yields nothing: its own diagnostics already block
 * the run, and no view should trust relationships the DOM parser recovered.
 */
const packageReach = async (context: ValidationContext, profile?: Profile): Promise<PackageReach | undefined> => {
    const names = context.partNames;
    const lookup = equivalentPartLookup(names);
    const packageRelsName = findEquivalentPart(names, "_rels/.rels");
    if (packageRelsName === undefined) {
        return undefined;
    }
    const bytes = context.bytes(packageRelsName);
    if (bytes === undefined || (await checkXmlSyntax(bytes, packageRelsName)).some((issue) => issue.severity === "error")) {
        return undefined;
    }
    const packageRels = relationshipsXml(context, packageRelsName, profile);
    if (!packageRels?.ok) {
        return undefined;
    }
    const mainParts = new Set<string>();
    for (const relationship of getRelationshipChildren(packageRels.document)) {
        if (relationship.getAttribute("TargetMode") === "External" || !OFFICE_DOCUMENT_TYPES.has(relationship.getAttribute("Type") ?? "")) {
            continue;
        }
        // Package relationships resolve from the root, never from _rels/.
        const target = resolveTarget(lookup, ".", relationship.getAttribute("Target") ?? "");
        if (target !== undefined) {
            mainParts.add(target);
        }
    }
    return { names, lookup, mainParts: [...mainParts] };
};

/** One relationship target, with the type segment that reached it. */
export type WordTypedTarget = {
    /** The stored part-name spelling, as the package view holds it. */
    readonly part: string;
    /**
     * The relationship type's one segment under an Office relationship-type
     * namespace, as `officeRelationshipTypeSegment` reads it (#519) — a type
     * with any other shape reaches no typed target — or `officeDocument` for
     * a story root and `relationships` for its own relationships part (#549).
     */
    readonly segment: string;
    /**
     * `"glossary"` when a reached glossary document's own relationship —
     * not the main document's — names the part (#579). Absent for
     * main-document targets.
     */
    readonly via?: "glossary";
    /**
     * The stored name of the reached glossary document whose own
     * relationships named the part — set on `via: "glossary"` entries so a
     * consumer can judge the part through the root that reads it (#698).
     * Absent for main-document targets.
     */
    readonly root?: string;
};

/**
 * The parts the named relationship segments reach (#432): every main document
 * (`officeDocument`), and the target of every main-document relationship of
 * one of the other segments, Transitional or Strict. One entry per
 * (part, segment) pair, in traversal order: a part reached twice through the
 * same type keeps the entry that reached it first, while a part reached
 * through two types keeps both, so the order of the relationships in the
 * relationships part cannot decide which type a consumer judges the part
 * against (#593). A target climbing past the package root names the part
 * its climb stops at, as `resolveTarget` resolves every target (#986).
 */
export const wordTypedRelationshipTargets = async (
    context: ValidationContext,
    segments: ReadonlySet<string>,
    profile?: Profile,
): Promise<readonly WordTypedTarget[]> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return [];
    }
    const targets: WordTypedTarget[] = [];
    const seen = new Set<string>();
    const add = (part: string, segment: string): void => {
        const key = `${canonicalPartName(part)}\n${segment}`;
        if (seen.has(key)) {
            return;
        }
        seen.add(key);
        // The stored spelling is what the finding reports, as every other
        // check reports the part names the package view holds.
        targets.push({ part, segment });
    };
    if (segments.has("officeDocument")) {
        for (const mainPart of reach.mainParts) {
            add(mainPart, "officeDocument");
        }
    }
    const namedSegments = [...segments].filter((segment) => segment !== "officeDocument");
    if (namedSegments.length === 0) {
        return targets;
    }
    const wanted = new Set(namedSegments);
    for (const mainPart of reach.mainParts) {
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, mainPart), profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const segment = officeRelationshipTypeSegment(relationship.getAttribute("Type") ?? "");
            if (segment === undefined || !wanted.has(segment)) {
                continue;
            }
            const target = resolveTarget(reach.lookup, path.posix.dirname(mainPart), relationship.getAttribute("Target") ?? "");
            if (target !== undefined) {
                add(target, segment);
            }
        }
    }
    return targets;
};

/**
 * Canonical names of the main documents themselves, wherever they are
 * stored, and of every part their own relationships parts reach, whatever
 * the relationship's type (#742): the exclusion set of the glossary
 * story-styles map, so a reached glossary can claim as a story part never a
 * part the main document also reads. Ungated by usage, as
 * `wordTypedRelationshipTargets` reads the story segments — whether the
 * main document uses a lazily parsed part cannot decide whose story the
 * part is — and unfiltered by type, so a part the main document reaches
 * through a non-story type cannot be claimed through a story one. External
 * targets name no package part, and a package view that cannot be read
 * yields the empty set, as `wordTypedRelationshipTargets` does; its own
 * diagnostics block the run.
 */
export const wordMainReachedParts = async (context: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return new Set();
    }
    const parts = new Set<string>();
    for (const mainPart of reach.mainParts) {
        parts.add(canonicalPartName(mainPart));
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, mainPart), profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const target = resolveTarget(reach.lookup, path.posix.dirname(mainPart), relationship.getAttribute("Target") ?? "");
            if (target !== undefined) {
                parts.add(canonicalPartName(target));
            }
        }
    }
    return parts;
};

/**
 * The Microsoft relationship types that pair a comments part with its
 * extension parts (#690, #696). They live under `schemas.microsoft.com`,
 * so `wordTypedRelationshipTargets` never reaches their targets and
 * `wordExactTypeTargets` reads them by exact Type.
 */
export const COMMENTS_EXTENDED_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
export const COMMENTS_IDS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
export const COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";

/**
 * The parts a main-document relationship of exactly `type` names, in
 * traversal order, once each (#690). The Microsoft comment-thread types
 * (`commentsExtended`, `commentsIds`, `commentsExtensible`) carry no Office
 * relationship-type segment, so `wordTypedRelationshipTargets` cannot reach
 * them; this walk reads the same relationships parts the same way (External
 * targets skipped, each target resolved against its main document's folder
 * and taken as the OPC-equivalent stored entry) and compares the Type
 * verbatim. A glossary's relationships are never read.
 */
export const wordExactTypeTargets = async (context: ValidationContext, type: string, profile?: Profile): Promise<readonly string[]> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return [];
    }
    const targets = new Set<string>();
    for (const mainPart of reach.mainParts) {
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, mainPart), profile)) {
            if (relationship.getAttribute("TargetMode") === "External" || relationship.getAttribute("Type") !== type) {
                continue;
            }
            const target = resolveTarget(reach.lookup, path.posix.dirname(mainPart), relationship.getAttribute("Target") ?? "");
            if (target !== undefined) {
                targets.add(target);
            }
        }
    }
    return [...targets];
};

/**
 * The parts a reached glossary document's own relationships part names
 * (#579): every glossary a main-document `glossaryDocument` relationship
 * reaches is a second story root (#549, session 17), and each of its
 * relationships whose type segment is in `segments` yields its stored
 * target, marked `via: "glossary"` so a finding can name the glossary's
 * relationship rather than the main document's. One entry per
 * (part, segment) pair, in traversal order; a part reached twice keeps every
 * segment that reached it (#593). A nested `glossaryDocument` is observed,
 * not read (session 20),
 * so the walk does not recurse, and `officeDocument` never matches: a
 * glossary's relationship of that segment would misattribute the package
 * root. A part the main document also names by the same segment keeps
 * both entries, so a caller can judge it through whichever root reads it
 * (#593).
 */
export const glossaryTypedRelationshipTargets = async (
    context: ValidationContext,
    segments: ReadonlySet<string>,
    profile?: Profile,
): Promise<readonly WordTypedTarget[]> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return [];
    }
    const wanted = new Set([...segments].filter((segment) => segment !== "officeDocument"));
    if (wanted.size === 0) {
        return [];
    }
    // The reached glossaries, canonical name to stored spelling — the same
    // discovery `wordReadParts` makes for its second story-root pass.
    const glossaries = new Map<string, string>();
    for (const mainPart of reach.mainParts) {
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, mainPart), profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            if (officeRelationshipTypeSegment(relationship.getAttribute("Type") ?? "") !== "glossaryDocument") {
                continue;
            }
            const target = resolveTarget(reach.lookup, path.posix.dirname(mainPart), relationship.getAttribute("Target") ?? "");
            if (target !== undefined) {
                glossaries.set(canonicalPartName(target), target);
            }
        }
    }
    const targets: WordTypedTarget[] = [];
    const seen = new Set<string>();
    for (const glossary of glossaries.values()) {
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, glossary), profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const segment = officeRelationshipTypeSegment(relationship.getAttribute("Type") ?? "");
            if (segment === undefined || !wanted.has(segment)) {
                continue;
            }
            const target = resolveTarget(reach.lookup, path.posix.dirname(glossary), relationship.getAttribute("Target") ?? "");
            if (target === undefined) {
                continue;
            }
            const canonical = canonicalPartName(target);
            const key = `${canonical}\n${segment}`;
            if (canonical === canonicalPartName(glossary) || seen.has(key)) {
                continue;
            }
            seen.add(key);
            targets.push({ part: target, segment, via: "glossary", root: glossary });
        }
    }
    return targets;
};

/**
 * The parts a story root — the main document or a reached glossary document
 * — reads through its own relationships part: the sidecar itself, and each
 * relationship's stored target as `{ part, segment }`, marked `via` when the
 * root is a reached glossary (#593). `eagerTypes` names
 * the eagerly parsed segments whose targets are added: `EAGER_TYPES` for a
 * main document, `GLOSSARY_EAGER_TYPES` for a glossary — a glossary's
 * `styles`, `settings`, `fontTable`, `webSettings`, `numbering` and
 * `theme` are observed (#562, session 19; #569, session 20) while its
 * `customXml` and nested `glossaryDocument` are observed, not read
 * (session 20) and stay unfollowed (#549).
 *
 * Lazy usage comes from the effective XML Word consumes, after markup
 * compatibility has selected its branches, but only once the part's bytes
 * pass the XML checker — a DOM may recover bytes the checker rejects (for
 * example an unsupported declared encoding), and such a root cannot
 * establish that lazy targets are unused, so they stay reached in full,
 * partial and fallback runs (#536).
 */
const storyRootReadParts = async (
    context: ValidationContext,
    reach: PackageReach,
    root: string,
    eagerTypes: ReadonlySet<string>,
    via?: "glossary",
    profile?: Profile,
): Promise<readonly WordTypedTarget[]> => {
    const typed = (part: string, segment: string): WordTypedTarget =>
        via === undefined ? { part, segment } : { part, segment, via, root };
    const targets: WordTypedTarget[] = [typed(root, "officeDocument")];
    const sidecar = sidecarOf(reach.names, root);
    if (sidecar !== undefined) {
        // Word reads a story root's own relationships part to reach its
        // stories, and offers recovery when it does not parse — the main
        // document's (#514, session 15) and a reached glossary's (#549,
        // session 17) alike.
        targets.push(typed(sidecar, "relationships"));
    }
    const story = await readableStory(context, root);
    const footerIds = story ? referenceIds(story, "footerReference") : undefined;
    const headerIds = story ? referenceIds(story, "headerReference") : undefined;
    const usesComments = story ? hasReference(story, "commentReference") : undefined;
    const usesFootnotes = story ? hasReference(story, "footnoteReference") : undefined;
    const usesEndnotes = story ? hasReference(story, "endnoteReference") : undefined;
    // Whether a lazily parsed relationship's target is used; `undefined`
    // when that cannot be decided, which keeps the target reached (#437).
    const lazyUsed = (segment: string, id: string): boolean | undefined => {
        switch (segment) {
            case "footer":
                return footerIds?.has(id);
            case "header":
                return headerIds?.has(id);
            case "comments":
                return usesComments;
            case "footnotes":
                return usesFootnotes;
            case "endnotes":
                return usesEndnotes;
            default:
                return undefined;
        }
    };

    const baseDir = path.posix.dirname(root);
    for (const relationship of relationshipsOf(context, sidecar, profile)) {
        if (relationship.getAttribute("TargetMode") === "External") {
            continue;
        }
        const segment = officeRelationshipTypeSegment(relationship.getAttribute("Type") ?? "");
        if (segment === undefined) {
            continue;
        }
        const target = resolveTarget(reach.lookup, baseDir, relationship.getAttribute("Target") ?? "");
        if (target === undefined) {
            continue;
        }
        if (EAGER_TYPES.has(segment)) {
            if (eagerTypes.has(segment)) {
                targets.push(typed(target, segment));
            }
        } else if (LAZY_TYPES.has(segment) && lazyUsed(segment, relationship.getAttribute("Id") ?? "") !== false) {
            targets.push(typed(target, segment));
        }
    }
    return targets;
};

/**
 * The story roots — every main document, then every reached glossary — and
 * the typed parts each reads, computed once. A glossary's entries carry
 * `via: "glossary"` so a (part, segment) pair the main document and a
 * glossary both read stays distinguishable (#593).
 */
const wordStoryReadTargets = async (
    context: ValidationContext,
    profile?: Profile,
): Promise<{ readonly targets: readonly WordTypedTarget[]; readonly reach: PackageReach } | undefined> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return undefined;
    }
    const targets: WordTypedTarget[] = [];
    // Glossary documents reached eagerly through a main-document
    // glossaryDocument relationship, canonical name to stored spelling, for
    // the second story-root pass below (#539).
    const glossaryParts = new Map<string, string>();
    for (const mainPart of reach.mainParts) {
        for (const target of await storyRootReadParts(context, reach, mainPart, EAGER_TYPES, undefined, profile)) {
            targets.push(target);
            if (target.segment === "glossaryDocument") {
                glossaryParts.set(canonicalPartName(target.part), target.part);
            }
        }
    }
    // Each reached glossary document is a second story root (#549, session
    // 17): its own relationships part is read — even when the glossary
    // references nothing (unclosed, zero-byte and unreferenced sidecars all
    // get repair) — and so are the header, footer, comments, footnotes and
    // endnotes its effective XML uses, as a main document's are, and the
    // styles, settings, fontTable, webSettings, numbering and theme it
    // relates eagerly, used or not (#562, session 19; the theme #569,
    // session 20) — customXml and a nested glossaryDocument are observed,
    // not read (session 20) and stay unfollowed. A header it relates but
    // does not reference is not read (clean open), and references from
    // inside a footnote are not followed (session 16). A glossary that
    // fails byte-level XML checks leaves usage unknown, as an unparseable
    // main does (#536), so its lazy targets stay reached.
    for (const glossaryPart of glossaryParts.values()) {
        targets.push(...(await storyRootReadParts(context, reach, glossaryPart, GLOSSARY_EAGER_TYPES, "glossary", profile)));
    }
    return { targets, reach };
};

/**
 * Canonical names of the story roots Word reads at open: every main
 * document, wherever it is stored, plus every glossary document a
 * main-document `glossaryDocument` relationship reaches. Each is a root
 * whose own relationships part Word reads (#514, #549), so a duplicate Id
 * there makes Word offer recovery as the main document's does (#617,
 * session 30). Reads the one story-root traversal above — no second glossary
 * resolver — and an unreached glossary stays absent, as a header or an
 * orphan `word/document.xml` does.
 */
export const wordStoryRootParts = async (context: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> => {
    const parts = new Set<string>();
    for (const target of (await wordStoryReadTargets(context, profile))?.targets ?? []) {
        // Only a story root itself is typed `officeDocument`; no relationship
        // segment can add another such entry.
        if (target.segment === "officeDocument") {
            parts.add(canonicalPartName(target.part));
        }
    }
    return parts;
};

/**
 * The parts Word reads at open through relationships, as typed targets
 * (#593): each entry says which relationship type reads the part — and which
 * story root's, through `via` — so a part Word reads through one type but
 * not another keeps both answers, and a consumer that gates lazily parsed
 * relationships can decide per (part, segment) pair instead of per part. A
 * part read through two types yields an entry per type, in document order.
 * The unobserved-type pass adds no typed read and stays in `wordReadParts`.
 */
export const wordReadTypedTargets = async (context: ValidationContext, profile?: Profile): Promise<readonly WordTypedTarget[]> =>
    (await wordStoryReadTargets(context, profile))?.targets ?? [];

/**
 * Canonical names of the parts Word may read (#1116, #1125): every part the
 * read model reads, and every part an internal relationship names, from any
 * relationships part, wherever it is stored and whether or not Word reads its
 * source — except a header, footer, comments or notes relationship of a story
 * root's own relationships part that the read model reads, which names its
 * part only through the read model, so a story the root relates but does not
 * use stays out (session 84 s84a). The read model reads a relationships part
 * as the profile does (`relationshipsXml`): outside lenient, through the
 * markup-compatibility view, where a relationship inside a selected Fallback
 * or Choice, or a processed wrapper, is a direct child (#634). A stored
 * relationship is the one it reads when it is in the package-relationships
 * namespace and its `Id`, `Type` and `Target` match (#1158); session 86
 * (s86c) opened clean an unused header related through a selected Fallback
 * or an understood Choice, and refused it used. The read model decides a root's comments and
 * notes use from the root alone, and so does Word: session 85 (s85a, s85b)
 * opened clean a related comments, footnotes or endnotes part holding an
 * unknown `mc:MustUnderstand` whose only reference sat in a read header, and
 * refused it referenced from the body, so such a part stays out too (#1155,
 * undoing #1145's conservative half). Session 86 (s86a, s86b, s86c) showed
 * the same for a reference in a used footer, for a reached glossary's used
 * header, and for malformed compatibility (#1164). A reference inside a notes or comments part names nothing either,
 * as session 16 (s16a) opened clean a malformed endnotes part only a footnote
 * referenced. Every other `Relationship` element of the stored DOM counts,
 * nested ones included, so a relationship inside a Choice Word does not
 * select still names its part unless it repeats a read one: the scan errs wide, and only the exception
 * narrows the list, where the read model decides that a part is unused. That
 * holds in any namespace: the read model reads only the
 * package-relationships `Relationship` children of a `Relationships` root in
 * that namespace, in its view, so it decides nothing about one in no namespace, in
 * another, or under a root outside it, and that one names its part (#1135). A
 * target spelling a URI scheme names the part word-valid's
 * file-reference walk resolves it to: every scheme but `http:`, `https:` and
 * `mailto:` reads as a relative reference (#988, #1008), which the read model
 * does not follow, so a story root's own story relationship to one counts
 * (#1127). `undefined` when that cannot be decided: the package
 * relationships part is missing or does not parse, it names no stored main
 * document (#816), or a relationships part fails the byte-level XML checks.
 */
export const wordPossiblyReadParts = async (context: ValidationContext, profile?: Profile): Promise<ReadonlySet<string> | undefined> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined || reach.mainParts.length === 0) {
        return undefined;
    }
    const { names, lookup } = reach;
    const read = await wordReadTypedTargets(context, profile);
    const rootSidecars = new Set(
        read.filter((target) => target.segment === "relationships").map((target) => canonicalPartName(target.part)),
    );
    const named = new Set(read.map((target) => canonicalPartName(target.part)));
    for (const name of names) {
        if (!isRelationshipsPartName(name)) {
            continue;
        }
        const bytes = context.bytes(name);
        const parsed = context.xml(name);
        if (bytes === undefined || !parsed?.ok || (await checkXmlSyntax(bytes, name)).some((issue) => issue.severity === "error")) {
            return undefined;
        }
        const rootSidecar = rootSidecars.has(canonicalPartName(name));
        const baseDir = relationshipsBaseDir(name);
        // The read model's own relationships, matched by attributes: its view
        // is a markup-compatibility clone, which lifts a relationship inside a
        // selected branch or a processed wrapper to a direct child (#1158).
        const view = relationshipsXml(context, name, profile);
        const readable = new Set(view?.ok ? getRelationshipChildren(view.document).map(relationshipKey) : []);
        for (const relationship of getElementsByTagNameNSAll(parsed.document, "*", "Relationship")) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const reference = relationship.getAttribute("Target") ?? "";
            const target = resolveTarget(lookup, baseDir, reference);
            const segment = officeRelationshipTypeSegment(relationship.getAttribute("Type") ?? "");
            if (
                target !== undefined &&
                rootSidecar &&
                segment !== undefined &&
                LAZY_TYPES.has(segment) &&
                relationship.namespaceURI === NS.PR &&
                readable.has(relationshipKey(relationship))
            ) {
                continue;
            }
            const part = target ?? schemeTargetPart(lookup, baseDir, reference);
            if (part !== undefined) {
                named.add(canonicalPartName(part));
            }
        }
    }
    return named;
};

/**
 * The part a target spelling a URI scheme names under word-valid's
 * file-reference walk, which skips only `http:`, `https:` and `mailto:`,
 * ASCII case-insensitively, and resolves any other scheme as a relative
 * reference (#988, #1008). `undefined` for those three schemes and for a
 * target with no scheme, which `resolveTarget` reads.
 */
const schemeTargetPart = (lookup: (partName: string) => string | undefined, baseDir: string, target: string): string | undefined =>
    hasUriScheme(target) && !/^(?:https?|mailto):/i.test(target) ? resolveTarget(lookup, baseDir, `./${target}`) : undefined;

/**
 * A relationship's identity across a stored part and its markup-compatibility
 * clone. `TargetMode` is left out: an external relationship names nothing.
 */
const relationshipKey = (relationship: Element): string =>
    JSON.stringify(["Id", "Type", "Target"].map((attribute) => relationship.getAttribute(attribute)));

/** The lazily read stories, beside the story roots' own `officeDocument` entries. */
const STORY_SEGMENTS: ReadonlySet<string> = new Set(["header", "footer", "comments", "footnotes", "endnotes"]);

/**
 * One computation per package, gate state and profile: twenty-odd Word rules
 * walk the same list in a run, each through its own view when a part fails
 * the XML checks (#1128).
 */
const readStoryPartsCache = new WeakMap<object, Map<string, Promise<readonly string[]>>>();

/**
 * The story parts Word reads (#1106, #1113): every main document and
 * reached glossary wherever stored, and the header, footer, comments and
 * notes parts their references make Word read, each once, in the context's
 * part-name order, as the by-name list keeps them.
 * Word refused a hyperlink's shape (session 82) and eleven rules' shapes
 * (session 83) in a main document stored as `word/main.xml`, and opened them
 * clean in an unrelated `word/header9.xml` or `word/header83.xml`. With no
 * main document the read model reads nothing, so every part named like a
 * story stands in, headers, footers, comments and notes parts included; that
 * reading is unprobed. It is wider than the note checks' #816 fallback, which
 * judges references in none of those four kinds (#1119).
 */
export const wordReadStoryParts = (context: ValidationContext, profile?: Profile): Promise<readonly string[]> => {
    const byProfile = readStoryPartsCache.get(context.gateKey) ?? new Map<string, Promise<readonly string[]>>();
    readStoryPartsCache.set(context.gateKey, byProfile);
    const key = profile ?? "";
    const cached = byProfile.get(key);
    if (cached !== undefined) {
        return cached;
    }
    const parts = readStoryParts(context, profile);
    byProfile.set(key, parts);
    return parts;
};

const readStoryParts = async (context: ValidationContext, profile?: Profile): Promise<readonly string[]> => {
    const targets = await wordReadTypedTargets(context, profile);
    if (!targets.some((target) => target.segment === "officeDocument")) {
        return context.storyParts();
    }
    const read = new Map<string, string>();
    for (const target of targets) {
        const canonical = canonicalPartName(target.part);
        if ((target.segment === "officeDocument" || STORY_SEGMENTS.has(target.segment)) && !read.has(canonical)) {
            read.set(canonical, target.part);
        }
    }
    const order = new Map(context.partNames.map((name, index) => [name, index]));
    const position = (part: string): number => order.get(part) ?? context.partNames.length;
    return [...read.values()].sort((a, b) => position(a) - position(b));
};

/**
 * Canonical names of the parts a known Office relationship of an unobserved
 * type names from any relationships part, wherever they are stored: the
 * unobserved-type pass of `wordReadParts`, before its `word/` filter.
 */
const unobservedRelationshipTargets = (context: ValidationContext, reach: PackageReach, profile?: Profile): Set<string> => {
    const { names, lookup } = reach;
    const targets = new Set<string>();
    for (const name of names) {
        if (!isRelationshipsPartName(name)) {
            continue;
        }
        const baseDir = relationshipsBaseDir(name);
        for (const relationship of relationshipsOf(context, name, profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const type = relationship.getAttribute("Type") ?? "";
            // Only an exact namespace-plus-segment type can be an observed
            // type; anything else in a known Office namespace — including a
            // type with an extra path segment — is unobserved (#519).
            const segment = officeRelationshipTypeSegment(type);
            if (!isKnownOfficeNamespace(type) || (segment !== undefined && OBSERVED_TYPES.has(segment))) {
                continue;
            }
            const target = resolveTarget(lookup, baseDir, relationship.getAttribute("Target") ?? "");
            if (target === undefined) {
                continue;
            }
            targets.add(canonicalPartName(target));
        }
    }
    return targets;
};

/**
 * Canonical names of the parts Word reads at open through relationships.
 * Mirrors `propertiesParts`: a package relationships part that does not parse
 * yields nothing (its own diagnostics already block the run, #388), and a
 * main document that does not parse keeps its lazily parsed parts reached —
 * whether it uses them cannot be decided, and no new false negative may open.
 */
export const wordReadParts = async (context: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> => {
    const parts = new Set<string>();
    const story = await wordStoryReadTargets(context, profile);
    if (story === undefined) {
        return parts;
    }
    const { reach } = story;
    for (const { part } of story.targets) {
        parts.add(canonicalPartName(part));
    }

    // Unobserved types keep the old `word/`-folder behaviour: a part under
    // word/ that a known Office relationship names from anywhere stays
    // blocking. From any part: a header's image or a settings part's people
    // list are named this way.
    for (const canonical of unobservedRelationshipTargets(context, reach, profile)) {
        if (canonical.startsWith("word/")) {
            parts.add(canonical);
        }
    }

    return parts;
};

/** The Microsoft `hdphoto` relationship type (an HD Photo a picture's `a14:imgLayer` embeds). */
const HDPHOTO_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2007/relationships/hdphoto";

/**
 * How Word 16.115 loads the part a relationship of `type` reaches, for a part
 * whose own relationships part is missing (#645, #647):
 *  - `"chart"`: only when a Word-read part references the relationship — an
 *    unreferenced chart opened clean, a drawing's `c:chart r:id` got the
 *    recovery offer (sessions 36, 37 and 39);
 *  - `"unobserved-reference"`: a `diagramData`, `oleObject` or `hdphoto`
 *    reach nothing references opened clean, under `word/` and outside it
 *    (session 44 s44b/s44c), and so did a picture's `a:blip r:embed` naming
 *    one (session 45 s45b/s45d, #702 S1) — session 52's D1 row opened the
 *    same clean reach again, a chart part with no sidecar an `hdphoto`
 *    relationship a picture's `a:blip r:embed` names; every other
 *    reference shape was never shown to Word;
 *  - `"never"`: an `image` reach opened clean unreferenced and when a
 *    picture's `a:blip r:embed` names it (session 44 s44a/s44b), a VML
 *    picture's `v:imagedata r:id` and a linked picture's `a:blip r:link`
 *    (session 46 s46d) — session 52's E1 and E2 rows opened the same
 *    no-sidecar chart clean when a `w:name r:id` names the reach and when
 *    a VML picture's `v:imagedata r:id` inside a selected `mc:Fallback`
 *    does — Word reads the bytes as an image and never loads
 *    the target's relationships; the shapes Word refuses (`c:chart r:id`,
 *    `dgm:relIds`, `o:OLEObject r:id` naming the relationship) are refused
 *    for the reference's Type, whatever the target (#705, sessions 45 and
 *    46), so #701's extension came out again (#706);
 *  - `undefined`: every other type keeps the reach-based rule.
 * The Office types are read by segment, Transitional or ISO Strict, as #645
 * read the chart type.
 */
const referenceLoad = (type: string): "chart" | "unobserved-reference" | "never" | undefined => {
    if (type === HDPHOTO_RELATIONSHIP_TYPE) {
        return "unobserved-reference";
    }
    switch (officeRelationshipTypeSegment(type)) {
        case "chart":
            return "chart";
        case "diagramData":
        case "oleObject":
            return "unobserved-reference";
        case "image":
            return "never";
        default:
            return undefined;
    }
};

/**
 * One relationship that reaches a part through a type Word loads only when it
 * is referenced: the part its relationships part belongs to, and its `Id`.
 * `source` is `undefined` for `_rels/.rels`, which owns no part, and for a
 * relationships part that is not a `_rels` sidecar. `reference` says what
 * counts as a reference to the Id:
 *  - `"declared"` (`chart`): #645's reference attributes
 *    (`wordRelationshipReferenceIds`);
 *  - `"non-picture"` (`diagramData`, `oleObject`, `hdphoto`): any attribute
 *    naming the Id but a picture's `a:blip r:embed`, the one reference Word
 *    was shown clean (#702 S1, session 45; the hdphoto reach again in
 *    session 52 D1); their other reference shapes
 *    (`dgm:relIds r:dm`, …) are unobserved, so they keep blocking (#647).
 * An `image` relationship is never returned: Word never loads the part
 * through it, however its Id is named (#706, sessions 44 and 46).
 */
export type ReferenceLoadedSource = {
    readonly source: string | undefined;
    readonly id: string;
    readonly reference: "declared" | "non-picture";
};

/**
 * The relationships that reach `part` through a type Word loads only when
 * referenced, when every internal relationship that reaches it is of such a
 * type or of the `image` type — `undefined` otherwise, including when none
 * does (#645, #647). Session 36 (s36b/s36d) observed that Word 16.115 never
 * loads a chart no drawing references: a chart part related from the main
 * document's sidecar but referenced by nothing opened clean, while the same
 * chart an inline drawing's `c:chart r:id` referenced got the recovery offer;
 * session 37 (s37b) saw the same split for a drawing only in an unselected
 * `mc:Choice` (clean) and in a selected one (recovery), so the caller reads
 * the reference in the processed view. Session 44 opened clean every
 * unreferenced image reach and every one a picture's `a:blip r:embed` names,
 * and every unreferenced `diagramData`, `oleObject` and `hdphoto` reach
 * ({@link referenceLoad}); an image relationship is never returned, so a
 * part only images reach yields no sources (#706). The caller drops
 * the finding to a warning when no Word-read part's effective XML names one
 * of the returned Ids, and keeps it when `undefined` leaves the part on the
 * pre-#645 verdict.
 */
export const wordReferenceLoadedSources = async (
    context: ValidationContext,
    part: string,
    profile?: Profile,
): Promise<readonly ReferenceLoadedSource[] | undefined> => {
    const reach = await packageReach(context, profile);
    if (reach === undefined) {
        return undefined;
    }
    const { names, lookup } = reach;
    // `part` arrives as a finding's stored spelling; `\` reads as `/` as every
    // resolver here reads it, before the OPC-equivalence key is taken.
    const canonical = canonicalPartName(part.replace(/\\/g, "/"));
    const sources: ReferenceLoadedSource[] = [];
    const seen = new Set<string>();
    let reached = false;
    for (const name of names) {
        if (!isRelationshipsPartName(name)) {
            continue;
        }
        const baseDir = relationshipsBaseDir(name);
        // The part this relationships part belongs to, when it is a `_rels`
        // sidecar — the folder segment compared canonically, as the sidecar
        // lookup compares it, so an escaped spelling owns its source part
        // too (#772, #781); `undefined` for `_rels/.rels` (the package's
        // own) and for a `.rels` name no `_rels` folder owns.
        const dir = path.posix.dirname(name);
        const source = isRelationshipsFolderName(path.posix.basename(dir))
            ? lookup(path.posix.join(path.posix.dirname(dir), path.posix.basename(name).replace(/\.rels$/i, "")))
            : undefined;
        for (const relationship of relationshipsOf(context, name, profile)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const target = resolveTarget(lookup, baseDir, relationship.getAttribute("Target") ?? "");
            if (target === undefined || canonicalPartName(target) !== canonical) {
                continue;
            }
            const load = referenceLoad(relationship.getAttribute("Type") ?? "");
            if (load === undefined) {
                return undefined;
            }
            reached = true;
            if (load === "never") {
                continue;
            }
            const key = `${source ?? ""}\n${relationship.getAttribute("Id") ?? ""}\n${load}`;
            if (seen.has(key)) {
                continue;
            }
            seen.add(key);
            const reference = load === "chart" ? "declared" : "non-picture";
            sources.push({ source, id: relationship.getAttribute("Id") ?? "", reference });
        }
    }
    return reached ? sources : undefined;
};

/**
 * Canonical names of the parts a known Office relationship of an unobserved
 * type names, wherever they are stored (#638). `rels-missing-sidecar` blocked
 * in every part before #638, and Word was never shown such a part outside
 * `word/` with a reference and no sidecar, so that finding keeps blocking
 * there: no new false negative opens. `wordReadParts` keeps the `word/`-folder
 * rule its XML-syntax consumers have used since #437. A package with no main
 * document yields nothing (its own diagnostics already block the run, #388).
 */
export const wordUnobservedRelationshipTargets = async (context: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> => {
    const story = await wordStoryReadTargets(context, profile);
    return story === undefined ? new Set() : unobservedRelationshipTargets(context, story.reach, profile);
};
