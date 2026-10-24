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
 *    (#539 session 16, header; #549 session 17, the rest). A header it
 *    relates but does not reference is not read (clean open), and
 *    references from inside a footnote are not followed (session 16);
 *  - a part under `word/` that any relationship of a KNOWN Office type names
 *    with none of the observed types above (people, chart, image, …). These
 *    are unobserved, so they keep the old `word/`-folder behaviour —
 *    blocking — and open no new false negative. A type has a segment only
 *    when it is exactly an Office relationship-type namespace plus one
 *    segment, so `…/relationships/custom/styles` is an unobserved type too,
 *    not the `styles` type (#519).
 *
 * Compute once per run, as `propertiesParts` (#409) is, and pass the set to
 * `isWordBlockingIssue`, so full and partial runs agree (#381). Evidence:
 * docs/plans/evidence/word-probes-2026-10-24.md, sessions 3c, 4, 5, 8b, 9a;
 * session 15 (the main document's own relationships part, #514; header and
 * endnotes lazy, theme and glossaryDocument eager, #516); session 16 (the
 * glossary's header reference, #539); session 17 (the glossary's
 * relationships part and its comments, footnotes and endnotes, #549).
 */

import path from "node:path";

import { canonicalPartName, equivalentPartLookup, findEquivalentPart, isRelationshipsPartName } from "./content-types";
import type { ValidationContext } from "./validation-context";
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
const officeRelationshipTypeSegment = (type: string): string | undefined => {
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
/** An absolute URI names no package part, whatever its TargetMode says. */
const hasUriScheme = (target: string): boolean => /^[a-z][a-z\d+.-]*:/i.test(target);

/** The folder a relationships part's targets resolve against: its source part's folder. */
const relationshipsBaseDir = (relsPart: string): string => {
    const dir = path.posix.dirname(relsPart);
    return path.posix.basename(dir).toLowerCase() === "_rels" ? path.posix.dirname(dir) : dir;
};

/**
 * Resolve one relationship target against the folder of its source part, as
 * the rest of the codebase does (`\` reads as `/`, a query or fragment is
 * dropped, dot segments resolve), then take the OPC-equivalent stored entry
 * (`canonicalPartName` comparison); `undefined` when it names no stored part.
 */
const resolveTarget = (lookup: (partName: string) => string | undefined, baseDir: string, target: string): string | undefined => {
    if (hasUriScheme(target)) {
        return undefined;
    }
    const [reference = ""] = target.replace(/\\/g, "/").split(/[?#]/, 1);
    if (reference === "") {
        return undefined;
    }
    const joined = reference.startsWith("/") ? reference : path.posix.join(baseDir, reference);
    return lookup(joined);
};

/** The relationships part of a source part, found by OPC part-name equivalence. */
const sidecarOf = (names: readonly string[], part: string): string | undefined =>
    findEquivalentPart(names, path.posix.join(path.posix.dirname(part), "_rels", `${path.posix.basename(part)}.rels`));

/** Each relationship a part's relationships part holds, with its parsed DOM, or `undefined`. */
const relationshipsOf = (context: ValidationContext, relsPart: string | undefined): readonly Element[] => {
    if (relsPart === undefined) {
        return [];
    }
    const parsed = context.xml(relsPart);
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
const packageReach = async (context: ValidationContext): Promise<PackageReach | undefined> => {
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
    const packageRels = context.xml(packageRelsName);
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
    /** The relationship type's last path segment, or `officeDocument` for the main document. */
    readonly segment: string;
};

/** Canonical names of the main document parts, wherever they are stored. */
export const wordMainDocumentParts = async (context: ValidationContext): Promise<ReadonlySet<string>> => {
    const reach = await packageReach(context);
    return new Set((reach?.mainParts ?? []).map(canonicalPartName));
};

/**
 * The parts the named relationship segments reach (#432): every main document
 * (`officeDocument`), and the target of every main-document relationship of
 * one of the other segments, Transitional or Strict. One entry per part, in
 * traversal order; a part reached twice keeps the segment that reached it
 * first.
 */
export const wordTypedRelationshipTargets = async (
    context: ValidationContext,
    segments: ReadonlySet<string>,
): Promise<readonly WordTypedTarget[]> => {
    const reach = await packageReach(context);
    if (reach === undefined) {
        return [];
    }
    const targets: WordTypedTarget[] = [];
    const seen = new Set<string>();
    const add = (part: string, segment: string): void => {
        const canonical = canonicalPartName(part);
        if (seen.has(canonical)) {
            return;
        }
        seen.add(canonical);
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
        for (const relationship of relationshipsOf(context, sidecarOf(reach.names, mainPart))) {
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
 * The parts a story root — the main document or a reached glossary document
 * — reads through its own relationships part: the sidecar itself, and each
 * relationship's stored target as `{ part, segment }`. `followEager` adds
 * the targets of eagerly parsed types and is a main-document rule only: an
 * eager relationship from a glossary is unobserved (#549).
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
    followEager: boolean,
): Promise<readonly WordTypedTarget[]> => {
    const targets: WordTypedTarget[] = [{ part: root, segment: "officeDocument" }];
    const sidecar = sidecarOf(reach.names, root);
    if (sidecar !== undefined) {
        // Word reads a story root's own relationships part to reach its
        // stories, and offers recovery when it does not parse — the main
        // document's (#514, session 15) and a reached glossary's (#549,
        // session 17) alike.
        targets.push({ part: sidecar, segment: "relationships" });
    }
    const bytes = context.bytes(root);
    const readable = bytes !== undefined && !(await checkXmlSyntax(bytes, root)).some((issue) => issue.severity === "error");
    const story = readable ? context.effectiveXml(root) : undefined;
    const footerIds = story?.ok ? referenceIds(story.document, "footerReference") : undefined;
    const headerIds = story?.ok ? referenceIds(story.document, "headerReference") : undefined;
    const usesComments = story?.ok ? hasReference(story.document, "commentReference") : undefined;
    const usesFootnotes = story?.ok ? hasReference(story.document, "footnoteReference") : undefined;
    const usesEndnotes = story?.ok ? hasReference(story.document, "endnoteReference") : undefined;
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
    for (const relationship of relationshipsOf(context, sidecar)) {
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
            if (followEager) {
                targets.push({ part: target, segment });
            }
        } else if (LAZY_TYPES.has(segment) && lazyUsed(segment, relationship.getAttribute("Id") ?? "") !== false) {
            targets.push({ part: target, segment });
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
export const wordReadParts = async (context: ValidationContext): Promise<ReadonlySet<string>> => {
    const parts = new Set<string>();
    const reach = await packageReach(context);
    if (reach === undefined) {
        return parts;
    }
    const { names, lookup } = reach;
    // Glossary documents reached eagerly through a main-document
    // glossaryDocument relationship, canonical name to stored spelling, for
    // the second story-root pass below (#539).
    const glossaryParts = new Map<string, string>();

    for (const mainPart of reach.mainParts) {
        for (const { part, segment } of await storyRootReadParts(context, reach, mainPart, true)) {
            parts.add(canonicalPartName(part));
            if (segment === "glossaryDocument") {
                glossaryParts.set(canonicalPartName(part), part);
            }
        }
    }

    // Each reached glossary document is a second story root (#549, session
    // 17): its own relationships part is read — even when the glossary
    // references nothing (unclosed, zero-byte and unreferenced sidecars all
    // get repair) — and so are the header, footer, comments, footnotes and
    // endnotes its effective XML uses, as a main document's are. A header it
    // relates but does not reference is not read (clean open), and
    // references from inside a footnote are not followed (session 16). A
    // glossary that fails byte-level XML checks leaves usage unknown, as an
    // unparseable main does (#536), so its lazy targets stay reached.
    for (const glossaryPart of glossaryParts.values()) {
        for (const { part } of await storyRootReadParts(context, reach, glossaryPart, false)) {
            parts.add(canonicalPartName(part));
        }
    }

    // Unobserved types keep the old `word/`-folder behaviour: a part under
    // word/ that a known Office relationship names from anywhere stays
    // blocking. From any part: a header's image or a settings part's people
    // list are named this way.
    for (const name of names) {
        if (!isRelationshipsPartName(name)) {
            continue;
        }
        const baseDir = relationshipsBaseDir(name);
        for (const relationship of relationshipsOf(context, name)) {
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
            const canonical = canonicalPartName(target);
            if (canonical.startsWith("word/")) {
                parts.add(canonical);
            }
        }
    }

    return parts;
};
