/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The parts Microsoft Word reads at open, through relationships (#437).
 *
 * Word 16.115 does not read a part because it lies under `word/`: an
 * unreferenced truncated `word/xtra.xml` opens clean even with its own
 * Override, and so does one named only by a relationship of an unknown type.
 * What Word reads is what its relationships reach:
 *
 *  - the main document, wherever it is stored, and the parts a main-document
 *    relationship of an EAGER type names (styles, settings, fontTable,
 *    webSettings, numbering, customXml) — parsed at open, used or not;
 *  - the parts a main-document relationship of a LAZY type names, when the
 *    main document uses them (footer through a `w:footerReference r:id`,
 *    comments through any `w:commentReference`, footnotes through any
 *    `w:footnoteReference`) — related-but-unused opens clean;
 *  - a part under `word/` that any relationship of a KNOWN Office type names
 *    with none of the observed types above (header, endnotes, theme,
 *    glossaryDocument, people, chart, image, …). These are unobserved, so
 *    they keep the old `word/`-folder behaviour — blocking — and open no new
 *    false negative.
 *
 * Compute once per run, as `propertiesParts` (#409) is, and pass the set to
 * `isWordBlockingIssue`, so full and partial runs agree (#381). Evidence:
 * docs/plans/evidence/word-probes-2026-10-24.md, sessions 3c, 4, 5, 8b, 9a.
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
const EAGER_TYPES: ReadonlySet<string> = new Set(["styles", "settings", "fontTable", "webSettings", "numbering", "customXml"]);

/** Word parses a part of these types only when the document uses it. */
const LAZY_TYPES: ReadonlySet<string> = new Set(["footer", "comments", "footnotes"]);

/** Every relationship type whose reachability a probe has observed. */
const OBSERVED_TYPES: ReadonlySet<string> = new Set([...EAGER_TYPES, ...LAZY_TYPES]);

/** A relationship type in a namespace Office defines, whatever its segment. */
const isKnownOfficeNamespace = (type: string): boolean =>
    type.startsWith("http://schemas.openxmlformats.org/") ||
    type.startsWith("http://purl.oclc.org/ooxml/") ||
    type.startsWith("http://schemas.microsoft.com/office/");

/** A relationship type under one of the two Office relationship-type namespaces. */
const isOfficeRelationshipType = (type: string): boolean =>
    type.startsWith(RELATIONSHIP_TYPES_TRANSITIONAL) || type.startsWith(RELATIONSHIP_TYPES_STRICT);

/** The last path segment of a relationship-type URI: `…/relationships/styles` → `styles`. */
const typeSegment = (type: string): string => type.slice(type.lastIndexOf("/") + 1);

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

/** The relationship ids the main document's `w:footerReference` elements name. */
const footerReferenceIds = (document: Document): ReadonlySet<string> => {
    const ids = new Set<string>();
    for (const namespace of WORD_MAIN_NAMESPACES) {
        for (const elem of getElementsByTagNameNSAll(document, namespace, "footerReference")) {
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
 * Canonical names of the parts Word reads at open through relationships.
 * Mirrors `propertiesParts`: a package relationships part that does not parse
 * yields nothing (its own diagnostics already block the run, #388), and a
 * main document that does not parse keeps its lazily parsed parts reached —
 * whether it uses them cannot be decided, and no new false negative may open.
 */
export const wordReadParts = async (context: ValidationContext): Promise<ReadonlySet<string>> => {
    const parts = new Set<string>();
    const names = context.partNames;
    const lookup = equivalentPartLookup(names);

    const packageRelsName = findEquivalentPart(names, "_rels/.rels");
    if (packageRelsName === undefined) {
        return parts;
    }
    // Do not use relationships recovered by the DOM parser. An unreadable
    // package relationships part already blocks Word-valid on its own (#388).
    const bytes = context.bytes(packageRelsName);
    if (bytes === undefined || (await checkXmlSyntax(bytes, packageRelsName)).some((issue) => issue.severity === "error")) {
        return parts;
    }
    const packageRels = context.xml(packageRelsName);
    if (!packageRels?.ok) {
        return parts;
    }

    // The main document, wherever the package relationship stores it.
    const mainParts = new Set<string>();
    for (const relationship of getRelationshipChildren(packageRels.document)) {
        if (relationship.getAttribute("TargetMode") === "External" || !OFFICE_DOCUMENT_TYPES.has(relationship.getAttribute("Type") ?? "")) {
            continue;
        }
        // Package relationships resolve from the root, never from _rels/.
        const target = resolveTarget(lookup, ".", relationship.getAttribute("Target") ?? "");
        if (target === undefined) {
            continue;
        }
        mainParts.add(target);
        parts.add(canonicalPartName(target));
    }

    // Main-document relationships: eager types always, lazy types only when
    // the document uses them. Usage is read from the effective view Word
    // consumes, after markup compatibility has selected its branches.
    for (const mainPart of mainParts) {
        const main = context.has(mainPart) ? context.effectiveXml(mainPart) : undefined;
        const footerIds = main?.ok ? footerReferenceIds(main.document) : undefined;
        const usesComments = main?.ok ? hasReference(main.document, "commentReference") : undefined;
        const usesFootnotes = main?.ok ? hasReference(main.document, "footnoteReference") : undefined;

        const sidecarName = findEquivalentPart(
            names,
            path.posix.join(path.posix.dirname(mainPart), "_rels", `${path.posix.basename(mainPart)}.rels`),
        );
        if (sidecarName === undefined) {
            continue;
        }
        const sidecar = context.xml(sidecarName);
        if (!sidecar?.ok) {
            continue;
        }

        const baseDir = path.posix.dirname(mainPart);
        for (const relationship of getRelationshipChildren(sidecar.document)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const type = relationship.getAttribute("Type") ?? "";
            if (!isOfficeRelationshipType(type)) {
                continue;
            }
            const target = resolveTarget(lookup, baseDir, relationship.getAttribute("Target") ?? "");
            if (target === undefined) {
                continue;
            }
            const segment = typeSegment(type);
            if (EAGER_TYPES.has(segment)) {
                parts.add(canonicalPartName(target));
            } else if (segment === "footer") {
                if (footerIds === undefined || footerIds.has(relationship.getAttribute("Id") ?? "")) {
                    parts.add(canonicalPartName(target));
                }
            } else if (segment === "comments") {
                if (usesComments === undefined || usesComments) {
                    parts.add(canonicalPartName(target));
                }
            } else if (segment === "footnotes") {
                if (usesFootnotes === undefined || usesFootnotes) {
                    parts.add(canonicalPartName(target));
                }
            }
        }
    }

    // Unobserved types keep the old `word/`-folder behaviour: a part under
    // word/ that a known Office relationship names from anywhere stays
    // blocking. From any part: a header's image, a settings part's people
    // list, or the main document's own theme are all named this way.
    for (const name of names) {
        if (!isRelationshipsPartName(name)) {
            continue;
        }
        const rels = context.xml(name);
        if (!rels?.ok) {
            continue;
        }
        const baseDir = relationshipsBaseDir(name);
        for (const relationship of getRelationshipChildren(rels.document)) {
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const type = relationship.getAttribute("Type") ?? "";
            if (!isKnownOfficeNamespace(type) || OBSERVED_TYPES.has(typeSegment(type))) {
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
