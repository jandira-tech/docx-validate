/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Structured identity for the story reference hand-off (#616).
 *
 * `validateAllRelationshipIds` (validators/base.ts) and the
 * `word-header-footer-relationship` rule (checks/word-header-footer.ts) can
 * both report the same `headerReference`/`footerReference` element, and the
 * `word-part-reference-relationship` rule (checks/word-part-reference.ts)
 * the same `c:chart`/`o:OLEObject` (#705), the
 * `word-attached-toolbar-relationship` rule (checks/word-attached-toolbar.ts)
 * the same `wne:toolbarData` (#716) and the `word-hyperlink-relationship`
 * rule (checks/word-hyperlink.ts) the same `w:hyperlink` (#1089). A full
 * non-lenient run keeps the Word rule's finding and drops the legacy
 * `rels-id-mismatch` copy (#616), or the `rels-id-missing` copy for an
 * `r:id` naming no Relationship (#619), through
 * `withoutRepeatedHeaderFooterFindings` (validators/docx.ts). A full
 * non-lenient PPTX run hands a slide master's `p:sldLayoutId` to
 * `sldlayout-missing-rid` (validators/pptx.ts) the same way (#1093). Both
 * filter through `withoutHandedOffCopies` below. An `r:id` is well-formed XML with any character in
 * it, so the hand-off must never parse the two findings' message prose:
 * quoted ids defeated the former regexes (a double report) and a crafted
 * `r:id` could imitate another element's message prefix (a silent loss).
 * Both producers therefore register the element they just reported into a
 * module-level `WeakMap` keyed by the issue object itself, and the helper
 * reads the identity back from the very same objects. The public
 * `ValidationIssue` shape stays unchanged, and an issue reaches the helper
 * unwrapped only if some intermediate copy replaced it — such a finding has
 * no identity and is never dropped.
 */

import type { ValidationIssue, ValidationResult } from "../../lib/types";

/** The story reference, Word part reference and slide layout id local names the hand-off recognizes. */
export type StoryReferenceLocalName =
    | "headerReference"
    | "footerReference"
    | "chart"
    | "OLEObject"
    | "toolbarData"
    | "hyperlink"
    | "sldLayoutId";

/** `true` for the story reference, Word part reference and slide layout id local names the hand-off recognizes. */
export const isStoryReferenceLocalName = (name: string): name is StoryReferenceLocalName =>
    name === "headerReference" ||
    name === "footerReference" ||
    name === "chart" ||
    name === "OLEObject" ||
    name === "toolbarData" ||
    name === "hyperlink" ||
    name === "sldLayoutId";

/**
 * The story reference element a finding names: the part path (backslashes
 * normalized to `/`, as `withoutRepeatedCompatibilityFindings` compares
 * paths), the element's namespace URI, its local name, the raw `r:id`
 * attribute value, and the namespace URI that attribute was read in. The
 * Word rule reads the markup-compatibility view and both relationships
 * namespaces, the legacy walk the raw part, in both namespaces outside
 * lenient (#1065), so an `mc:Choice` reference with an ISO Strict `r:id`
 * and its `mc:Fallback` twin with a Transitional one differ only in that
 * last field (review BE F1); the hand-off's pairing deliberately ignores
 * that field (#727), so both spellings of one element pair as one defect.
 */
export interface StoryReferenceIdentity {
    readonly part: string;
    readonly namespaceURI: string;
    readonly localName: StoryReferenceLocalName;
    readonly rid: string;
    readonly ridNamespaceURI: string;
    /**
     * The element's other non-empty id values, when one finding stands for
     * ids in both relationships namespaces: session 55's dual-namespace
     * `c:chart` (#756, #1074), and every element a Word rule judges by its
     * first non-empty id (#1079). Since #1065 the legacy walk makes a copy
     * for each id outside lenient, and the hand-off drops the copy for any
     * of them.
     */
    readonly otherRids?: readonly string[];
}

/**
 * The non-empty `id` values `elem` carries in `namespaces` other than the
 * judged `rid`: the `otherRids` of a finding that stands for the whole
 * element (#1074, #1079).
 */
export const otherRelationshipIds = (elem: Element, namespaces: readonly string[], rid: string): string[] =>
    namespaces.map((ns) => elem.getAttributeNS(ns, "id")).filter((value): value is string => !!value && value !== rid);

/**
 * The qualified name `elem` spells its `localName` attribute with in the
 * relationships namespace `namespaceURI`, so a Word rule's message names the
 * attribute the part holds, as `rels-id-missing` does (#1090, #1101): `rs:id`
 * for the Strict namespace bound to `rs`, `r:id` for it bound to `r`. An empty
 * `namespaceURI` names a literal `r:` attribute.
 */
export const relationshipAttributeName = (elem: Element, namespaceURI: string, localName = "id"): string =>
    (namespaceURI ? elem.getAttributeNodeNS(namespaceURI, localName)?.name : undefined) ?? `r:${localName}`;

/** Identity by issue object, so a copied issue is an unregistered issue (#616). */
const identities = new WeakMap<ValidationIssue, StoryReferenceIdentity>();

/**
 * Register the story reference element an issue names. Both producers call
 * this with the issue object they return; the part path is normalized here
 * so both sides compare through one spelling.
 */
export const registerStoryReferenceIdentity = (issue: ValidationIssue, identity: StoryReferenceIdentity): void => {
    identities.set(issue, { ...identity, part: identity.part.replace(/\\/g, "/") });
};

/** The registered identity of `issue`, or `undefined` — an unregistered finding is never dropped. */
export const storyReferenceIdentity = (issue: ValidationIssue): StoryReferenceIdentity | undefined => identities.get(issue);

/**
 * `relationshipIds` without each `rels-id-mismatch` or `rels-id-missing` copy
 * of an element that a finding in `findings` whose code is in `producers`
 * already reported. A copy is dropped when its registered identity has that
 * finding's part, element namespace URI and local name, and either its `rid`
 * or one of its `otherRids` (#616, #1074, #1079). The namespace the id was
 * read in is not compared (#727). A copy with no registered identity is
 * never dropped. `valid` is kept, as the producer's own error fails the run.
 */
export const withoutHandedOffCopies = (
    relationshipIds: ValidationResult,
    findings: ValidationResult,
    producers: ReadonlySet<string>,
): ValidationResult => {
    const key = (identity: StoryReferenceIdentity): string =>
        `${identity.part}\u0000${identity.namespaceURI}\u0000${identity.localName}\u0000${identity.rid}`;
    const reported = new Set<string>();
    for (const issue of findings.issues) {
        if (issue.code === undefined || !producers.has(issue.code)) continue;
        const identity = storyReferenceIdentity(issue);
        if (!identity) continue;
        reported.add(key(identity));
        for (const rid of identity.otherRids ?? []) reported.add(key({ ...identity, rid }));
    }
    const issues = relationshipIds.issues.filter((issue) => {
        if (issue.code !== "rels-id-mismatch" && issue.code !== "rels-id-missing") return true;
        const identity = storyReferenceIdentity(issue);
        if (!identity) return true;
        return !reported.has(key(identity));
    });
    return { valid: relationshipIds.valid, issues };
};
