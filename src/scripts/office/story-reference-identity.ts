/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Structured identity for the story reference hand-off (#616).
 *
 * `validateAllRelationshipIds` (validators/base.ts) and the
 * `word-header-footer-relationship` rule (checks/word-header-footer.ts) can
 * both report the same `headerReference`/`footerReference` element. A full
 * non-lenient run keeps the Word rule's finding and drops the legacy
 * `rels-id-mismatch` copy (#616), or the `rels-id-missing` copy for an
 * `r:id` naming no Relationship (#619), through
 * `withoutRepeatedHeaderFooterFindings` (validators/docx.ts). An `r:id` is well-formed XML with any character in
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

import type { ValidationIssue } from "../../lib/types";

/** The two story reference local names the hand-off recognizes. */
export type StoryReferenceLocalName = "headerReference" | "footerReference";

/** `true` for the two story reference local names the hand-off recognizes. */
export const isStoryReferenceLocalName = (name: string): name is StoryReferenceLocalName =>
    name === "headerReference" || name === "footerReference";

/**
 * The story reference element a finding names: the part path (backslashes
 * normalized to `/`, as `withoutRepeatedCompatibilityFindings` compares
 * paths), the element's namespace URI, its local name, the raw `r:id`
 * attribute value, and the namespace URI that attribute was read in. The
 * Word rule reads the markup-compatibility view and both relationships
 * namespaces, the legacy walk the raw part and only the Transitional one,
 * so an `mc:Choice` reference with an ISO Strict `r:id` and its
 * `mc:Fallback` twin with a Transitional one differ only in that last
 * field (review BE F1).
 */
export interface StoryReferenceIdentity {
    readonly part: string;
    readonly namespaceURI: string;
    readonly localName: StoryReferenceLocalName;
    readonly rid: string;
    readonly ridNamespaceURI: string;
}

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
