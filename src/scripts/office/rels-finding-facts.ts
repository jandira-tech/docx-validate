/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Structured facts for the relationships findings whose word-valid verdict
 * `isWordBlockingIssue` once read from message text (#734). The
 * file-reference walk (`validateFileReferences`, validators/base.ts) embeds
 * the raw relationship `Target` in its `rels-broken` message, so a package
 * could spell the blocking text into a tolerated finding. The producer
 * therefore registers the structural fact on the issue object it returns,
 * through a module-level `WeakMap` keyed by the issue object itself, exactly
 * as the story-reference hand-off carries its identity (#616,
 * story-reference-identity.ts). The public `ValidationIssue` shape stays
 * unchanged, and an issue reaches the predicate unwrapped only if some
 * intermediate copy replaced it — such a finding has no fact and keeps its
 * word-valid warning.
 */

import type { ValidationIssue } from "../../lib/types";

/** Facts by issue object, so a copied issue is an unregistered issue (#734). */
const brokenTargets = new WeakMap<ValidationIssue, string>();

/**
 * Register the part name a broken relationship target resolved to, in the
 * walk's own `resolvePartName` spelling. The verdict never reads the raw
 * Target the message embeds.
 */
export const registerRelsBrokenTarget = (issue: ValidationIssue, resolvedTarget: string): void => {
    brokenTargets.set(issue, resolvedTarget);
};

/** The registered resolved part name of `issue`, or `undefined` — an unregistered finding never blocks (#734). */
export const relsBrokenTarget = (issue: ValidationIssue): string | undefined => brokenTargets.get(issue);
