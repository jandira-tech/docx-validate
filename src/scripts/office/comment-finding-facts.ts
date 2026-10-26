/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Structured reference facts for comment-marker-missing findings (#784). */
import type { ValidationIssue } from "../../lib/types";

const missingReferences = new WeakSet<ValidationIssue>();

/** Register a missing comment named by w:commentReference, rather than a range marker alone. */
export const registerCommentReferenceMissing = (issue: ValidationIssue): void => {
    missingReferences.add(issue);
};

/** Unregistered findings keep their word-valid warning, as do range-only markers. */
export const isCommentReferenceMissing = (issue: ValidationIssue): boolean => missingReferences.has(issue);
