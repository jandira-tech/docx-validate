/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { ValidationIssue } from "./types";

/**
 * Drop the `*-parse` findings a check reports for a part the XML diagnostics
 * already rejected: the check could not read the part, which is the same
 * defect reported again.
 */
export const withoutConsequentialParseFailures = (issues: readonly ValidationIssue[], malformed: ReadonlySet<string>): ValidationIssue[] =>
    issues.filter((issue) => !(issue.path !== undefined && malformed.has(issue.path) && issue.code?.endsWith("-parse")));
