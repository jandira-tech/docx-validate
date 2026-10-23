/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Runs a selection of named checks against one context and aggregates their
 * findings. Each selected check runs at most once. A check whose prerequisite
 * did not pass is skipped rather than run against input it cannot handle, a
 * check that throws becomes an error finding instead of crashing the run, and
 * an archive-only check on an unpacked directory is reported unavailable, not
 * passed.
 */

import type { CheckDefinition, CheckRegistry } from "./check-registry";
import type { PackageFormat, ValidationIssue } from "./types";
import type { ValidationContext } from "./validation-context";

export type CheckStatus = "passed" | "failed" | "skipped" | "unavailable";

export type CheckOutcome = {
    readonly id: string;
    readonly status: CheckStatus;
    readonly issues: readonly ValidationIssue[];
    /** Why a check was skipped or unavailable. */
    readonly reason?: string;
};

export type CheckRun = {
    readonly valid: boolean;
    readonly issues: readonly ValidationIssue[];
    readonly outcomes: readonly CheckOutcome[];
    /** Ids that ran (passed or failed). */
    readonly executed: readonly string[];
    /** Applicable ids the selection left out. Empty for a complete run. */
    readonly omitted: readonly string[];
    readonly skipped: readonly string[];
    readonly unavailable: readonly string[];
};

export type RunChecksOptions = {
    readonly format: PackageFormat;
    /** Check ids or group names. Omit to run every applicable check. */
    readonly select?: readonly string[];
};

export const runChecks = async (ctx: ValidationContext, registry: CheckRegistry, opts: RunChecksOptions): Promise<CheckRun> => {
    const applicable = registry.list(opts.format).map((def) => def.id);
    const selected = opts.select === undefined ? applicable : registry.resolve(opts.select, opts.format);

    const outcomes: CheckOutcome[] = [];
    const status = new Map<string, CheckStatus>();
    // Registration order puts every prerequisite before its dependents.
    for (const def of registry.list(opts.format).filter((candidate) => selected.includes(candidate.id))) {
        const outcome = await runOne(ctx, def, status);
        status.set(def.id, outcome.status);
        outcomes.push(outcome);
    }

    const issues = outcomes.flatMap((outcome) => outcome.issues);
    const ids = (wanted: CheckStatus | readonly CheckStatus[]): string[] =>
        outcomes
            .filter((outcome) => (typeof wanted === "string" ? [wanted] : wanted).includes(outcome.status))
            .map((outcome) => outcome.id);
    return {
        valid: issues.every((issue) => issue.severity !== "error"),
        issues,
        outcomes,
        executed: ids(["passed", "failed"]),
        omitted: applicable.filter((id) => !selected.includes(id)),
        skipped: ids("skipped"),
        unavailable: ids("unavailable"),
    };
};

const runOne = async (ctx: ValidationContext, def: CheckDefinition, status: ReadonlyMap<string, CheckStatus>): Promise<CheckOutcome> => {
    if (def.archiveOnly && ctx.archiveBytes === undefined) {
        return { id: def.id, status: "unavailable", issues: [], reason: "needs the packed archive; the package was not loaded from one" };
    }
    // Resolution includes every prerequisite, so each one already has a status.
    for (const prerequisite of def.requires ?? []) {
        const prior = status.get(prerequisite);
        if (prior !== "passed") {
            return { id: def.id, status: "skipped", issues: [], reason: `prerequisite ${prerequisite} ${prior}` };
        }
    }
    let issues: readonly ValidationIssue[];
    try {
        issues = await def.run(ctx);
    } catch (err) {
        issues = [
            {
                severity: "error",
                code: "check-crashed",
                message: `Check ${def.id} failed to run: ${err instanceof Error ? err.message : String(err)}`,
            },
        ];
    }
    return { id: def.id, status: issues.some((issue) => issue.severity === "error") ? "failed" : "passed", issues };
};
