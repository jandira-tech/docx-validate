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
 * did not pass is skipped rather than run against input it cannot handle,
 * except that a failed part gate (`gatesByPart`) only hides the parts it
 * names from its dependents, a
 * check that throws becomes an error finding instead of crashing the run, and
 * an archive-only check on an unpacked directory is reported unavailable, not
 * passed.
 */

import type { CheckDefinition, CheckRegistry } from "./check-registry";
import type { PackageFormat, Profile, ValidationIssue } from "./types";
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
    /**
     * A check whose failure ends the run: nothing registered after it runs,
     * as a lenient run stops at its first malformed part (#391). A selected
     * check the halt kept from running is neither executed nor omitted.
     */
    readonly haltAfterFailing?: string;
    /**
     * The run's resolved profile, handed to each check's `run`: a check that
     * reads relationships parts uses the markup-compatibility view under
     * `word-valid` and `strict`, the stored view under `lenient` (#634).
     */
    readonly profile?: Profile;
};

export const runChecks = async (ctx: ValidationContext, registry: CheckRegistry, opts: RunChecksOptions): Promise<CheckRun> => {
    const applicable = registry.list(opts.format).map((def) => def.id);
    const selected = opts.select === undefined ? applicable : registry.resolve(opts.select, opts.format);

    const outcomes: CheckOutcome[] = [];
    const status = new Map<string, CheckStatus>();
    const byId = new Map<string, CheckOutcome>();
    // Registration order puts every prerequisite before its dependents.
    for (const def of registry.list(opts.format).filter((candidate) => selected.includes(candidate.id))) {
        const outcome = await runOne(ctx, def, status, byId, registry, opts.profile);
        status.set(def.id, outcome.status);
        byId.set(def.id, outcome);
        outcomes.push(outcome);
        if (def.id === opts.haltAfterFailing && outcome.status === "failed") {
            break;
        }
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

/** The parts a failed outcome's error findings name, with the first reason for each; `undefined` when one names no part. */
const partFailures = (outcome: CheckOutcome | undefined): Map<string, string> | undefined => {
    const failures = new Map<string, string>();
    for (const issue of outcome?.issues ?? []) {
        if (issue.severity !== "error") {
            continue;
        }
        if (issue.path === undefined) {
            return undefined;
        }
        if (!failures.has(issue.path)) {
            failures.set(issue.path, issue.message);
        }
    }
    return failures.size > 0 ? failures : undefined;
};

const runOne = async (
    ctx: ValidationContext,
    def: CheckDefinition,
    status: ReadonlyMap<string, CheckStatus>,
    outcomes: ReadonlyMap<string, CheckOutcome>,
    registry: CheckRegistry,
    profile?: Profile,
): Promise<CheckOutcome> => {
    if (def.archiveOnly && ctx.archiveBytes === undefined) {
        return { id: def.id, status: "unavailable", issues: [], reason: "needs the packed archive; the package was not loaded from one" };
    }
    // Resolution includes every prerequisite, so each one already has a status.
    const unparseable = new Map<string, string>();
    for (const prerequisite of def.requires ?? []) {
        const prior = status.get(prerequisite);
        if (prior === "passed") {
            continue;
        }
        // A failed part gate excludes the parts it names; the dependent still
        // judges the rest. Findings without a part leave nothing to exclude.
        const failed = prior === "failed" && registry.get(prerequisite)?.gatesByPart ? partFailures(outcomes.get(prerequisite)) : undefined;
        if (!failed) {
            return { id: def.id, status: "skipped", issues: [], reason: `prerequisite ${prerequisite} ${prior}` };
        }
        for (const [part, reason] of failed) {
            unparseable.set(part, reason);
        }
    }
    let issues: readonly ValidationIssue[];
    try {
        issues = await def.run(unparseable.size > 0 ? ctx.withUnparseable(unparseable) : ctx, profile);
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
