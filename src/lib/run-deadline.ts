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
 * The wall-clock budget of one validation run (#630).
 *
 * The bound is cooperative. Each step of the run asks {@link RunDeadline.reached}
 * before it starts: the package stages, every named check, every rule of the
 * format validator, every part of the schema walk, and Word verification. A
 * step that is already running is never interrupted, so a run overruns its
 * budget by at most its longest single step. A step the budget stopped is
 * recorded, and the run reports every recorded step in one error-severity
 * {@link VALIDATION_DEADLINE} finding: a run cut short is never a pass, whatever
 * the steps that did run found (timeout > engine error > invalid > valid).
 */

import type { ValidationIssue } from "./types";

/** Code of the one finding a run its deadline stopped reports. */
export const VALIDATION_DEADLINE = "validation-deadline";

/** How many stopped steps the finding names before it counts the rest. */
const NAMED_STEPS = 8;

/**
 * Reject a budget that is not a finite number of milliseconds, 0 or more.
 *
 * @throws {RangeError}
 */
export const assertDeadlineMs: (value: unknown) => asserts value is number = (value: unknown): asserts value is number => {
    if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new RangeError(`deadlineMs must be a finite number of milliseconds, 0 or more; got ${String(value)}`);
    }
};

export class RunDeadline {
    /** The budget the run was given, in milliseconds. */
    public readonly budgetMs: number;
    private readonly expiresAt: number;
    private readonly stopped: string[] = [];

    /**
     * Start the budget now. `now` is the clock the budget is measured on;
     * tests inject one that moves only when they move it.
     *
     * @throws {RangeError} when `budgetMs` is not a finite number, 0 or more.
     */
    public constructor(
        budgetMs: number,
        private readonly now: () => number = Date.now,
    ) {
        assertDeadlineMs(budgetMs);
        this.budgetMs = budgetMs;
        this.expiresAt = now() + budgetMs;
    }

    /** Whether the budget is spent. Records nothing. */
    public expired(): boolean {
        return this.now() >= this.expiresAt;
    }

    /**
     * Whether `step` must not start because the budget is spent. When it must
     * not, `step` is recorded as not run.
     */
    public reached(step: string): boolean {
        if (!this.expired()) {
            return false;
        }
        this.record(step);
        return true;
    }

    /** Record `step` as not run, for a caller that already knows the budget is spent. */
    public record(step: string): void {
        this.stopped.push(step);
    }

    /** Milliseconds left in the budget, never below zero. */
    public remainingMs(): number {
        return Math.max(0, this.expiresAt - this.now());
    }

    /** The steps the budget stopped, in the order they were stopped. */
    public get notRun(): readonly string[] {
        return [...this.stopped];
    }

    /**
     * The finding for the steps the budget stopped, or `undefined` when every
     * step ran: a budget spent by the last step still leaves a complete run.
     */
    public issue(): ValidationIssue | undefined {
        if (this.stopped.length === 0) {
            return undefined;
        }
        const named = this.stopped.slice(0, NAMED_STEPS);
        const more = this.stopped.length - named.length;
        const steps = more > 0 ? `${named.join("; ")}; and ${more} more` : named.join("; ");
        const count = this.stopped.length === 1 ? "1 step" : `${this.stopped.length} steps`;
        return {
            severity: "error",
            code: VALIDATION_DEADLINE,
            message:
                `Validation stopped at its ${this.budgetMs} ms deadline: ${count} did not run: ${steps}. ` +
                "The result covers only the steps that ran, so a run cut short is never a pass.",
        };
    }
}
