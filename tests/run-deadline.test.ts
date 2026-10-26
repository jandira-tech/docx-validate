/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { assertDeadlineMs, RunDeadline, VALIDATION_DEADLINE } from "../src/lib/run-deadline";

/** A clock that only moves when the test moves it. */
function manualClock(start = 1_000): { now: () => number; advance: (ms: number) => void } {
    let current = start;
    return {
        now: () => current,
        advance: (ms) => {
            current += ms;
        },
    };
}

describe("RunDeadline", () => {
    it("lets steps run until the budget is spent, then records each step it stops", () => {
        const clock = manualClock();
        const deadline = new RunDeadline(100, clock.now);
        expect(deadline.reached("first")).toBe(false);
        clock.advance(99);
        expect(deadline.reached("second")).toBe(false);
        expect(deadline.notRun).toEqual([]);
        clock.advance(1);
        expect(deadline.reached("third")).toBe(true);
        expect(deadline.reached("fourth")).toBe(true);
        expect(deadline.notRun).toEqual(["third", "fourth"]);
    });

    it("stops the very first step of a zero budget", () => {
        const deadline = new RunDeadline(0, manualClock().now);
        expect(deadline.expired()).toBe(true);
        expect(deadline.reached("first")).toBe(true);
        expect(deadline.notRun).toEqual(["first"]);
    });

    it("answers expired() without recording a step, and record() without asking the clock", () => {
        const clock = manualClock();
        const deadline = new RunDeadline(10, clock.now);
        expect(deadline.expired()).toBe(false);
        clock.advance(10);
        expect(deadline.expired()).toBe(true);
        expect(deadline.notRun).toEqual([]);
        deadline.record("schema validation of 2 XML parts it had not reached");
        expect(deadline.notRun).toEqual(["schema validation of 2 XML parts it had not reached"]);
    });

    it("reports the remaining budget, never below zero", () => {
        const clock = manualClock();
        const deadline = new RunDeadline(50, clock.now);
        expect(deadline.remainingMs()).toBe(50);
        clock.advance(20);
        expect(deadline.remainingMs()).toBe(30);
        clock.advance(500);
        expect(deadline.remainingMs()).toBe(0);
    });

    it("hands out a copy of the stopped steps", () => {
        const deadline = new RunDeadline(0, manualClock().now);
        deadline.reached("first");
        (deadline.notRun as string[]).push("tampered");
        expect(deadline.notRun).toEqual(["first"]);
    });
});

describe("RunDeadline.issue", () => {
    it("is undefined while every step has run", () => {
        const clock = manualClock();
        const deadline = new RunDeadline(10, clock.now);
        deadline.reached("first");
        clock.advance(50);
        // The budget is spent, but no step was stopped: the run is complete.
        expect(deadline.issue()).toBeUndefined();
    });

    it("is one error naming the budget and the steps that did not run", () => {
        const deadline = new RunDeadline(0, manualClock().now);
        deadline.reached("word-anchor-attributes");
        deadline.reached("Word verification");
        const issue = deadline.issue();
        expect(issue?.severity).toBe("error");
        expect(issue?.code).toBe(VALIDATION_DEADLINE);
        expect(issue?.path).toBeUndefined();
        expect(issue?.message).toContain("0 ms deadline");
        expect(issue?.message).toContain("2 steps did not run: word-anchor-attributes; Word verification.");
        expect(issue?.message).toMatch(/never a pass/);
    });

    it("names at most eight steps and counts the rest", () => {
        const deadline = new RunDeadline(0, manualClock().now);
        for (const step of ["s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8", "s9", "s10"]) deadline.reached(step);
        const message = deadline.issue()?.message ?? "";
        expect(message).toContain("10 steps did not run: s1; s2; s3; s4; s5; s6; s7; s8; and 2 more.");
        expect(message).not.toContain("s9");
    });

    it("says step, not steps, for one", () => {
        const deadline = new RunDeadline(0, manualClock().now);
        deadline.reached("opening the archive and every later step");
        expect(deadline.issue()?.message).toContain("1 step did not run: opening the archive and every later step.");
    });
});

describe("assertDeadlineMs", () => {
    it("accepts zero and any finite positive budget", () => {
        expect(() => assertDeadlineMs(0)).not.toThrow();
        expect(() => assertDeadlineMs(1)).not.toThrow();
        expect(() => assertDeadlineMs(2.5)).not.toThrow();
        expect(() => assertDeadlineMs(600_000)).not.toThrow();
    });

    it.each([[-1], [Number.NaN], [Number.POSITIVE_INFINITY], [Number.NEGATIVE_INFINITY], ["100"], [null]])(
        "rejects %s with a RangeError",
        (value) => {
            expect(() => assertDeadlineMs(value)).toThrow(RangeError);
            expect(() => assertDeadlineMs(value)).toThrow(/deadlineMs must be a finite number of milliseconds, 0 or more/);
        },
    );

    it("guards the constructor too", () => {
        expect(() => new RunDeadline(-5)).toThrow(RangeError);
    });
});
