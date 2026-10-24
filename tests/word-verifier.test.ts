/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for `src/lib/word-verifier.ts` (issue #273, plan task 12).
 *
 * Everything is deterministic: a scripted `FakeTransport` drives Word
 * observations and a `FakeClock` advances time only inside `sleep`, so no
 * test ever launches Word or waits on the real clock.
 */

import { existsSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import type { Clock, WordSnapshot, WordTransport } from "../src/lib/word-verifier";
import { classifyWordDialog, parseWordSnapshot, verifyInWord, WORD_VERIFY_CLEANUP_RESERVE_MS } from "../src/lib/word-verifier";

/** A wall-clock instant, distinct from any deadline-clock reading (#329). */
const OPEN_WALL_MS = Date.parse("2026-10-23T12:00:00.000Z");

const BYTES = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xde, 0xad, 0xbe, 0xef]);

const PENDING: WordSnapshot = { processAlive: false, documentNames: [], dialogText: "", dialogButtons: [] };

type SnapshotScript = WordSnapshot | ((documentName: string) => WordSnapshot);

interface FakeSpec {
    platform?: string;
    available?: boolean;
    version?: { version: string; build?: string } | undefined;
    snapshots?: SnapshotScript[];
    /** What an observation taken before the open sees (#315). Default: Word not running, no dialog. */
    before?: WordSnapshot;
    /**
     * The crash-report lookup after the Word process vanishes (#305):
     * "report" — a new Word crash report exists; "none" (default) — none
     * does (an external kill or quit); "hang" — the lookup never settles.
     */
    crashReport?: "report" | "none" | "hang";
}

const openDoc = (documentName: string, dialogText = "", dialogButtons: string[] = []): WordSnapshot => ({
    processAlive: true,
    documentNames: [documentName],
    dialogText,
    dialogButtons,
});

class FakeTransport implements WordTransport {
    readonly calls: string[] = [];
    private readonly spec: { platform: string; available: boolean; version: { version: string; build?: string } | undefined };
    private readonly snapshots: SnapshotScript[];
    private readonly before: WordSnapshot;
    private readonly crashReport: "report" | "none" | "hang";
    private observeCount = 0;
    private opened = false;

    constructor(spec: FakeSpec = {}) {
        this.spec = {
            platform: spec.platform ?? "darwin",
            available: spec.available ?? true,
            version: "version" in spec ? spec.version : { version: "16.89", build: "24091315" },
        };
        this.snapshots = spec.snapshots ?? [];
        this.before = spec.before ?? PENDING;
        this.crashReport = spec.crashReport ?? "none";
    }

    platform(): string {
        return this.spec.platform;
    }

    async available(): Promise<boolean> {
        this.calls.push("available");
        return this.spec.available;
    }

    async version(): Promise<{ version: string; build?: string } | undefined> {
        this.calls.push("version");
        return this.spec.version;
    }

    async open(file: string): Promise<void> {
        this.opened = true;
        this.calls.push(`open:${file}${existsSync(file) ? " [exists]" : " [missing]"}`);
    }

    async observe(documentName: string): Promise<WordSnapshot> {
        this.calls.push(`observe:${documentName}`);
        // Observations before the open do not consume the scripted sequence.
        if (!this.opened) return this.before;
        const entry = this.snapshots[this.observeCount] ?? this.snapshots[this.snapshots.length - 1];
        this.observeCount += 1;
        if (entry === undefined) return PENDING;
        return typeof entry === "function" ? entry(documentName) : entry;
    }

    async close(documentName: string): Promise<void> {
        this.calls.push(`close:${documentName}`);
        if (this.closeError !== undefined) throw this.closeError;
    }

    async crashReportedSince(sinceMs: number): Promise<boolean> {
        this.calls.push(`crashReportedSince:${sinceMs}`);
        if (this.crashReport === "hang") {
            return new Promise<boolean>(() => undefined);
        }
        return this.crashReport === "report";
    }

    closeError?: unknown;
}

class FakeClock {
    nowValue = 0;
    readonly sleeps: number[] = [];
    private timers: { at: number; resolve: () => void }[] = [];

    now(): number {
        return this.nowValue;
    }

    async sleep(ms: number): Promise<void> {
        this.sleeps.push(ms);
        this.nowValue += ms;
        this.fireDueTimers();
    }

    advance(ms: number): void {
        this.nowValue += ms;
        this.fireDueTimers();
    }

    timer(ms: number): Promise<void> {
        return new Promise((resolve) => {
            this.timers.push({ at: this.nowValue + ms, resolve });
        });
    }

    private fireDueTimers(): void {
        const due = this.timers.filter((t) => t.at <= this.nowValue);
        if (due.length === 0) return;
        this.timers = this.timers.filter((t) => t.at > this.nowValue);
        for (const timer of due) timer.resolve();
    }
}

/**
 * Scope a transport's `observe` to the run after the open: the observation
 * taken before the open (#315) sees Word not running and no dialog.
 */
const observingAfterOpen = (transport: WordTransport): WordTransport => {
    let opened = false;
    return {
        ...transport,
        open: (file) => {
            opened = true;
            return transport.open(file);
        },
        observe: (name) => (opened ? transport.observe(name) : Promise.resolve(PENDING)),
    };
};

/** Wall-clock-backed fake: `now` advances with real time, `timer` uses real setTimeout. */
const wallClock = (): Clock => {
    const t0 = Date.now();
    return {
        now: () => Date.now() - t0,
        sleep: async (ms) => {
            await new Promise((resolve) => setTimeout(resolve, ms));
        },
        timer: (ms) =>
            new Promise((resolve) => {
                setTimeout(resolve, ms);
            }),
    };
};

describe("verifyInWord", () => {
    it("reports a clean open once the temporary document stays open and dialog-free for the observation period", async () => {
        const transport = new FakeTransport({
            snapshots: [PENDING, (name) => openDoc(name), (name) => openDoc(name), (name) => openDoc(name), (name) => openDoc(name)],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("clean-open");
        expect(result.satisfied).toBe(true);
        expect(result.observationMs).toBeGreaterThanOrEqual(3000);
    });

    it("opens a unique temporary copy and only counts that document name — an unrelated open document does not satisfy", async () => {
        const transport = new FakeTransport({
            snapshots: [
                () => ({ processAlive: true, documentNames: ["Quarterly Report.docx"], dialogText: "", dialogButtons: [] }),
                (name) => ({ processAlive: true, documentNames: ["Quarterly Report.docx", name], dialogText: "", dialogButtons: [] }),
                (name) => ({ processAlive: true, documentNames: ["Quarterly Report.docx", name], dialogText: "", dialogButtons: [] }),
                (name) => ({ processAlive: true, documentNames: ["Quarterly Report.docx", name], dialogText: "", dialogButtons: [] }),
            ],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("clean-open");

        const openCall = transport.calls.find((c) => c.startsWith("open:"));
        expect(openCall).toBeDefined();
        const openedPath = openCall!.slice(5).replace(" [exists]", "").replace(" [missing]", "");
        const documentName = path.basename(openedPath);
        expect(documentName).toMatch(/^docx-verify-[0-9a-f]+-[0-9a-f]+\.docx$/);
        expect(documentName).not.toBe("Quarterly Report.docx");
        // The temp copy must exist on disk when Word is asked to open it.
        expect(openCall).toContain("[exists]");
        // Every observation and the cleanup close are scoped to our document name.
        for (const call of transport.calls.filter((c) => c.startsWith("observe:"))) {
            expect(call).toBe(`observe:${documentName}`);
        }
        expect(transport.calls.filter((c) => c.startsWith("close:"))).toEqual([`close:${documentName}`]);
    });

    it("identifies our document by its window title when Word's document list comes back empty", async () => {
        // Real Word sometimes answers `name of every document` with `missing value`
        // while the document window is up; the title is "<stem>  -  <mode>".
        const windowOnly = (name: string): WordSnapshot => ({
            processAlive: true,
            documentNames: [],
            windowNames: [`${name.replace(/\.docx$/, "")}  -  Compatibility Mode`],
            dialogText: "",
            dialogButtons: [],
        });
        const transport = new FakeTransport({ snapshots: [windowOnly, windowOnly, windowOnly, windowOnly, windowOnly] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock() });
        expect(result.outcome).toBe("clean-open");
    });

    it("does not accept a window whose title only resembles our document", async () => {
        const lookalike = (name: string): WordSnapshot => ({
            processAlive: true,
            documentNames: [],
            windowNames: [`${name.replace(/\.docx$/, "")}0  -  Compatibility Mode`, "Quarterly Report  -  Compatibility Mode"],
            dialogText: "",
            dialogButtons: [],
        });
        const transport = new FakeTransport({ snapshots: [lookalike] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock(), timeoutMs: 5000 });
        expect(result.outcome).toBe("timeout");
    });

    it("does not count as open while only an unrelated document is visible (times out instead of clean-open)", async () => {
        const transport = new FakeTransport({
            snapshots: [() => ({ processAlive: true, documentNames: ["Someone Elses Doc.docx"], dialogText: "", dialogButtons: [] })],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5000 });
        expect(result.outcome).toBe("timeout");
        expect(result.satisfied).toBe(false);
    });

    it("reports repair-warning when the unreadable-content dialog appears during the observation period", async () => {
        const transport = new FakeTransport({
            snapshots: [
                (name) => openDoc(name),
                (name) => openDoc(name),
                (name) =>
                    openDoc(name, "Word found unreadable content in document.docx. Do you want to recover the contents?", ["Yes", "No"]),
            ],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("repair-warning");
        expect(result.satisfied).toBe(false);
        expect(result.details).toContain("unreadable content");
    });

    it.each([
        "Word experienced an error trying to open the file.",
        "cannot be opened because there are problems with the contents",
        "The file is corrupt and cannot be opened",
    ])("reports open-error for the blocking open dialog: %s", async (dialogText) => {
        const transport = new FakeTransport({
            snapshots: [(name) => openDoc(name, dialogText, ["OK"])],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("open-error");
        expect(result.satisfied).toBe(false);
    });

    it("reports password-prompt when Word asks for a password", async () => {
        const transport = new FakeTransport({
            snapshots: [{ processAlive: true, documentNames: [], dialogText: "Password", dialogButtons: ["OK", "Cancel"] }],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("password-prompt");
        expect(result.satisfied).toBe(false);
    });

    it("reports unknown-dialog for dialog text that matches no known signature", async () => {
        const transport = new FakeTransport({
            snapshots: [(name) => openDoc(name, "Microsoft Word wants to access your contacts", ["Allow", "Don't Allow"])],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("unknown-dialog");
        expect(result.satisfied).toBe(false);
        expect(result.details).toContain("wants to access your contacts");
    });

    it("reports timeout when the document never appears inside the budget", async () => {
        const transport = new FakeTransport({
            snapshots: [{ processAlive: true, documentNames: [], dialogText: "", dialogButtons: [] }],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 10_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.satisfied).toBe(false);
        // Observation runs until the slice kept for the cleanup close (#308).
        expect(result.observationMs).toBeGreaterThanOrEqual(10_000 - WORD_VERIFY_CLEANUP_RESERVE_MS);
        expect(result.observationMs).toBeLessThanOrEqual(10_000);
    });

    it("defaults to a 45-second total budget", async () => {
        const transport = new FakeTransport({ snapshots: [] });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("timeout");
        expect(clock.nowValue).toBeGreaterThanOrEqual(45_000 - WORD_VERIFY_CLEANUP_RESERVE_MS);
        expect(clock.nowValue).toBeLessThanOrEqual(45_000);
    });

    it("closes only the document this run opened — never quits Word, never touches other documents", async () => {
        const transport = new FakeTransport({
            snapshots: [
                (name) => ({
                    processAlive: true,
                    documentNames: ["Keep Me Open.docx", name],
                    dialogText: "Word experienced an error trying to open the file.",
                    dialogButtons: ["OK"],
                }),
            ],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("open-error");

        const openedPath = transport.calls
            .find((c) => c.startsWith("open:"))!
            .slice(5)
            .replace(" [exists]", "");
        const documentName = path.basename(openedPath);
        const closes = transport.calls.filter((c) => c.startsWith("close:"));
        expect(closes).toEqual([`close:${documentName}`]);
        // Nothing in the call log references the other document or a quit/kill action.
        expect(transport.calls.join("\n")).not.toContain("Keep Me Open.docx".replace(".docx", "") + "|quit");
        expect(transport.calls.every((c) => !c.includes("quit") && !c.includes("pkill") && !c.startsWith("close:Keep"))).toBe(true);
    });

    it("reports unavailable and never opens Word on a non-macOS platform", async () => {
        const transport = new FakeTransport({ platform: "linux" });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("unavailable");
        expect(result.satisfied).toBe(false);
        expect(result.platform).toBe("linux");
        expect(result.details).toContain("linux");
        expect(transport.calls.some((c) => c.startsWith("open:"))).toBe(false);
    });

    it("reports unavailable and never opens Word when Word is not installed", async () => {
        const transport = new FakeTransport({ available: false });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("unavailable");
        expect(result.satisfied).toBe(false);
        expect(transport.calls.some((c) => c.startsWith("open:"))).toBe(false);
    });

    it("records the file sha256, Word version/build, platform, request flag and observation timestamp", async () => {
        const transport = new FakeTransport({
            version: { version: "16.89", build: "24091315" },
            snapshots: [(name) => openDoc(name)],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.requested).toBe(true);
        expect(result.fileSha256).toMatch(/^[0-9a-f]{64}$/);
        expect(result.wordVersion).toBe("16.89");
        expect(result.wordBuild).toBe("24091315");
        expect(result.platform).toBe("darwin");
        expect(new Date(result.observedAt).toString()).not.toBe("Invalid Date");
        expect(result.observationMs).toBeGreaterThan(0);
        expect(typeof result.details).toBe("string");
    });

    it("requires the document to stay open and dialog-free continuously — a drop resets the streak", async () => {
        const transport = new FakeTransport({
            snapshots: [
                (name) => openDoc(name),
                (name) => openDoc(name),
                { processAlive: true, documentNames: [], dialogText: "", dialogButtons: [] },
                (name) => openDoc(name),
                (name) => openDoc(name),
                (name) => openDoc(name),
                (name) => openDoc(name),
            ],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock, observationMs: 3000, pollMs: 1000, timeoutMs: 30_000 });
        // The two pre-drop clean polls do not count toward the 3 s streak that
        // restarts when the document reappears; the run takes ≥ 6 s of clean presence.
        expect(result.outcome).toBe("clean-open");
        expect(result.observationMs).toBeGreaterThanOrEqual(6000);
    });

    it("keeps the outcome when cleanup close throws, and records the failure in details", async () => {
        const transport = new FakeTransport({
            snapshots: [(name) => openDoc(name, "The file is corrupt and cannot be opened", ["OK"])],
        });
        transport.closeError = new Error("document already gone");
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("open-error");
        expect(result.details).toContain("cleanup close failed: document already gone");
    });

    it("stringifies non-Error cleanup failures", async () => {
        const transport = new FakeTransport({
            snapshots: [(name) => openDoc(name, "The file is corrupt and cannot be opened", ["OK"])],
        });
        transport.closeError = "plain failure";
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.details).toContain("cleanup close failed: plain failure");
    });

    it("closes after observation failure", async () => {
        let closed = false;
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async () => {
                throw new Error("Apple event timed out");
            },
            close: async () => {
                closed = true;
            },
            crashReportedSince: async () => false,
        });
        await expect(verifyInWord(new Uint8Array([1]), { transport, clock: new FakeClock() })).rejects.toThrow("Apple event");
        expect(closed).toBe(true);
    });

    it("closes after launch failure and propagates the launch error", async () => {
        let closed = false;
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => {
                throw new Error("open -g failed");
            },
            observe: async () => PENDING,
            close: async () => {
                closed = true;
            },
            crashReportedSince: async () => false,
        };
        await expect(verifyInWord(BYTES, { transport, clock: new FakeClock() })).rejects.toThrow("open -g failed");
        expect(closed).toBe(true);
    });

    it("keeps the original observation failure when cleanup close also throws during unwind", async () => {
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async () => {
                throw new Error("Apple event timed out");
            },
            close: async () => {
                throw new Error("close also failed");
            },
            crashReportedSince: async () => false,
        });
        await expect(verifyInWord(BYTES, { transport, clock: new FakeClock() })).rejects.toThrow("Apple event");
    });

    it("counts launch against the total budget", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => {
                clock.advance(10_000);
            },
            observe: async (name) => ({
                processAlive: true,
                documentNames: [name],
                dialogText: "",
                dialogButtons: [],
            }),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(new Uint8Array([1]), {
            transport,
            clock,
            timeoutMs: 5_000,
            observationMs: 1_000,
        });
        expect(result.outcome).toBe("timeout");
    });

    it("ends as timeout when an observation settles only after the deadline", async () => {
        // The observe call resolves cleanly 40 ms in, but the 20 ms budget is
        // already spent by then: the run must end as timeout instead of
        // acting on the stale observation.
        const t0 = Date.now();
        const clock: Clock = {
            now: () => Date.now() - t0,
            sleep: async (ms) => {
                await new Promise((resolve) => setTimeout(resolve, ms));
            },
            timer: (ms) =>
                new Promise((resolve) => {
                    setTimeout(resolve, ms);
                }),
        };
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async (name) => {
                await new Promise((resolve) => setTimeout(resolve, 40));
                return { processAlive: true, documentNames: [name], dialogText: "", dialogButtons: [] };
            },
            close: async () => undefined,
            crashReportedSince: async () => false,
        });
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 20, pollMs: 5, observationMs: 0 });
        expect(result.outcome).toBe("timeout");
    });

    it("resolves as timeout when launch never settles", async () => {
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: () => new Promise<void>(() => undefined),
            observe: async (name) => ({
                processAlive: true,
                documentNames: [name],
                dialogText: "",
                dialogButtons: [],
            }),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("launch did not settle");
    });

    it("resolves as timeout when an observation never settles", async () => {
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: () => new Promise<WordSnapshot>(() => undefined),
            close: async () => undefined,
            crashReportedSince: async () => false,
        });
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("observation did not settle");
    });

    it("silences an observation that rejects after losing the race", async () => {
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: () =>
                new Promise<WordSnapshot>((_, reject) => {
                    setTimeout(() => reject(new Error("late Apple event failure")), 40);
                }),
            close: async () => undefined,
            crashReportedSince: async () => false,
        });
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        // Stay alive past the late rejection: an unhandled rejection fails the file.
        await new Promise((resolve) => setTimeout(resolve, 60));
    });

    it("abandons a cleanup close that outlives the budget and keeps the outcome", async () => {
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async (name) => ({
                processAlive: true,
                documentNames: [name],
                dialogText: "",
                dialogButtons: [],
            }),
            close: () => new Promise<void>(() => undefined),
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5, observationMs: 0 });
        expect(result.outcome).toBe("clean-open");
        expect(result.details).toContain("cleanup close exceeded the budget");
    });

    // #300 closes what a run opened; #301 bounds the run. The last slice of
    // the budget is kept for the cleanup close, so a launch or observation
    // that uses up its share still leaves the document closed.
    it.each([
        ["launch never settles", { open: () => new Promise<void>(() => undefined) }],
        ["an observation never settles", { observe: () => new Promise<WordSnapshot>(() => undefined) }],
    ] as const)("still attempts the cleanup close when %s", async (_case, hang) => {
        const closed: string[] = [];
        const transport: WordTransport = observingAfterOpen({
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async (name) => ({ processAlive: true, documentNames: [name], dialogText: "", dialogButtons: [] }),
            close: async (name) => {
                closed.push(name);
            },
            crashReportedSince: async () => false,
            ...hang,
        });
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 40, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(closed).toHaveLength(1);
        expect(closed[0]).toMatch(/^docx-verify-.*\.docx$/);
    });

    it("still attempts the cleanup close when launch overruns the whole budget", async () => {
        const clock = new FakeClock();
        const closed: string[] = [];
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => {
                clock.advance(10_000);
            },
            observe: async (name) => ({ processAlive: true, documentNames: [name], dialogText: "", dialogButtons: [] }),
            close: async (name) => {
                closed.push(name);
            },
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5_000, observationMs: 1_000 });
        expect(result.outcome).toBe("timeout");
        expect(closed).toHaveLength(1);
    });

    it("awaits transport calls without a race when the clock provides no timer", async () => {
        const transport = new FakeTransport({ snapshots: [(name) => openDoc(name)] });
        const clock: Clock = { now: () => 0, sleep: async () => undefined };
        const result = await verifyInWord(BYTES, { transport, clock, observationMs: 0 });
        expect(result.outcome).toBe("clean-open");
    });

    it("uses the real clock and macOS transport defaults only through injection points", async () => {
        // No clock injected: the system clock drives the (tiny) timeout budget.
        const transport = new FakeTransport({ snapshots: [] });
        const result = await verifyInWord(BYTES, { transport, timeoutMs: 2, pollMs: 1 });
        expect(result.outcome).toBe("timeout");
        expect(result.satisfied).toBe(false);
    });

    it("omits wordVersion/wordBuild when the transport cannot report a version", async () => {
        const transport = new FakeTransport({ version: undefined, snapshots: [(name) => openDoc(name)] });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("clean-open");
        expect(result.wordVersion).toBeUndefined();
        expect(result.wordBuild).toBeUndefined();
    });
});

describe("verifyInWord: a dialog already showing before the open (#315)", () => {
    const STALE_DIALOG: WordSnapshot = {
        processAlive: true,
        documentNames: [],
        dialogText: "Word experienced an error trying to open the file.",
        dialogButtons: ["OK"],
    };

    it("reports unavailable and neither opens nor closes anything", async () => {
        const transport = new FakeTransport({ before: STALE_DIALOG, snapshots: [(name) => openDoc(name)] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock() });
        expect(result.outcome).toBe("unavailable");
        expect(result.satisfied).toBe(false);
        expect(result.details).toContain("already showing before the open");
        expect(result.details).toContain("Word experienced an error trying to open the file");
        expect(transport.calls.filter((c) => c.startsWith("observe:"))).toHaveLength(1);
        expect(transport.calls.some((c) => c.startsWith("open:") || c.startsWith("close:"))).toBe(false);
    });

    it("observes once before the open, then opens as usual when no dialog is showing", async () => {
        const transport = new FakeTransport({ snapshots: [(name) => openDoc(name)] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock() });
        expect(result.outcome).toBe("clean-open");
        const firstObserve = transport.calls.findIndex((c) => c.startsWith("observe:"));
        const open = transport.calls.findIndex((c) => c.startsWith("open:"));
        expect(firstObserve).toBeGreaterThanOrEqual(0);
        expect(firstObserve).toBeLessThan(open);
    });

    it("ends as timeout, without opening or closing, when the observation before the open never settles", async () => {
        const calls: string[] = [];
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => {
                calls.push("open");
            },
            observe: () => new Promise<WordSnapshot>(() => undefined),
            close: async () => {
                calls.push("close");
            },
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("before the open");
        expect(calls).toEqual([]);
    });

    it("propagates a failed observation before the open without opening or closing", async () => {
        const calls: string[] = [];
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => {
                calls.push("open");
            },
            observe: async () => {
                throw new Error("Apple event timed out");
            },
            close: async () => {
                calls.push("close");
            },
            crashReportedSince: async () => false,
        };
        await expect(verifyInWord(BYTES, { transport, clock: new FakeClock() })).rejects.toThrow("Apple event");
        expect(calls).toEqual([]);
    });
});

describe("verifyInWord: a vanished Word process needs crash evidence (#305)", () => {
    // The process is seen alive with our document, then vanishes.
    const VANISHED: SnapshotScript[] = [PENDING, (name) => openDoc(name), PENDING];

    it("reports crash when the process disappears and a new Word crash report was written after the open", async () => {
        const transport = new FakeTransport({ crashReport: "report", snapshots: VANISHED });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock, wallNow: () => OPEN_WALL_MS });
        expect(result.outcome).toBe("crash");
        expect(result.satisfied).toBe(false);
        expect(result.details).toContain("crash report");
        // Only a report written after this run's open counts. Report files
        // carry wall-clock mtimes, so the lookup receives the open instant
        // in wall time, not in the deadline clock's time (#329).
        expect(transport.calls).toContain(`crashReportedSince:${OPEN_WALL_MS}`);
    });

    it("reports unavailable, not crash, when Word vanishes without a crash report (an external kill or quit)", async () => {
        const transport = new FakeTransport({ snapshots: VANISHED });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("unavailable");
        expect(result.satisfied).toBe(false);
        expect(result.details).toContain("Word exited without a crash report");
    });

    it("ends as timeout when the crash-report lookup does not settle within the budget", async () => {
        const transport = new FakeTransport({ crashReport: "hang", snapshots: VANISHED });
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 40, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("the crash-report lookup did not settle");
    });
});

describe("verifyInWord: the version and availability calls race the budget (#320)", () => {
    it("keeps a version that settles within the budget", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => {
                clock.advance(2_000);
                return { version: "16.89", build: "24091315" };
            },
            available: async () => true,
            open: async () => undefined,
            observe: async (name) => openDoc(name),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 10_000 });
        expect(result.outcome).toBe("clean-open");
        expect(result.wordVersion).toBe("16.89");
        expect(result.wordBuild).toBe("24091315");
    });

    it("omits the version when its call settles only after the budget is spent", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => {
                // The budget expires while the call is in flight; it answers only after.
                await Promise.resolve().then(() => clock.advance(5_000));
                return { version: "16.89" };
            },
            available: async () => true,
            open: async () => undefined,
            observe: async (name) => openDoc(name),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.wordVersion).toBeUndefined();
        expect(result.observationMs).toBeLessThanOrEqual(5_000);
    });

    it("ends as timeout when the availability call settles only after the budget is spent", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => undefined,
            available: async () => {
                // The budget expires while the call is in flight; it answers only after.
                await Promise.resolve().then(() => clock.advance(5_000));
                return true;
            },
            open: async () => undefined,
            observe: async (name) => openDoc(name),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("the availability check did not settle");
        expect(result.observationMs).toBeLessThanOrEqual(5_000);
    });

    it("settles within the budget when the availability call never settles", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => undefined,
            available: () => {
                const never = new Promise<boolean>(() => undefined);
                // The budget expires while the call is still in flight.
                queueMicrotask(() => clock.advance(5_000));
                return never;
            },
            open: async () => undefined,
            observe: async (name) => openDoc(name),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("the availability check did not settle");
        expect(clock.nowValue).toBeLessThanOrEqual(5_000);
    });

    it("settles within the budget when the version call never settles", async () => {
        const clock = new FakeClock();
        const transport: WordTransport = {
            platform: () => "darwin",
            version: () => {
                const never = new Promise<{ version: string }>(() => undefined);
                queueMicrotask(() => clock.advance(5_000));
                return never;
            },
            available: async () => true,
            open: async () => undefined,
            observe: async (name) => openDoc(name),
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 5_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.wordVersion).toBeUndefined();
        expect(clock.nowValue).toBeLessThanOrEqual(5_000);
    });
});

describe("classifyWordDialog", () => {
    it("returns null when no dialog text is present", () => {
        expect(classifyWordDialog("")).toBeNull();
    });

    it("checks the unreadable-content signature before the password signature", () => {
        expect(classifyWordDialog("Word found unreadable content in Password-protected.docx")).toBe("repair-warning");
    });
});

describe("parseWordSnapshot", () => {
    it("parses the observe line protocol into a snapshot", () => {
        const out = ["process=1", "window=Document1", "text=Some prompt text", "button=OK", "doc=Report.docx", "doc=Extra.docx"].join("\n");
        expect(parseWordSnapshot(out)).toEqual({
            processAlive: true,
            documentNames: ["Report.docx", "Extra.docx"],
            windowNames: ["Document1"],
            dialogText: "Some prompt text",
            dialogButtons: ["OK"],
        });
    });

    it("reports a dead process when Word is not running", () => {
        expect(parseWordSnapshot("process=0")).toEqual({
            processAlive: false,
            documentNames: [],
            windowNames: [],
            dialogText: "",
            dialogButtons: [],
        });
    });

    it("joins multiple static-text fragments into one dialog text", () => {
        const out = "process=1\ntext=First line\ntext=Second line";
        expect(parseWordSnapshot(out).dialogText).toBe("First line Second line");
    });
});

describe("verifyInWord: a late-settling open is closed after it settles (#328)", () => {
    it("waits for an open that lost its race before closing the document", async () => {
        const clock = new FakeClock();
        const calls: string[] = [];
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => undefined,
            available: async () => true,
            open: () => {
                calls.push("open");
                // The work budget expires while the open is in flight; Word
                // finishes opening the document only afterwards.
                queueMicrotask(() => clock.advance(8_500));
                return new Promise<void>((resolve) => {
                    setTimeout(() => {
                        calls.push("open settled");
                        resolve();
                    }, 20);
                });
            },
            observe: async () => PENDING,
            close: async () => {
                calls.push("close");
            },
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 10_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("launch did not settle");
        // Closing first would leave the late-opened document open in Word.
        expect(calls).toEqual(["open", "open settled", "close"]);
    });

    it("still closes, and says so, when the late open has not settled by the deadline", async () => {
        const transport: WordTransport = {
            platform: () => "darwin",
            version: async () => undefined,
            available: async () => true,
            open: () => new Promise<void>(() => undefined),
            observe: async () => PENDING,
            close: async () => undefined,
            crashReportedSince: async () => false,
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 40, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("the late open had not settled when the cleanup close ran");
    });
});

describe("verifyInWord: observedAt is wall time, not deadline time (#329)", () => {
    it("stamps observedAt from the injected wall clock while the deadline clock reads 0", async () => {
        const transport = new FakeTransport({ snapshots: [(name) => openDoc(name)] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock(), wallNow: () => OPEN_WALL_MS });
        expect(result.observedAt).toBe("2026-10-23T12:00:00.000Z");
    });

    it("defaults observedAt to the system wall clock, so a deadline clock at 0 never yields 1970", async () => {
        const before = Date.now();
        const transport = new FakeTransport({ snapshots: [(name) => openDoc(name)] });
        const result = await verifyInWord(BYTES, { transport, clock: new FakeClock() });
        expect(Date.parse(result.observedAt)).toBeGreaterThanOrEqual(before);
        expect(Date.parse(result.observedAt)).toBeLessThanOrEqual(Date.now());
    });
});
