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
import { classifyWordDialog, parseWordSnapshot, verifyInWord } from "../src/lib/word-verifier";

const BYTES = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xde, 0xad, 0xbe, 0xef]);

const PENDING: WordSnapshot = { processAlive: false, documentNames: [], dialogText: "", dialogButtons: [] };

type SnapshotScript = WordSnapshot | ((documentName: string) => WordSnapshot);

interface FakeSpec {
    platform?: string;
    available?: boolean;
    version?: { version: string; build?: string } | undefined;
    snapshots?: SnapshotScript[];
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
    private observeCount = 0;

    constructor(spec: FakeSpec = {}) {
        this.spec = {
            platform: spec.platform ?? "darwin",
            available: spec.available ?? true,
            version: "version" in spec ? spec.version : { version: "16.89", build: "24091315" },
        };
        this.snapshots = spec.snapshots ?? [];
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
        this.calls.push(`open:${file}${existsSync(file) ? " [exists]" : " [missing]"}`);
    }

    async observe(documentName: string): Promise<WordSnapshot> {
        this.calls.push(`observe:${documentName}`);
        const entry = this.snapshots[this.observeCount] ?? this.snapshots[this.snapshots.length - 1];
        this.observeCount += 1;
        if (entry === undefined) return PENDING;
        return typeof entry === "function" ? entry(documentName) : entry;
    }

    async close(documentName: string): Promise<void> {
        this.calls.push(`close:${documentName}`);
        if (this.closeError !== undefined) throw this.closeError;
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

    it("reports crash when the Word process disappears after opening the file", async () => {
        const transport = new FakeTransport({
            snapshots: [PENDING, (name) => openDoc(name), PENDING],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("crash");
        expect(result.satisfied).toBe(false);
    });

    it("reports timeout when the document never appears inside the budget", async () => {
        const transport = new FakeTransport({
            snapshots: [{ processAlive: true, documentNames: [], dialogText: "", dialogButtons: [] }],
        });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock, timeoutMs: 10_000 });
        expect(result.outcome).toBe("timeout");
        expect(result.satisfied).toBe(false);
        expect(result.observationMs).toBeGreaterThanOrEqual(10_000);
    });

    it("defaults to a 45-second total budget", async () => {
        const transport = new FakeTransport({ snapshots: [] });
        const clock = new FakeClock();
        const result = await verifyInWord(BYTES, { transport, clock });
        expect(result.outcome).toBe("timeout");
        expect(clock.nowValue).toBeGreaterThanOrEqual(45_000);
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
        const transport: WordTransport = {
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
        };
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
        };
        await expect(verifyInWord(BYTES, { transport, clock: new FakeClock() })).rejects.toThrow("open -g failed");
        expect(closed).toBe(true);
    });

    it("keeps the original observation failure when cleanup close also throws during unwind", async () => {
        const transport: WordTransport = {
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
        };
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
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: async (name) => {
                await new Promise((resolve) => setTimeout(resolve, 40));
                return { processAlive: true, documentNames: [name], dialogText: "", dialogButtons: [] };
            },
            close: async () => undefined,
        };
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
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("launch did not settle");
    });

    it("resolves as timeout when an observation never settles", async () => {
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: () => new Promise<WordSnapshot>(() => undefined),
            close: async () => undefined,
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5 });
        expect(result.outcome).toBe("timeout");
        expect(result.details).toContain("observation did not settle");
    });

    it("silences an observation that rejects after losing the race", async () => {
        const transport: WordTransport = {
            platform: () => "darwin",
            available: async () => true,
            version: async () => undefined,
            open: async () => undefined,
            observe: () =>
                new Promise<WordSnapshot>((_, reject) => {
                    setTimeout(() => reject(new Error("late Apple event failure")), 40);
                }),
            close: async () => undefined,
        };
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
        };
        const result = await verifyInWord(BYTES, { transport, clock: wallClock(), timeoutMs: 20, pollMs: 5, observationMs: 0 });
        expect(result.outcome).toBe("clean-open");
        expect(result.details).toContain("cleanup close exceeded the budget");
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
