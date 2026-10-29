/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `deadlineMs` bounds the whole `validate()` run (#630).
 *
 * The bound is cooperative: every step asks the run's deadline before it
 * starts, and a step already running finishes. A step the deadline stopped is
 * named in one error-severity `validation-deadline` finding, so a run cut
 * short is never a pass, whatever the steps that did run found.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import { RunDeadline, VALIDATION_DEADLINE } from "../src/lib/run-deadline";
import type { Profile, ValidationIssue } from "../src/lib/types";
import type { WordVerification } from "../src/lib/word-verifier";
import type { XsdValidator } from "../src/lib/xsd-validator";
import { jubarte } from "../src/index";
import { unpack } from "../src/scripts/office/unpack";
import { runValidateFromArgv, validate, type ValidateRunResult } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { PPTXSchemaValidator } from "../src/scripts/office/validators/pptx";
import { Validate } from "../src/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");

const deadlineIssues = (issues: readonly ValidationIssue[]): ValidationIssue[] =>
    issues.filter((issue) => issue.code === VALIDATION_DEADLINE);
const codes = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "");

/** An XSD engine that answers every part valid after `delayMs` of real time. */
const slowEngine = (delayMs: number): XsdValidator => ({
    validate: () => new Promise((resolve) => setTimeout(() => resolve([]), delayMs)),
});

/** Write `parts` as an unpacked package under `dir`. */
function writeUnpacked(dir: string, parts: Readonly<Record<string, string>>): string {
    for (const [name, content] of Object.entries(parts)) {
        mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
        writeFileSync(path.join(dir, name), content);
    }
    return dir;
}

/** Run the CLI with stdout/stderr captured, restoring both afterwards. */
async function capture(argv: readonly string[]): Promise<{ code: number; stdout: string; stderr: string }> {
    const out: string[] = [];
    const err: string[] = [];
    const stdoutWrite = process.stdout.write.bind(process.stdout);
    const stderrWrite = process.stderr.write.bind(process.stderr);
    process.stdout.write = ((chunk: string | Uint8Array): boolean => {
        out.push(typeof chunk === "string" ? chunk : Buffer.from(chunk).toString("utf-8"));
        return true;
    }) as typeof process.stdout.write;
    process.stderr.write = ((chunk: string | Uint8Array): boolean => {
        err.push(typeof chunk === "string" ? chunk : Buffer.from(chunk).toString("utf-8"));
        return true;
    }) as typeof process.stderr.write;
    try {
        const code = await runValidateFromArgv(argv);
        return { code, stdout: out.join(""), stderr: err.join("") };
    } finally {
        process.stdout.write = stdoutWrite;
        process.stderr.write = stderrWrite;
    }
}

/** A Word verifier that records how it was called. */
function recordingVerifier(): {
    verify: (bytes: Uint8Array, options?: { timeoutMs?: number }) => Promise<WordVerification>;
    calls: unknown[][];
} {
    const calls: unknown[][] = [];
    const verify = async (...args: [Uint8Array, { timeoutMs?: number }?]): Promise<WordVerification> => {
        calls.push(args);
        return {
            requested: true,
            outcome: "clean-open",
            satisfied: true,
            fileSha256: "0".repeat(64),
            platform: "darwin",
            observedAt: "2026-10-25T00:00:00.000Z",
            observationMs: 3000,
            details: "injected",
        };
    };
    return { verify, calls };
}

/** Everything of a result except the deadline's own disclosure. */
const verdict = (result: ValidateRunResult): unknown => {
    const { deadline: _deadline, ...metadata } = result.metadata ?? {};
    return { valid: result.valid, issues: result.issues, metadata, xsdWalk: result.xsdWalk };
};

describe("validate(): deadlineMs bounds the whole run (#630)", { timeout: 300_000 }, () => {
    it("changes nothing when no deadline is given", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            const result = await validate(file);
            expect(deadlineIssues(result.issues)).toEqual([]);
            expect(result.metadata?.deadline).toBeUndefined();
        });
    });

    it("gives the same verdict inside a generous budget as without one, under every profile", async () => {
        const fixtures = [
            "broken/comments.unmatched-comment-marker.docx",
            "broken/package.zip-eocd-comment-overrun.docx",
            "broken/package.word-textbox-deleted-text-shape.docx",
            "broken/package.word-duplicate-settings.docx",
            "broken/package.word-footnote-reference-missing.docx",
            "working/package.word-move-from-text.docx",
        ];
        const profiles: Profile[] = ["word-valid", "strict", "lenient"];
        for (const fixture of fixtures) {
            for (const profile of profiles) {
                const file = path.join(FIXTURES, fixture);
                const unbounded = await validate(file, { profile });
                const bounded = await validate(file, { profile, deadlineMs: 600_000 });
                expect(verdict(bounded), `${fixture} under ${profile}`).toEqual(verdict(unbounded));
                expect(bounded.metadata?.deadline).toEqual({ budgetMs: 600_000, notRun: [] });
            }
        }
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.pptx");
            writeFileSync(file, await buildPptx());
            for (const profile of profiles) {
                const unbounded = await validate(file, { profile });
                const bounded = await validate(file, { profile, deadlineMs: 600_000 });
                expect(verdict(bounded), `control.pptx under ${profile}`).toEqual(verdict(unbounded));
            }
        });
    });

    it("never passes a clean package whose budget ran out before anything ran", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            expect((await validate(file)).valid).toBe(true);

            const result = await validate(file, { deadlineMs: 0 });
            expect(result.valid).toBe(false);
            expect(result.issues).toEqual(deadlineIssues(result.issues));
            expect(deadlineIssues(result.issues)).toHaveLength(1);
            expect(deadlineIssues(result.issues)[0]).toMatchObject({ severity: "error" });
            expect(deadlineIssues(result.issues)[0]?.message).toContain("opening the archive and every later step");
            expect(result.metadata?.executedChecks).toEqual([]);
            expect(result.metadata?.schemaCompleted).toBe(false);
            expect(result.metadata?.deadline?.budgetMs).toBe(0);
            expect(result.metadata?.deadline?.notRun).toContain("opening the archive and every later step");
        });
    });

    it("bounds a directory, a partial run and a PPTX the same way", async () => {
        await withTempDir(async (dir) => {
            const unpacked = writeUnpacked(path.join(dir, "unpacked.docx"), MINIMAL_PARTS);
            const fromDirectory = await validate(unpacked, { deadlineMs: 0 });
            expect(fromDirectory.valid).toBe(false);
            expect(deadlineIssues(fromDirectory.issues)).toHaveLength(1);
            expect(fromDirectory.metadata?.deadline?.notRun).toContain("reading the directory and every later step");

            const docx = path.join(dir, "control.docx");
            writeFileSync(docx, await buildDocx());
            const quick = await validate(docx, { quick: true, deadlineMs: 0 });
            expect(quick.valid).toBe(false);
            expect(deadlineIssues(quick.issues)).toHaveLength(1);
            expect(quick.metadata?.executedChecks).toEqual([]);

            const pptx = path.join(dir, "control.pptx");
            writeFileSync(pptx, await buildPptx());
            const slides = await validate(pptx, { deadlineMs: 0 });
            expect(slides.valid).toBe(false);
            expect(deadlineIssues(slides.issues)).toHaveLength(1);
        });
    });

    it("stops a slow schema walk between parts and reports one deadline finding", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            const started = Date.now();
            const result = await validate(file, { profile: "strict", xsdValidator: slowEngine(1_500), deadlineMs: 1_000 });
            const elapsed = Date.now() - started;
            // Five schema-mapped parts take 7.5 s unbounded. The walk stops
            // after the part that spent the budget, or never starts.
            expect(elapsed).toBeLessThan(6_000);
            expect(result.xsdWalk.parts.length).toBeLessThanOrEqual(1);
            expect(result.metadata?.schemaCompleted).toBe(false);
            expect(result.valid).toBe(false);
            expect(deadlineIssues(result.issues)).toHaveLength(1);
            expect(deadlineIssues(result.issues)[0]?.message).toMatch(/schema validation|validateAgainstXsd/);
        });
    });

    it("hands Word verification what is left of the budget, and skips it once the budget is spent", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());

            const unbounded = recordingVerifier();
            await validate(file, { verifyWord: true, wordVerifier: unbounded.verify });
            // Without a deadline the verifier is called exactly as before.
            expect(unbounded.calls).toHaveLength(1);
            expect(unbounded.calls[0]).toHaveLength(1);

            const bounded = recordingVerifier();
            const generous = await validate(file, { verifyWord: true, wordVerifier: bounded.verify, deadlineMs: 600_000 });
            expect(generous.valid).toBe(true);
            expect(bounded.calls).toHaveLength(1);
            const timeoutMs = (bounded.calls[0]?.[1] as { timeoutMs?: number } | undefined)?.timeoutMs;
            // Never past the deadline, and never longer than the verifier's own default budget.
            expect(timeoutMs).toBeGreaterThan(0);
            expect(timeoutMs).toBeLessThanOrEqual(45_000);

            const spent = recordingVerifier();
            const stopped = await validate(file, { verifyWord: true, wordVerifier: spent.verify, deadlineMs: 0 });
            expect(spent.calls).toEqual([]);
            expect(stopped.valid).toBe(false);
            expect(stopped.metadata?.wordVerification).toMatchObject({ requested: true, outcome: "unavailable", satisfied: false });
            expect(stopped.metadata?.wordVerification?.details).toMatch(/deadline/);
            expect(stopped.metadata?.deadline?.notRun).toContain("Word verification");
            expect(deadlineIssues(stopped.issues)).toHaveLength(1);
        });
    });

    it("refuses a deadline with --auto-repair before touching the file", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            const bytes = await buildDocx();
            writeFileSync(file, bytes);
            await expect(validate(file, { autoRepair: true, deadlineMs: 60_000 })).rejects.toThrow(ModeConflictError);
            await expect(validate(file, { autoRepair: true, deadlineMs: 60_000 })).rejects.toThrow(/--auto-repair.*--deadline-ms/);
            expect(readFileSync(file).equals(bytes)).toBe(true);
        });
    });

    it.each([[-1], [Number.NaN], [Number.POSITIVE_INFINITY]])("rejects deadlineMs %s before reading anything", async (deadlineMs) => {
        await expect(validate(path.join(FIXTURES, "does-not-exist.docx"), { deadlineMs })).rejects.toThrow(RangeError);
    });
});

describe("the validators honour a shared RunDeadline (#630)", { timeout: 120_000 }, () => {
    it("stops the schema walk between parts and counts the parts it never reached", async () => {
        await withTempDir(async (dir) => {
            const unpacked = writeUnpacked(dir, MINIMAL_PARTS);
            // Each engine call spends 100 ms of a clock that moves only there.
            let now = 0;
            const engine: XsdValidator = {
                validate: async () => {
                    now += 100;
                    return [];
                },
            };
            const deadline = new RunDeadline(250, () => now);
            const validator = new DOCXSchemaValidator({ unpackedDir: unpacked, profile: "strict", xsdValidator: engine, deadline });
            await validator.validateAgainstXsd();
            expect(validator.xsdWalk.parts).toHaveLength(3);
            expect(validator.xsdWalk.notExamined).toBe(2);
            expect(deadline.notRun).toEqual(["schema validation of 2 XML parts it had not reached"]);
        });
    });

    it("skips the checks after the budget, so their findings are absent, and reports that", async () => {
        await withTempDir(async (dir) => {
            const fixture = path.join(FIXTURES, "broken/package.word-footnote-reference-missing.docx");
            const unpacked = path.join(dir, "unpacked");
            await unpack(fixture, unpacked, { mergeRuns: false, simplifyRedlines: false });

            const complete = await new DOCXSchemaValidator({ unpackedDir: unpacked, profile: "word-valid" }).validate();
            expect(codes(complete.issues)).toContain("word-footnote-reference-missing");

            // Every look at this clock moves it one tick: the budget of three
            // lets two steps start.
            let tick = 0;
            const deadline = new RunDeadline(3, () => tick++);
            const validator = new DOCXSchemaValidator({ unpackedDir: unpacked, profile: "word-valid", deadline });
            const result = await validator.validate();
            expect(codes(result.issues)).not.toContain("word-footnote-reference-missing");
            expect(result.valid).toBe(false);
            expect(deadlineIssues(result.issues)).toHaveLength(1);
            expect(deadline.notRun).toContain("DOCX Word rules");
            expect(deadline.notRun).not.toContain("DOCX validateXml");
            expect(validator.executedChecks).toContain("xml-syntax");
        });
    });
});

describe("the PPTX validator honours a shared RunDeadline (#630)", { timeout: 120_000 }, () => {
    it("stops its rules once the budget is spent and reports that", async () => {
        await withTempDir(async (dir) => {
            const unpacked = writeUnpacked(dir, MINIMAL_PPTX_PARTS);
            expect((await new PPTXSchemaValidator({ unpackedDir: unpacked, profile: "lenient" }).validate()).valid).toBe(true);

            let tick = 0;
            const deadline = new RunDeadline(3, () => tick++);
            const result = await new PPTXSchemaValidator({ unpackedDir: unpacked, profile: "lenient", deadline }).validate();
            expect(result.valid).toBe(false);
            expect(deadlineIssues(result.issues)).toHaveLength(1);
            expect(deadline.notRun).toContain("PPTX validateAgainstXsd");
            expect(deadline.notRun).not.toContain("PPTX validateXml");
        });
    });
});

describe("the bytes APIs and the CLI forward the deadline (#630)", { timeout: 120_000 }, () => {
    it("Validate and jubarte.validate bound their run", async () => {
        const bytes = await buildDocx();
        const viaClass = await new Validate({ deadlineMs: 0 }).run(bytes);
        expect(viaClass.valid).toBe(false);
        expect(deadlineIssues(viaClass.issues)).toHaveLength(1);
        expect(viaClass.metadata?.deadline?.budgetMs).toBe(0);

        const viaNamespace = await jubarte.validate(bytes, { deadlineMs: 0 });
        expect(deadlineIssues(viaNamespace.issues)).toHaveLength(1);

        const generous = await new Validate({ deadlineMs: 600_000 }).run(bytes);
        expect(generous.valid).toBe(true);
        expect(generous.metadata?.deadline?.notRun).toEqual([]);

        await expect(new Validate({ deadlineMs: -1 }).run(bytes)).rejects.toThrow(RangeError);
    });

    it("--deadline-ms bounds a CLI run, rejects a bad value and refuses --auto-repair", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());

            const stopped = await capture([file, "--deadline-ms", "0"]);
            expect(stopped.code).toBe(1);
            expect(stopped.stderr).toContain("0 ms deadline");

            const json = await capture([file, "--deadline-ms", "0", "--json"]);
            expect(json.code).toBe(1);
            const document = JSON.parse(json.stdout) as { metadata: { deadline?: { budgetMs: number } } };
            expect(document.metadata.deadline?.budgetMs).toBe(0);

            const generous = await capture([file, "--deadline-ms", "600000"]);
            expect(generous.code).toBe(0);
            expect(generous.stdout).toContain("All validations PASSED!");

            for (const bad of ["abc", "-1", "Infinity", ""]) {
                const rejected = await capture([file, "--deadline-ms", bad]);
                expect(rejected.code, bad).toBe(1);
                expect(rejected.stderr, bad).toContain("Invalid --deadline-ms");
            }

            const repair = await capture([file, "--deadline-ms", "5000", "--auto-repair"]);
            expect(repair.code).toBe(1);
            expect(repair.stderr).toMatch(/--auto-repair.*--deadline-ms/);
        });
    });
});
