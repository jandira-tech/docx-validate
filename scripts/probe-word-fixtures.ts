/*
 * Probe utility: run the current validator over every .docx fixture, then
 * open the same files in installed Microsoft Word and record whether Word
 * opens them cleanly or shows a blocking corruption/recovery dialog.
 *
 * macOS only. Requires Accessibility automation permission for the terminal
 * running this script.
 *
 * Run with:
 *   bunx tsx scripts/probe-word-fixtures.ts \
 *     --profile word-valid \
 *     --out /tmp/docx-word-probe-results.jsonl \
 *     --force-close-existing-word
 */

import { execFile } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync, appendFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { Command } from "commander";

import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_FIXTURES = path.resolve(HERE, "..", "tests/fixtures");

type WordOutcome = "clean-open" | "unreadable-content-warning" | "open-error" | "password-required" | "unknown-dialog" | "timeout";

interface CliOptions {
    root: string;
    out: string;
    profile: Profile;
    only: "all" | "validator-valid" | "validator-invalid";
    match?: string;
    limit?: string;
    timeoutMs: string;
    pollMs: string;
    validTimeoutRetries: string;
    resume: boolean;
    dryRun: boolean;
    forceCloseExistingWord: boolean;
    quitAtEnd: boolean;
}

interface ValidatorOutcome {
    valid: boolean;
    errorCodes: string[];
    threw?: string;
}

interface WordProbe {
    outcome: WordOutcome;
    clean: boolean;
    details: string;
    durationMs: number;
}

interface ProbeRecord {
    relativePath: string;
    file: string;
    validator: ValidatorOutcome & { profile: Profile };
    word: WordProbe | null;
    aligned: boolean | null;
    mismatch: "validator-false-positive" | "validator-false-negative" | null;
}

interface Summary {
    totalConsidered: number;
    probed: number;
    skippedByResume: number;
    validatorValid: number;
    validatorInvalid: number;
    wordClean: number;
    wordNotClean: number;
    aligned: number;
    falsePositive: number;
    falseNegative: number;
    outcomes: Record<string, number>;
}

function walkDocx(dir: string, out: string[]): void {
    for (const entry of readdirSync(dir)) {
        if (entry.startsWith("~$")) continue;
        const file = path.join(dir, entry);
        const stat = statSync(file);
        if (stat.isDirectory()) {
            walkDocx(file, out);
        } else if (entry.toLowerCase().endsWith(".docx")) {
            out.push(file);
        }
    }
}

async function runValidator(file: string, profile: Profile): Promise<ValidatorOutcome> {
    try {
        const result = await validate(file, { profile });
        const errorCodes = Array.from(
            new Set(
                result.issues
                    .filter((i) => i.severity === "error")
                    .map((i) => i.code)
                    .filter((c): c is string => Boolean(c)),
            ),
        ).sort();
        return { valid: result.valid, errorCodes };
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        return { valid: false, errorCodes: [], threw: msg.slice(0, 240) };
    }
}

function buildCommand(): Command {
    return new Command()
        .name("probe-word-fixtures")
        .description("Compare docx-validate results against installed Microsoft Word open behavior")
        .option("--root <dir>", "Fixture root to scan recursively for .docx files", DEFAULT_FIXTURES)
        .option("--out <file>", "JSONL log path", path.resolve(HERE, "..", "tests/word-probe-results.jsonl"))
        .option("--profile <profile>", "Validator profile: lenient, strict, or word-valid", "word-valid")
        .option("--only <group>", "Probe all, validator-valid, or validator-invalid files", "all")
        .option("--match <text>", "Only probe files whose relative path contains this text")
        .option("--limit <n>", "Maximum number of files to probe")
        .option("--timeout-ms <ms>", "Per-file Word open timeout", "45000")
        .option("--poll-ms <ms>", "Word polling interval", "1000")
        .option("--valid-timeout-retries <n>", "Retry timeout outcomes for validator-valid files", "1")
        .option("--resume", "Skip files already present in the JSONL log", false)
        .option("--dry-run", "Run the validator and write predicted groups without opening Word", false)
        .option("--force-close-existing-word", "Quit an existing Word session before probing", false)
        .option("--no-quit-at-end", "Leave Microsoft Word running after the probe");
}

function parseOptions(argv: readonly string[]): CliOptions {
    const command = buildCommand();
    command.parse(argv as string[], { from: "user" });
    const opts = command.opts<CliOptions>();
    if (opts.profile !== "lenient" && opts.profile !== "strict" && opts.profile !== "word-valid") {
        throw new Error(`Invalid --profile: ${String(opts.profile)}. Must be lenient, strict, or word-valid.`);
    }
    if (opts.only !== "all" && opts.only !== "validator-valid" && opts.only !== "validator-invalid") {
        throw new Error(`Invalid --only: ${String(opts.only)}. Must be all, validator-valid, or validator-invalid.`);
    }
    return opts;
}

function readCompleted(outFile: string): Set<string> {
    if (!existsSync(outFile)) return new Set();
    const done = new Set<string>();
    for (const line of readFileSync(outFile, "utf-8").split(/\r?\n/)) {
        if (!line.trim()) continue;
        try {
            const parsed = JSON.parse(line) as Partial<ProbeRecord>;
            if (parsed.relativePath) done.add(parsed.relativePath);
        } catch {
            // Ignore partial/truncated JSONL lines so an interrupted run can resume.
        }
    }
    return done;
}

async function osascript(script: string): Promise<string> {
    const { stdout } = await execFileAsync("/usr/bin/osascript", ["-e", script], {
        maxBuffer: 1024 * 1024 * 8,
    });
    return stdout.trim();
}

async function wordWindowCount(): Promise<number> {
    const out = await osascript(`
tell application "System Events"
  if not (exists process "Microsoft Word") then return "0"
  tell process "Microsoft Word"
    return (count of windows) as text
  end tell
end tell
`);
    return Number.parseInt(out, 10) || 0;
}

async function quitWord(): Promise<void> {
    await osascript(`
tell application "Microsoft Word"
  try
    quit saving no
  end try
end tell
`);
}

async function forceQuitWord(): Promise<void> {
    try {
        await execFileAsync("/usr/bin/pkill", ["-9", "-x", "Microsoft Word"]);
    } catch {
        // No Word process is also fine.
    }
}

async function cleanupWord(): Promise<void> {
    await osascript(`
tell application "System Events"
  if exists process "Microsoft Word" then
    tell process "Microsoft Word"
      repeat 5 times
        set clickedButton to false
        repeat with w in windows
          repeat with buttonName in {"Don't Save", "No", "OK", "Cancel"}
            try
              if exists button buttonName of w then
                click button buttonName of w
                set clickedButton to true
                exit repeat
              end if
            end try
          end repeat
          if clickedButton then exit repeat
        end repeat
        if not clickedButton then exit repeat
        delay 0.2
      end repeat
    end tell
  end if
end tell
tell application "Microsoft Word"
  try
    repeat 20 times
      if (count of documents) = 0 then exit repeat
      close active document saving no
    end repeat
  end try
end tell
`);
}

async function inspectWord(): Promise<{ status: WordOutcome | "pending"; details: string }> {
    const out = await osascript(`
tell application "System Events"
  if not (exists process "Microsoft Word") then return "pending|NO_PROCESS"
  tell process "Microsoft Word"
    set winCount to count of windows
    if winCount = 0 then
      set docDetails to "window_count=0" & linefeed
      tell application "Microsoft Word"
        try
          set docDetails to docDetails & "word_documents=" & ((count of documents) as text) & linefeed
          if (count of documents) > 0 then
            set docDetails to docDetails & "active_document=" & (name of active document as text) & linefeed
            return "clean-open|" & docDetails
          end if
        end try
      end tell
      return "pending|" & docDetails
    end if

    set details to "window_count=" & winCount & linefeed
    set allText to ""
    set hasNamedDoc to false
    repeat with w in windows
      set windowName to ""
      try
        set windowName to name of w as text
      end try
      set details to details & "name=" & windowName & linefeed
      if windowName is not "" and windowName is not "Microsoft Word" then set hasNamedDoc to true
      try
        set details to details & "buttons=" & ((name of buttons of w) as text) & linefeed
      end try
      try
        set textValue to ((value of static texts of w) as text)
        set allText to allText & textValue
        set details to details & "static=" & textValue & linefeed
      end try
    end repeat

    if allText contains "Word found unreadable content" then
      repeat with w in windows
        try
          if exists button "No" of w then click button "No" of w
        end try
      end repeat
      return "unreadable-content-warning|" & details
    end if

    if allText contains "Word experienced an error trying to open the file" then
      repeat with w in windows
        try
          if exists button "OK" of w then click button "OK" of w
        end try
      end repeat
      return "open-error|" & details
    end if

    if allText contains "cannot be opened because there are problems with the contents" then
      repeat with w in windows
        try
          if exists button "OK" of w then click button "OK" of w
        end try
      end repeat
      return "open-error|" & details
    end if

    if allText contains "The file is corrupt and cannot be opened" then
      repeat with w in windows
        try
          if exists button "OK" of w then click button "OK" of w
        end try
      end repeat
      return "open-error|" & details
    end if

    if allText contains "Password" or allText contains "password" or allText contains "encrypted" then
      repeat with w in windows
        try
          if exists button "Cancel" of w then click button "Cancel" of w
        end try
      end repeat
      return "password-required|" & details
    end if

    if hasNamedDoc and allText is "" then
      return "clean-open|" & details
    end if

    if allText is not "" then return "unknown-dialog|" & details
    return "pending|" & details
  end tell
end tell
`);
    const [status, ...rest] = out.split("|");
    return {
        status: status as WordOutcome | "pending",
        details: rest.join("|"),
    };
}

function sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

async function probeWord(file: string, timeoutMs: number, pollMs: number): Promise<WordProbe> {
    const started = Date.now();
    let lastDetails = "";
    await execFileAsync("/usr/bin/open", ["-a", "Microsoft Word", file]);
    while (Date.now() - started < timeoutMs) {
        await sleep(pollMs);
        const observed = await inspectWord();
        lastDetails = observed.details;
        if (observed.status !== "pending") {
            await cleanupWord();
            if (observed.status !== "clean-open") {
                await forceQuitWord();
                await sleep(1000);
            }
            return {
                outcome: observed.status,
                clean: observed.status === "clean-open",
                details: observed.details,
                durationMs: Date.now() - started,
            };
        }
    }
    await cleanupWord();
    await forceQuitWord();
    await sleep(1000);
    return {
        outcome: "timeout",
        clean: false,
        details: lastDetails,
        durationMs: Date.now() - started,
    };
}

function includeByGroup(valid: boolean, only: CliOptions["only"]): boolean {
    if (only === "all") return true;
    if (only === "validator-valid") return valid;
    return !valid;
}

function classify(record: ProbeRecord): ProbeRecord {
    if (!record.word) return record;
    const validatorClean = record.validator.valid;
    const wordClean = record.word.clean;
    const aligned = validatorClean === wordClean;
    let mismatch: ProbeRecord["mismatch"] = null;
    if (!aligned) {
        mismatch = validatorClean ? "validator-false-negative" : "validator-false-positive";
    }
    return { ...record, aligned, mismatch };
}

function updateSummary(summary: Summary, record: ProbeRecord): void {
    summary.totalConsidered += 1;
    if (record.validator.valid) summary.validatorValid += 1;
    else summary.validatorInvalid += 1;
    if (!record.word) return;
    summary.probed += 1;
    summary.outcomes[record.word.outcome] = (summary.outcomes[record.word.outcome] ?? 0) + 1;
    if (record.word.clean) summary.wordClean += 1;
    else summary.wordNotClean += 1;
    if (record.aligned) summary.aligned += 1;
    if (record.mismatch === "validator-false-positive") summary.falsePositive += 1;
    if (record.mismatch === "validator-false-negative") summary.falseNegative += 1;
}

function emptySummary(): Summary {
    return {
        totalConsidered: 0,
        probed: 0,
        skippedByResume: 0,
        validatorValid: 0,
        validatorInvalid: 0,
        wordClean: 0,
        wordNotClean: 0,
        aligned: 0,
        falsePositive: 0,
        falseNegative: 0,
        outcomes: {},
    };
}

async function run(argv: readonly string[]): Promise<number> {
    const opts = parseOptions(argv);
    const root = path.resolve(opts.root);
    const out = path.resolve(opts.out);
    const timeoutMs = Number.parseInt(opts.timeoutMs, 10);
    const pollMs = Number.parseInt(opts.pollMs, 10);
    const validTimeoutRetries = Number.parseInt(opts.validTimeoutRetries, 10);
    const limit = opts.limit ? Number.parseInt(opts.limit, 10) : undefined;
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error("--timeout-ms must be a positive integer");
    if (!Number.isFinite(pollMs) || pollMs <= 0) throw new Error("--poll-ms must be a positive integer");
    if (!Number.isFinite(validTimeoutRetries) || validTimeoutRetries < 0) {
        throw new Error("--valid-timeout-retries must be a non-negative integer");
    }
    if (limit !== undefined && (!Number.isFinite(limit) || limit <= 0)) throw new Error("--limit must be a positive integer");

    const files: string[] = [];
    walkDocx(root, files);
    files.sort();
    const filtered = opts.match ? files.filter((f) => path.relative(root, f).includes(opts.match ?? "")) : files;
    const completed = opts.resume ? readCompleted(out) : new Set<string>();

    mkdirSync(path.dirname(out), { recursive: true });
    if (!opts.resume) writeFileSync(out, "");

    if (!opts.dryRun) {
        const windowCount = await wordWindowCount();
        if (windowCount > 0) {
            if (!opts.forceCloseExistingWord) {
                throw new Error(
                    `Microsoft Word already has ${windowCount} window(s). Re-run with --force-close-existing-word if they are safe to close.`,
                );
            }
            await quitWord();
            await sleep(1500);
        }
    }

    const summary = emptySummary();
    let probedThisRun = 0;
    let seen = 0;
    for (const file of filtered) {
        const relativePath = path.relative(root, file);
        if (completed.has(relativePath)) {
            summary.skippedByResume += 1;
            continue;
        }

        const validator = { ...(await runValidator(file, opts.profile)), profile: opts.profile };
        if (!includeByGroup(validator.valid, opts.only)) continue;
        if (limit !== undefined && probedThisRun >= limit) break;

        seen += 1;
        process.stderr.write(`\n[${seen}] ${validator.valid ? "validator-valid" : "validator-invalid"} ${relativePath}\n`);

        const baseRecord: ProbeRecord = {
            relativePath,
            file,
            validator,
            word: null,
            aligned: null,
            mismatch: null,
        };

        let word: WordProbe | null = null;
        if (!opts.dryRun) {
            word = await probeWord(file, timeoutMs, pollMs);
            for (let retry = 0; retry < validTimeoutRetries && validator.valid && word.outcome === "timeout"; retry += 1) {
                process.stderr.write(`    timeout on validator-valid file; retrying (${retry + 1}/${validTimeoutRetries})\n`);
                word = await probeWord(file, timeoutMs, pollMs);
            }
        }

        const record = opts.dryRun ? baseRecord : classify({ ...baseRecord, word });

        appendFileSync(out, `${JSON.stringify(record)}\n`);
        updateSummary(summary, record);
        probedThisRun += 1;

        const wordLabel = record.word ? `${record.word.outcome}${record.word.clean ? "" : " (not clean)"}` : "dry-run";
        const alignLabel = record.aligned === null ? "" : record.aligned ? " aligned" : ` ${record.mismatch}`;
        process.stderr.write(`    Word: ${wordLabel}${alignLabel}\n`);
    }

    if (!opts.dryRun && opts.quitAtEnd) await quitWord();
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
    return summary.falsePositive === 0 && summary.falseNegative === 0 ? 0 : 2;
}

run(process.argv.slice(2)).then(
    (code) => {
        process.exitCode = code;
    },
    (err: unknown) => {
        const message = err instanceof Error ? err.message : String(err);
        process.stderr.write(`${message}\n`);
        process.exitCode = 1;
    },
);
