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
 * Optional verification of one OOXML package in real Microsoft Word
 * (issue #273, plan task 12).
 *
 * Deterministic under test: every side effect (process inspection, opening
 * the file, reading window/dialog state, closing the document) goes through
 * the injected `WordTransport`, and all timing goes through the injected
 * `Clock`. The production transport is `createMacWordTransport()` (osascript
 * + `open -g`), extracted from `scripts/probe-word-fixtures.ts` but
 * document-scoped: it closes only the document this run opened, never
 * quits or force-kills Word, and never touches other documents.
 */

import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { withTempDir } from "./run-cli";
import { nextSecureLongHexNumber } from "./secure-id";

/** What Word actually did with the file. `unavailable` means verification could not run at all. */
export type WordOutcome =
    | "clean-open"
    | "repair-warning"
    | "open-error"
    | "crash"
    | "password-prompt"
    | "unknown-dialog"
    | "timeout"
    | "unavailable";

/** Structured result of one verification run. `satisfied` is true iff the outcome is `clean-open`. */
export type WordVerification = {
    requested: true;
    outcome: WordOutcome;
    satisfied: boolean;
    fileSha256: string;
    wordVersion?: string;
    wordBuild?: string;
    platform: string;
    observedAt: string;
    observationMs: number;
    details: string;
};

/** One point-in-time observation of Word's state, scoped to what System Events can see. */
export type WordSnapshot = {
    processAlive: boolean;
    documentNames: string[];
    /**
     * Titles of Word's windows ("<stem>  -  Compatibility Mode"). Word can
     * answer its document list with `missing value` while the window is up,
     * so the window title is a second identity signal.
     */
    windowNames?: string[];
    /** "" when no dialog/window text is showing. */
    dialogText: string;
    dialogButtons: string[];
};

/** Side-effect boundary between the verifier and the OS/Word. */
export type WordTransport = {
    platform: () => string;
    available: () => Promise<boolean>;
    version: () => Promise<{ version: string; build?: string } | undefined>;
    open: (file: string) => Promise<void>;
    observe: (documentName: string) => Promise<WordSnapshot>;
    close: (documentName: string) => Promise<void>;
};

/** Time source. Tests inject a fake that advances only inside `sleep`. */
export type Clock = {
    now: () => number;
    sleep: (ms: number) => Promise<void>;
};

export type WordVerifyOptions = {
    transport?: WordTransport;
    clock?: Clock;
    /** Total budget including launch, observation and cleanup polling. Default 45 000 ms. */
    timeoutMs?: number;
    /** How long the document must stay open and dialog-free before the run counts as clean. Default 3 000 ms. */
    observationMs?: number;
    /** Interval between observations. Default 1 000 ms. */
    pollMs?: number;
};

export const WORD_VERIFY_TIMEOUT_MS = 45_000;
export const WORD_VERIFY_OBSERVATION_MS = 3_000;
export const WORD_VERIFY_POLL_MS = 1_000;

const systemClock: Clock = {
    now: () => Date.now(),
    sleep: (ms: number) => new Promise((resolve) => setTimeout(resolve, ms)),
};

/**
 * Map raw dialog/window text to a blocking-dialog outcome, or `null` when
 * the text is empty (no dialog). Message substrings mirror the probe's
 * AppleScript classifier.
 */
export const classifyWordDialog = (dialogText: string): WordOutcome | null => {
    if (dialogText === "") {
        return null;
    }
    if (dialogText.includes("Word found unreadable content")) {
        return "repair-warning";
    }
    if (
        dialogText.includes("Word experienced an error trying to open the file") ||
        dialogText.includes("cannot be opened because there are problems with the contents") ||
        dialogText.includes("The file is corrupt and cannot be opened")
    ) {
        return "open-error";
    }
    if (/password|encrypted/i.test(dialogText)) {
        return "password-prompt";
    }
    return "unknown-dialog";
};

const describeSnapshot = (snapshot: WordSnapshot): string => {
    const parts = [`processAlive=${snapshot.processAlive}`, `documents=${snapshot.documentNames.length}`];
    if (snapshot.documentNames.length > 0) {
        parts.push(`names=${snapshot.documentNames.join(",")}`);
    }
    if (snapshot.dialogText !== "") {
        parts.push(`dialog=${snapshot.dialogText.slice(0, 160)}`);
    }
    if (snapshot.dialogButtons.length > 0) {
        parts.push(`buttons=${snapshot.dialogButtons.join(",")}`);
    }
    return parts.join("; ");
};

/**
 * True when the snapshot shows this exact document: by name in Word's
 * document list, or by a window whose title is the name (or its stem)
 * followed by Word's "  -  " suffix. Generated names are unique, so a title
 * that merely starts with the stem does not count.
 */
const showsDocument = (snapshot: WordSnapshot, documentName: string): boolean => {
    if (snapshot.documentNames.includes(documentName)) {
        return true;
    }
    const stem = documentName.replace(/\.docx$/i, "");
    return (snapshot.windowNames ?? []).some((title) =>
        [documentName, stem].some((name) => title === name || title.startsWith(`${name}  -  `)),
    );
};

/**
 * Verify that `bytes` open cleanly in real Microsoft Word.
 *
 * The bytes are written to a uniquely named temporary copy (via
 * `withTempDir`) and only THAT document name can satisfy the run — an
 * unrelated open document does not count. Cleanup closes only that
 * document (`saving no`); Word itself is never quit or killed.
 */
export const verifyInWord = async (bytes: Uint8Array, opts: WordVerifyOptions = {}): Promise<WordVerification> => {
    const transport = opts.transport ?? createMacWordTransport();
    const clock = opts.clock ?? systemClock;
    const timeoutMs = opts.timeoutMs ?? WORD_VERIFY_TIMEOUT_MS;
    const observationMs = opts.observationMs ?? WORD_VERIFY_OBSERVATION_MS;
    const pollMs = opts.pollMs ?? WORD_VERIFY_POLL_MS;

    const platform = transport.platform();
    const fileSha256 = createHash("sha256").update(bytes).digest("hex");
    const version = await transport.version();
    const observedAt = new Date(clock.now()).toISOString();
    const finish = (outcome: WordOutcome, details: string, elapsedMs: number): WordVerification => ({
        requested: true,
        outcome,
        satisfied: outcome === "clean-open",
        fileSha256,
        ...(version?.version !== undefined ? { wordVersion: version.version } : {}),
        ...(version?.build !== undefined ? { wordBuild: version.build } : {}),
        platform,
        observedAt,
        observationMs: elapsedMs,
        details,
    });

    if (platform !== "darwin") {
        return finish("unavailable", `Word verification is macOS-only; platform is ${platform}`, 0);
    }
    if (!(await transport.available())) {
        return finish("unavailable", "Microsoft Word is not installed", 0);
    }

    return withTempDir(async (dir) => {
        const documentName = `docx-verify-${fileSha256.slice(0, 12)}-${nextSecureLongHexNumber().toString(16)}.docx`;
        const tempFile = path.join(dir, documentName);
        writeFileSync(tempFile, bytes);

        const started = clock.now();
        const deadline = started + timeoutMs;
        let sawProcess = false;
        let cleanSince: number | undefined;
        let outcome: WordOutcome = "timeout";
        let details = "no observation reached a conclusive state before the deadline";
        let elapsed = 0;
        let settled = false;

        try {
            await transport.open(tempFile);

            while (clock.now() < deadline) {
                await clock.sleep(pollMs);
                const snapshot = await transport.observe(documentName);
                details = describeSnapshot(snapshot);
                if (snapshot.processAlive) {
                    sawProcess = true;
                } else if (sawProcess) {
                    outcome = "crash";
                    details = `Word process exited after open; ${details}`;
                    break;
                }
                const dialog = classifyWordDialog(snapshot.dialogText);
                if (dialog !== null) {
                    outcome = dialog;
                    break;
                }
                if (snapshot.processAlive && showsDocument(snapshot, documentName)) {
                    cleanSince ??= clock.now();
                    if (clock.now() - cleanSince >= observationMs) {
                        outcome = "clean-open";
                        break;
                    }
                } else {
                    cleanSince = undefined;
                }
            }

            elapsed = clock.now() - started;
            settled = true;
        } finally {
            // Close runs even when launch or an observation throws. A
            // close failure after a settled run folds into details (as
            // before); during unwind it is swallowed so the original
            // error still propagates.
            try {
                await transport.close(documentName);
            } catch (err) {
                if (settled) {
                    details = `${details}; cleanup close failed: ${err instanceof Error ? err.message : String(err)}`;
                }
            }
        }
        return finish(outcome, details, elapsed);
    });
};

/**
 * Parse the `key=value` line protocol emitted by the transport's observe
 * AppleScript into a `WordSnapshot`. Exported for unit tests.
 */
export const parseWordSnapshot = (output: string): WordSnapshot => {
    const snapshot: WordSnapshot = { processAlive: false, documentNames: [], windowNames: [], dialogText: "", dialogButtons: [] };
    for (const line of output.split(/\r?\n/)) {
        if (line === "process=1") {
            snapshot.processAlive = true;
        } else if (line.startsWith("window=")) {
            snapshot.windowNames?.push(line.slice(7));
        } else if (line.startsWith("doc=")) {
            snapshot.documentNames.push(line.slice(4));
        } else if (line.startsWith("text=")) {
            const value = line.slice(5);
            snapshot.dialogText = snapshot.dialogText === "" ? value : `${snapshot.dialogText} ${value}`;
        } else if (line.startsWith("button=")) {
            snapshot.dialogButtons.push(line.slice(7));
        }
    }
    return snapshot;
};

/* v8 ignore start -- osascript/execFile boundary; exercised only by the optional WORD_VERIFY=1 integration test. */
/** osascript-based transport for macOS. See module docstring for scope guarantees. */
export const createMacWordTransport = (): WordTransport => {
    const execFileAsync = promisify(execFile);

    const osascript = async (script: string): Promise<string> => {
        const { stdout } = await execFileAsync("/usr/bin/osascript", ["-e", script], { maxBuffer: 1024 * 1024 * 8 });
        return stdout.trim();
    };

    const wordAppDirs = (): string[] => [
        "/Applications/Microsoft Word.app/Contents/Info.plist",
        path.join(homedir(), "Applications/Microsoft Word.app/Contents/Info.plist"),
    ];

    const readPlistKey = async (key: string): Promise<string | undefined> => {
        for (const plist of wordAppDirs()) {
            if (!existsSync(plist)) {
                continue;
            }
            try {
                const { stdout } = await execFileAsync("/usr/bin/defaults", ["read", plist, key]);
                const value = stdout.trim();
                if (value !== "") {
                    return value;
                }
            } catch {
                // Key absent — fall through to the next candidate location.
            }
        }
        return undefined;
    };

    const escapeAppleScriptString = (value: string): string => value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');

    return {
        platform: () => process.platform,

        available: () => Promise.resolve(wordAppDirs().some((plist) => existsSync(plist))),

        version: async () => {
            const wordVersion = await readPlistKey("CFBundleShortVersionString");
            if (wordVersion === undefined) {
                return undefined;
            }
            const build = await readPlistKey("CFBundleVersion");
            return build === undefined ? { version: wordVersion } : { version: wordVersion, build };
        },

        open: async (file: string) => {
            await execFileAsync("/usr/bin/open", ["-g", "-a", "Microsoft Word", file]);
        },

        observe: async (): Promise<WordSnapshot> => {
            const out = await osascript(`
set linesOut to ""
tell application "System Events"
  if not (exists process "Microsoft Word") then return "process=0"
  set linesOut to linesOut & "process=1" & linefeed
  tell process "Microsoft Word"
    repeat with w in windows
      try
        set linesOut to linesOut & "window=" & (name of w as text) & linefeed
      end try
      try
        repeat with t in static texts of w
          try
            set linesOut to linesOut & "text=" & (value of t as text) & linefeed
          end try
        end repeat
      end try
      try
        repeat with b in buttons of w
          try
            set linesOut to linesOut & "button=" & (name of b as text) & linefeed
          end try
        end repeat
      end try
    end repeat
  end tell
end tell
-- \`repeat with d in documents\` fails: Word's document collection does not
-- answer \`count\` (-1708). Ask for the names in one event instead.
tell application "Microsoft Word"
  try
    with timeout of 10 seconds
      set docNames to name of every document
    end timeout
    repeat with n in docNames
      set linesOut to linesOut & "doc=" & (n as text) & linefeed
    end repeat
  end try
end tell
return linesOut
`);
            return parseWordSnapshot(out);
        },

        close: async (documentName: string) => {
            const escaped = escapeAppleScriptString(documentName);
            await osascript(`
tell application "Microsoft Word"
  try
    with timeout of 10 seconds
      close document "${escaped}" saving no
    end timeout
  end try
end tell
`);
        },
    };
};
/* v8 ignore stop */
