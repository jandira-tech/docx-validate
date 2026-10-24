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
 * Hash-bound Word provenance for the fixtures manifest (#274).
 *
 * A Word observation recorded in `tests/fixtures-all.manifest.json` is only
 * evidence about the exact bytes that were observed. This module binds every
 * observation to the fixture's sha256: readers must treat an observation whose
 * hash no longer matches the file as stale (`unknown`), and accuracy accounting
 * must keep `unknown`/`timeout`/stale observations out of both the clean and
 * the rejected denominators.
 */

import type { WordOutcome } from "./word-verifier";

/** One observed Microsoft Word outcome, bound to the observed bytes. */
export type WordObservation = {
    /** Observed outcome, e.g. "clean-open", "open-error", "timeout". */
    readonly outcome: string;
    /** sha256 (hex) of the fixture bytes at observation time. */
    readonly sha256?: string;
    /** ISO timestamp of the observation, when known. */
    readonly observedAt?: string;
    /** Word build that produced the observation, when known. */
    readonly wordVersion?: string;
};

/**
 * A retired observation. Appended to a manifest entry's `history` array when a
 * newer observation (or a re-hash of changed bytes) supersedes it, so prior
 * Word evidence is never discarded.
 */
export type WordHistoryEntry = WordObservation & {
    /** sha256 (hex) of the bytes this entry was observed against. */
    readonly sha256: string;
};

/** The manifest `word` column: legacy bare outcome string, or structured. */
export type ManifestWordField = string | WordObservation;

/** Sentinel outcome for "no usable evidence" (never observed, or stale). */
export const UNKNOWN_OUTCOME = "unknown";

/**
 * Outcomes that carry no clean/rejected signal. `unknown` — never observed, or
 * the recorded hash no longer matches the bytes; `timeout` — the probe never
 * got an answer; `crash` — Word died before delivering a verdict on the
 * content. All are excluded from accuracy denominators.
 */
export const UNATTRIBUTABLE_OUTCOMES: readonly string[] = [UNKNOWN_OUTCOME, "timeout", "crash"];

/**
 * Outcomes that are Word's verdict against the content — the only outcomes
 * counted as rejections (#316). Anything else, known or not, is never a
 * rejection without evidence of Word's verdict.
 */
export const REJECTED_OUTCOMES: readonly string[] = ["open-error", "unreadable-content-warning", "password-required", "unknown-dialog"];

/**
 * Translate a verifier outcome (`WordOutcome`) into the manifest's outcome
 * vocabulary (#316). The two spell the recovery and password dialogs
 * differently (`repair-warning`/`unreadable-content-warning`,
 * `password-prompt`/`password-required`); the shared spellings pass through,
 * and outcomes that carry no Word verdict map to their honest no-signal
 * equivalents — `crash` stays `crash` (Word died before a verdict) and
 * `unavailable` becomes `unknown`. Use this wherever a verifier outcome
 * becomes manifest evidence.
 */
export const VERIFIER_OUTCOME_TO_MANIFEST: Readonly<Record<WordOutcome, string>> = {
    "clean-open": "clean-open",
    "repair-warning": "unreadable-content-warning",
    "open-error": "open-error",
    crash: "crash",
    "password-prompt": "password-required",
    "unknown-dialog": "unknown-dialog",
    timeout: "timeout",
    unavailable: UNKNOWN_OUTCOME,
};

export const toManifestOutcome = (outcome: WordOutcome): string => VERIFIER_OUTCOME_TO_MANIFEST[outcome];

/** Accuracy class of an (effective) Word outcome. */
export type WordOutcomeClass = "clean" | "rejected" | "unattributable";

/**
 * Read a manifest `word` column. Accepts the structured form
 * `{ outcome, sha256, observedAt?, wordVersion? }` and the legacy bare
 * outcome string (`"clean-open"`), which yields an unbound observation —
 * there is no hash to check staleness against, so its outcome is preserved
 * as-is. Returns `null` for absent/malformed fields (no usable evidence).
 */
export const parseWordField = (field: unknown): WordObservation | null => {
    if (typeof field === "string" && field.length > 0) {
        return { outcome: field };
    }
    if (typeof field !== "object" || field === null) {
        return null;
    }
    const candidate = field as { outcome?: unknown; sha256?: unknown; observedAt?: unknown; wordVersion?: unknown };
    if (typeof candidate.outcome !== "string" || candidate.outcome.length === 0) {
        return null;
    }
    const observation: { outcome: string; sha256?: string; observedAt?: string; wordVersion?: string } = { outcome: candidate.outcome };
    if (typeof candidate.sha256 === "string" && candidate.sha256.length > 0) {
        observation.sha256 = candidate.sha256;
    }
    if (typeof candidate.observedAt === "string") {
        observation.observedAt = candidate.observedAt;
    }
    if (typeof candidate.wordVersion === "string") {
        observation.wordVersion = candidate.wordVersion;
    }
    return observation;
};

/**
 * The outcome a reader may attribute to the *current* bytes. A structured
 * observation whose recorded sha256 differs from the file's current hash is
 * stale: it must be read as `unknown`, never as clean or rejected. Legacy
 * unbound observations (no recorded hash) cannot be stale-checked and are
 * returned as-is; binding them is the writer's migration job.
 */
export const effectiveOutcome = (observation: WordObservation | null, currentSha256: string | undefined): string => {
    if (!observation) {
        return UNKNOWN_OUTCOME;
    }
    if (observation.sha256 === undefined || currentSha256 === undefined) {
        return observation.outcome;
    }
    return observation.sha256 === currentSha256 ? observation.outcome : UNKNOWN_OUTCOME;
};

/**
 * Classify an (effective) outcome for accuracy accounting. `clean-open` is
 * the only clean outcome, and a rejection needs one of Word's verdict
 * outcomes ({@link REJECTED_OUTCOMES}): an open error, a recovery offer, a
 * password prompt or an unknown blocking dialog. Everything else — `unknown`,
 * `timeout`, `crash`, stale or not-yet-mapped spellings — is `unattributable`
 * and stays out of both denominators: never a rejection without evidence of
 * Word's verdict (#316).
 */
export const classifyWordOutcome = (outcome: string): WordOutcomeClass => {
    if (outcome === "clean-open") {
        return "clean";
    }
    return REJECTED_OUTCOMES.includes(outcome) ? "rejected" : "unattributable";
};

/** Clean/rejected/unattributable tally over manifest word fields. */
export type WordAccuracyCounts = {
    readonly clean: number;
    readonly rejected: number;
    readonly unattributable: number;
};
/** Inputs to a manifest-entry word-column merge (see {@link mergeWordObservation}). */
export type WordMergeInput = {
    /** Parsed prior `word` column, when the entry existed before. */
    readonly prior: WordObservation | null;
    /** Prior retired observations (`history`), preserved as-is. */
    readonly priorHistory?: readonly WordHistoryEntry[];
    /** Fresh probe outcome, when the probe JSONL has one for this fixture. */
    readonly freshOutcome?: string;
    /** sha256 (hex) of the fixture bytes today. */
    readonly currentSha256: string;
    /** Optional timestamp/Word build stamped onto a fresh observation. */
    readonly observedAt?: string;
    readonly wordVersion?: string;
};

/** Merged `word` column plus its (possibly extended) `history` array. */
export type WordMergeResult = {
    readonly word: WordObservation;
    readonly history: readonly WordHistoryEntry[];
};

export const mergeWordObservation = (input: WordMergeInput): WordMergeResult => {
    const history: WordHistoryEntry[] = [...(input.priorHistory ?? [])];
    const { prior } = input;
    const priorIsCurrent = prior !== null && (prior.sha256 === undefined || prior.sha256 === input.currentSha256);

    if (prior !== null && prior.sha256 !== undefined && prior.sha256 !== input.currentSha256) {
        history.push({ outcome: prior.outcome, sha256: prior.sha256, observedAt: prior.observedAt, wordVersion: prior.wordVersion });
    }

    if (input.freshOutcome !== undefined) {
        const fresh: WordHistoryEntry = {
            outcome: input.freshOutcome,
            sha256: input.currentSha256,
            ...(input.observedAt !== undefined ? { observedAt: input.observedAt } : {}),
            ...(input.wordVersion !== undefined ? { wordVersion: input.wordVersion } : {}),
        };
        if (prior !== null && priorIsCurrent) {
            const current: WordHistoryEntry = { ...prior, sha256: input.currentSha256 };
            // An unattributable re-probe (timeout/unknown) carries no signal
            // about these bytes, so it never displaces attributable evidence.
            if (classifyWordOutcome(fresh.outcome) === "unattributable" && classifyWordOutcome(current.outcome) !== "unattributable") {
                return { word: current, history: [...history, fresh] };
            }
            // Same bytes, newer observation: the prior one is retired, not dropped.
            history.push(current);
        }
        return { word: fresh, history };
    }

    if (prior === null) {
        return { word: { outcome: UNKNOWN_OUTCOME, sha256: input.currentSha256 }, history };
    }

    if (priorIsCurrent) {
        return { word: { ...prior, sha256: input.currentSha256 }, history };
    }

    return { word: { outcome: UNKNOWN_OUTCOME, sha256: input.currentSha256 }, history };
};

export const wordAccuracy = (fields: readonly { readonly field: unknown; readonly sha256: string | undefined }[]): WordAccuracyCounts => {
    const counts: { clean: number; rejected: number; unattributable: number } = { clean: 0, rejected: 0, unattributable: 0 };
    for (const entry of fields) {
        counts[classifyWordOutcome(effectiveOutcome(parseWordField(entry.field), entry.sha256))] += 1;
    }
    return counts;
};
