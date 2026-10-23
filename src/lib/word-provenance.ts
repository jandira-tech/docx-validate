/*
 * Hash-bound Word provenance for the fixtures manifest (#274).
 *
 * A Word observation recorded in `tests/fixtures-all.manifest.json` is only
 * evidence about the exact bytes that were observed. This module binds every
 * observation to the fixture's sha256: readers must treat an observation whose
 * hash no longer matches the file as stale (`unknown`), and accuracy accounting
 * must keep `unknown`/`timeout`/stale observations out of both the clean and
 * the rejected denominators.
 */

/** One observed Microsoft Word outcome, bound to the observed bytes. */
export interface WordObservation {
    /** Observed outcome, e.g. "clean-open", "open-error", "timeout". */
    outcome: string;
    /** sha256 (hex) of the fixture bytes at observation time. */
    sha256?: string;
    /** ISO timestamp of the observation, when known. */
    observedAt?: string;
    /** Word build that produced the observation, when known. */
    wordVersion?: string;
}

/**
 * A retired observation. Appended to a manifest entry's `history` array when a
 * newer observation (or a re-hash of changed bytes) supersedes it, so prior
 * Word evidence is never discarded.
 */
export interface WordHistoryEntry extends WordObservation {
    /** sha256 (hex) of the bytes this entry was observed against. */
    sha256: string;
}

/** The manifest `word` column: legacy bare outcome string, or structured. */
export type ManifestWordField = string | WordObservation;

/** Sentinel outcome for "no usable evidence" (never observed, or stale). */
export const UNKNOWN_OUTCOME = "unknown";

/**
 * Outcomes that carry no clean/rejected signal. `unknown` — never observed, or
 * the recorded hash no longer matches the bytes; `timeout` — the probe never
 * got an answer. Both are excluded from accuracy denominators.
 */
export const UNATTRIBUTABLE_OUTCOMES: readonly string[] = [UNKNOWN_OUTCOME, "timeout"];

/** Accuracy class of an (effective) Word outcome. */
export type WordOutcomeClass = "clean" | "rejected" | "unattributable";

/**
 * Read a manifest `word` column. Accepts the structured form
 * `{ outcome, sha256, observedAt?, wordVersion? }` and the legacy bare
 * outcome string (`"clean-open"`), which yields an unbound observation —
 * there is no hash to check staleness against, so its outcome is preserved
 * as-is. Returns `null` for absent/malformed fields (no usable evidence).
 */
export function parseWordField(field: unknown): WordObservation | null {
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
    const observation: WordObservation = { outcome: candidate.outcome };
    if (typeof candidate.sha256 === "string" && candidate.sha256.length > 0) observation.sha256 = candidate.sha256;
    if (typeof candidate.observedAt === "string") observation.observedAt = candidate.observedAt;
    if (typeof candidate.wordVersion === "string") observation.wordVersion = candidate.wordVersion;
    return observation;
}

/**
 * The outcome a reader may attribute to the *current* bytes. A structured
 * observation whose recorded sha256 differs from the file's current hash is
 * stale: it must be read as `unknown`, never as clean or rejected. Legacy
 * unbound observations (no recorded hash) cannot be stale-checked and are
 * returned as-is; binding them is the writer's migration job.
 */
export function effectiveOutcome(observation: WordObservation | null, currentSha256: string | undefined): string {
    if (!observation) return UNKNOWN_OUTCOME;
    if (observation.sha256 === undefined || currentSha256 === undefined) {
        return observation.outcome;
    }
    return observation.sha256 === currentSha256 ? observation.outcome : UNKNOWN_OUTCOME;
}
