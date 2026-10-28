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
 * One place that turns "what the caller asked for" into "what the run does".
 *
 * The library API (`validate(path, options)`) and the CLI
 * (`src/scripts/office/validate.ts`) both call {@link resolveMode}; neither
 * resolves profiles, defaults or selectors on its own, so the two can never
 * drift.
 *
 * Contract (plan § 1):
 *
 * | selection                     | profile      | scope                    |
 * | ----------------------------- | ------------ | ------------------------ |
 * | nothing                       | per format   | every applicable check   |
 * | `profile: "lenient"`          | lenient      | every applicable check   |
 * | `profile: "word-valid"`       | word-valid   | every applicable check   |
 * | `strict` / `profile:"strict"` | strict       | every applicable check   |
 * | `quick`                       | per format   | the quick preset groups  |
 * | `checks`                      | per format   | the named checks + prereq|
 *
 * `strict` is the complete mode that additionally requires schema validation to
 * complete, so it cannot be combined with partial check selection.
 */

import { DEFAULT_PROFILE, type PackageFormat, type Profile, type ValidationIssue } from "./types";

/**
 * The groups `--quick` runs: package integrity, XML syntax, content types,
 * relationships, table structure and drawing structure.
 *
 * A group with no registered member yet contributes nothing; the run reports
 * which checks it covered and which it omitted, so the gap is visible rather
 * than implied.
 */
export const QUICK_CHECK_GROUPS: readonly string[] = ["package", "xml", "content-types", "relationships", "tables", "drawings"];

/** What the caller asked for. Every field is optional. */
export type ModeRequest = {
    /** Explicit `--profile`. */
    profile?: Profile;
    /** `--strict`: shorthand for `profile: "strict"`. */
    strict?: boolean;
    /** `--quick`: run {@link QUICK_CHECK_GROUPS} and disclose the omissions. */
    quick?: boolean;
    /** `--checks`: run these ids or group names, plus their prerequisites. */
    checks?: readonly string[];
    /** `--auto-repair`: repair before validating. Only a full run repairs. */
    autoRepair?: boolean;
};

/**
 * Two selectors that contradict each other (`--quick` with `--checks`,
 * `--strict` or `--auto-repair` with partial selection, …). A configuration error: the caller
 * asked for something that has no meaning, so nothing is validated.
 */
export class ModeConflictError extends Error {
    public constructor(message: string) {
        super(message);
        this.name = "ModeConflictError";
    }
}

/** What the run will actually do. */
export type ResolvedMode = {
    /** The profile the schema validators run under. */
    profile: Profile;
    /**
     * Check ids or group names to run, or `undefined` for "every applicable
     * check". A non-`undefined` selection makes the run *partial*: the
     * applicable ids it leaves out must be reported as omissions.
     */
    select: readonly string[] | undefined;
};

/**
 * The profile a format defaults to.
 *
 * DOCX/DOCM default to `"word-valid"`: the whole point of the library is to
 * tell you whether Microsoft Word opens the file, and the brief's plan flips
 * the default deliberately once the Word rules run as first-class checks.
 * Every other suffix keeps the historical `"lenient"` default, which matches
 * real-world Office output, so PPTX behaviour is unchanged.
 *
 * Callers that want the pre-flip behaviour pass `profile: "lenient"`
 * explicitly; nothing silently falls back to it.
 */
export const defaultProfileForFormat = (format: PackageFormat): Profile => (format === "docx" ? "word-valid" : DEFAULT_PROFILE);

/**
 * The severity a package-name finding carries under a profile and a format:
 * an error under strict, and under word-valid for DOCX/DOCM — Word 16.115
 * offers recovery for every malformed name probed (#351, #408, #414) — and a
 * warning under lenient, and under word-valid for the unobserved PPTX/XLSX.
 * Word-valid then demotes what Word opens clean through `isWordBlockingIssue`,
 * as it does an Override PartName whose folder segment ends with `.` (#450).
 */
export const opcNameIssueSeverity = (profile: Profile, format: PackageFormat): "error" | "warning" =>
    profile === "strict" || (profile === "word-valid" && format === "docx") ? "error" : "warning";

/**
 * Promotions applied by full format validators to registry findings. A full
 * run that cannot stage the archive must apply these too (#475). Strict
 * rejects partial selection; the registry itself keeps this soft rule advisory.
 */
export const promoteFullCheckIssue = (issue: ValidationIssue, profile: Profile, format: PackageFormat): ValidationIssue =>
    format === "docx" && profile === "strict" && issue.code === "comment-thread-paraid-missing" ? { ...issue, severity: "error" } : issue;

/**
 * True when the request asks for part of the check set rather than all of it.
 * A partial run never constructs a schema validator, so callers use this to
 * skip the XSD-engine startup probe as well.
 */
export const isPartialRequest = (request: ModeRequest): boolean => request.quick === true || (request.checks?.length ?? 0) > 0;

/**
 * True when a full run of this format under this profile compiles and walks
 * XSD schemas.
 *
 * A DOCX walks them under `lenient` and `strict`. Under `word-valid` it is
 * judged by structural Word checks and compiles no schema, so it never needs
 * the XSD engine: such a run must neither probe the engine nor fail because
 * the engine is broken (#318). The PPTX validator walks its schemas under
 * every profile, and no format validator runs for an XLSX (#341).
 */
export const profileWalksSchemas = (format: PackageFormat, profile: Profile): boolean =>
    format === "pptx" || (format === "docx" && (profile === "lenient" || profile === "strict"));

/**
 * Reject a request whose selectors contradict each other. Strict is the
 * complete mode, so it cannot be paired with partial check selection, and it
 * cannot be paired with a profile that would silently win over it. A partial
 * run never repairs, so it cannot be asked to: the request is refused rather
 * than answered with `repairs: 0` and no word that nothing was tried (#652).
 *
 * @throws {ModeConflictError}
 */
export const assertModeCoherent = (request: ModeRequest): void => {
    const partial = isPartialRequest(request);
    if (request.quick === true && (request.checks?.length ?? 0) > 0) {
        throw new ModeConflictError("Cannot combine --quick with --checks: --quick is a fixed preset, --checks names individual checks.");
    }
    if (request.strict === true && partial) {
        throw new ModeConflictError(
            "Cannot combine --strict with --quick or --checks: strict validation cannot be limited to part of the check set.",
        );
    }
    if (request.strict === true && request.profile !== undefined && request.profile !== "strict") {
        throw new ModeConflictError(`Cannot combine --strict with --profile ${request.profile}: strict is already a complete mode.`);
    }
    if (request.profile === "strict" && partial) {
        throw new ModeConflictError(
            "Cannot combine profile 'strict' with --quick or --checks: strict validation cannot be limited to part of the check set.",
        );
    }
    if (request.autoRepair === true && partial) {
        throw new ModeConflictError("Cannot combine --auto-repair with --quick or --checks: only a full run repairs.");
    }
};

/**
 * Resolve a request into the mode a run executes under.
 *
 * Deliberately pure: no I/O, no registry, no filesystem. Callers resolve
 * selections against the check registry afterwards, which is where unknown ids
 * are reported.
 *
 * @throws {ModeConflictError} when the selectors contradict each other.
 */
export const resolveMode = (format: PackageFormat, request: ModeRequest): ResolvedMode => {
    assertModeCoherent(request);
    return {
        profile: request.strict === true ? "strict" : (request.profile ?? defaultProfileForFormat(format)),
        select: request.quick === true ? QUICK_CHECK_GROUPS : request.checks?.length ? [...request.checks] : undefined,
    };
};
/**
 * The severity a finding for spec-valid markup Word was observed to refuse
 * carries under a profile and a format: an error under word-valid and strict
 * for DOCX/DOCM — Word 16.115 offers recovery for Markup Compatibility in the
 * content-types stream (#459; probe session 14a) and for the part-name
 * spellings `opc-part-name-word-refused` and `ct-override-partname-word-refused`
 * report (#408, #443, #451; sessions 6–12), and strict is the superset of
 * word-valid — and a warning under lenient, and under every profile for the
 * unobserved PPTX/XLSX. Unlike {@link opcNameIssueSeverity}, strict does not
 * raise the severity for the other formats: the markup is spec-valid and
 * nothing is observed there.
 */
export const wordObservedIssueSeverity = (profile: Profile, format: PackageFormat): "error" | "warning" =>
    format === "docx" && (profile === "word-valid" || profile === "strict") ? "error" : "warning";
