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

import { DEFAULT_PROFILE, type PackageFormat, type Profile } from "./types";

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
};

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
 * True when the request asks for part of the check set rather than all of it.
 * A partial run never constructs a schema validator, so callers use this to
 * skip the XSD-engine startup probe as well.
 */
export const isPartialRequest = (request: ModeRequest): boolean => request.quick === true || (request.checks?.length ?? 0) > 0;

/**
 * Resolve a request into the mode a run executes under.
 *
 * Deliberately pure: no I/O, no registry, no filesystem. Callers resolve
 * selections against the check registry afterwards, which is where unknown ids
 * are reported.
 */
export const resolveMode = (format: PackageFormat, request: ModeRequest): ResolvedMode => ({
    profile: request.strict === true ? "strict" : (request.profile ?? defaultProfileForFormat(format)),
    select: request.quick === true ? QUICK_CHECK_GROUPS : request.checks?.length ? [...request.checks] : undefined,
});
