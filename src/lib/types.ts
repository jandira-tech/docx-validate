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
 * Shared types and OOXML namespace constants used across docx-validate-ts.
 *
 * Convention for porters: every validator returns `ValidationResult`. Do not
 * invent a per-validator shape. Add fields to `ValidationResult` (and update
 * `mergeResults` if you do) rather than forking.
 */

/**
 * Severity of a single validation finding. Mirrors the Python validators'
 * three-tier model: ERROR (XSD violation, malformed XML), WARNING (suspicious
 * but not a hard schema breach), INFO (advisory only — does not affect
 * `valid`).
 */
export type Severity = "error" | "warning" | "info";

/**
 * One finding from a validator. `path` is the part name inside the OOXML
 * package (e.g. "word/document.xml"); `line`/`column` are 1-based when
 * available. `code` is a stable string identifier so tests can assert on
 * specific findings without matching prose.
 */
export type ValidationIssue = {
    severity: Severity;
    message: string;
    path?: string;
    line?: number;
    column?: number;
    code?: string;
};

/**
 * What a real-Microsoft-Word observation did with the file, as it appears on
 * {@link ValidationMetadata}. A structural mirror of `WordVerification` from
 * `lib/word-verifier.ts` — declared here so `lib/types.ts` stays free of
 * imports from the verifier (which itself imports `lib/types.ts`).
 */
export type WordVerificationMetadata = {
    requested: boolean;
    satisfied: boolean;
    outcome: string;
    platform: string;
    details: string;
    observedAt: string;
    observationMs: number;
    fileSha256: string;
    wordVersion?: string;
    wordBuild?: string;
};

/**
 * Execution metadata for one validation run (plan § 1, tasks 10-11). Answers
 * "what actually ran", so a caller can tell a complete run from a partial one:
 * a partial run is still `valid` when its selected checks passed, but
 * `omittedChecks` is non-empty and the caller is expected to say so.
 *
 * Every field is optional so a validator that knows nothing about modes can
 * still return a `ValidationResult` unchanged.
 */
export type ValidationMetadata = {
    /** The profile the run was executed under. */
    profile?: Profile;
    /** Check ids this run covered, in registry order. */
    executedChecks?: readonly string[];
    /** Applicable check ids this run left out. Empty for a complete run. */
    omittedChecks?: readonly string[];
    /**
     * True only when every schema-eligible part was validated and no schema
     * failed to load. False when schema validation never started, stopped
     * early, or the engine could not compile a schema.
     */
    schemaCompleted?: boolean;
    /** Present when real-Word verification ran. */
    wordVerification?: WordVerificationMetadata;
};

/** Aggregate result returned by every validator. `valid` is false iff there is at least one `error` issue. */
export type ValidationResult = {
    valid: boolean;
    issues: ValidationIssue[];
    /** What the run executed. Attached by the top-level `validate()`, never by an inner validator. */
    metadata?: ValidationMetadata;
};

/**
 * What the XSD walk actually did. `mergeResults` keeps only `valid` and
 * `issues`, so `validate()` attaches this after the merge.
 *
 * `started` is true once `validateAgainstXsd` is entered. `parts` are the
 * relative paths sent to the engine. `eligible` is how many parts a full
 * walk would send (a schema is mapped for the part's conformance class; a
 * Strict root counts only under the strict profile).
 */
export type XsdWalkTrace = {
    started: boolean;
    parts: string[];
    eligible: number;
};

/** Walk trace for a path that never entered `validateAgainstXsd`. */
export const xsdWalkNotStarted = (): XsdWalkTrace => ({ started: false, parts: [], eligible: 0 });

/** Empty success result (zero issues, valid). */
export const OK_RESULT: ValidationResult = Object.freeze({
    valid: true,
    issues: [],
});

/**
 * Combine multiple results into one; `valid` is the AND of all inputs.
 *
 * Keeps only `valid` and `issues`, so per-run `metadata` (and the older
 * `xsdWalk` trace) is dropped here and re-attached by the caller that knows
 * what the whole run executed.
 */
export const mergeResults = (...results: ValidationResult[]): ValidationResult => {
    const issues = results.flatMap((r) => r.issues);
    const valid = results.every((r) => r.valid);
    return { valid, issues };
};

/**
 * OOXML namespace URIs. Used by xml-helpers `getElementsByTagNameNSAll`,
 * by `xpath` queries, and by the validators when emitting `code` strings.
 *
 * Names match the prefixes used in the OOXML spec (and in the Python source)
 * so cross-referencing the two trees is straightforward.
 */
export const NS = {
    /** WordprocessingML — `w:` */
    W: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    /** Word 2010 extensions — `w14:` (paraId, durableId, …) */
    W14: "http://schemas.microsoft.com/office/word/2010/wordml",
    /** Word 2012 extensions — `w15:` */
    W15: "http://schemas.microsoft.com/office/word/2012/wordml",
    /** Word 2014/16 — `w16:` family */
    W16CEX: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    W16CID: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    /** PresentationML — `p:` */
    P: "http://schemas.openxmlformats.org/presentationml/2006/main",
    /** SpreadsheetML — `s:` (also `x:` in some places) */
    S: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    /** DrawingML — `a:` */
    A: "http://schemas.openxmlformats.org/drawingml/2006/main",
    /** Relationships — `r:` */
    R: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    /** Package relationships (`.rels`) */
    PR: "http://schemas.openxmlformats.org/package/2006/relationships",
    /** Content Types */
    CT: "http://schemas.openxmlformats.org/package/2006/content-types",
    /** Markup Compatibility — `mc:` */
    MC: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    /** xml: namespace (xml:space, xml:lang) */
    XML: "http://www.w3.org/XML/1998/namespace",
} as const;

/** Default xpath select factory namespace map; mirror of `NS` keyed by lower-case prefix. */
export const XPATH_NS: Record<string, string> = {
    w: NS.W,
    w14: NS.W14,
    w15: NS.W15,
    w16cex: NS.W16CEX,
    w16cid: NS.W16CID,
    p: NS.P,
    s: NS.S,
    a: NS.A,
    r: NS.R,
    pr: NS.PR,
    ct: NS.CT,
    mc: NS.MC,
    xml: NS.XML,
};

/** Office package format inferred from a file path or content-types stream. */
export type PackageFormat = "docx" | "pptx" | "xlsx";

/**
 * Validation profile — selects how strictly the validators interpret
 * tolerated-but-non-canonical constructs.
 *
 * - `"lenient"` (default): matches real-world Microsoft Office output. UTF-8
 *   BOMs at the start of XML parts are accepted (the XML spec permits them
 *   and Word/Office 365 routinely emits them).
 * - `"strict"`: spec-purist. Reports every tolerated-but-non-canonical
 *   construct it can see. A leading BOM is still only a `warning`
 *   (`xml-bom-leading`) under every profile: XML 1.0 permits it and Microsoft
 *   Office emits it, so it never fails a document on its own. All other
 *   checks run identically.
 * - `"word-valid"`: empirical Microsoft Word openability profile. It keeps
 *   checks that correlate with Word's unreadable-content/open-error dialogs as
 *   errors and downgrades tolerated OOXML/schema quirks to warnings.
 */
export type Profile = "lenient" | "strict" | "word-valid";

/** Default validation profile. Lenient matches real-world Microsoft Office output. */
export const DEFAULT_PROFILE: Profile = "lenient";
