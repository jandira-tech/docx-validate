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
 * Port target: src/docx-validate/scripts/office/validate.py — task #15.
 *
 * Top-level validation CLI. Dispatches to the DOCX or PPTX schema validator
 * (and the redlining validator for DOCX with --original) based on the file
 * extension of either the unpacked target or the supplied --original.
 *
 * CLI:
 *   tsx validate.ts <path> [--original <file>] [-v|--verbose] [--auto-repair]
 *                          [--author <name>]
 *
 * `<path>` may be either an unpacked directory or a packed `.docx`/`.pptx`/
 * `.xlsx` file (the latter is unzipped to a temp dir for the duration of the
 * run). Exit code: 0 when every check passes, 1 otherwise.
 */

import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";
import JSZip from "jszip";

import { commanderExitCode, runCli, withTempDir } from "../../lib/run-cli";
import { withoutConsequentialParseFailures } from "../../lib/parse-failures";
import type { CheckRegistry } from "../../lib/check-registry";
import { runChecks, type CheckRun } from "../../lib/check-scheduler";
import {
    assertModeCoherent,
    isPartialRequest,
    ModeConflictError,
    profileWalksSchemas,
    QUICK_CHECK_GROUPS,
    resolveMode,
    type ResolvedMode,
} from "../../lib/mode-policy";
import {
    mergeResults,
    xsdWalkNotStarted,
    type PackageFormat,
    type Profile,
    type ValidationIssue,
    type ValidationResult,
    type XsdWalkTrace,
} from "../../lib/types";
import type { XsdValidator } from "../../lib/xsd-validator";
import { verifyInWord, type WordVerification } from "../../lib/word-verifier";
import { indexByPartName, normalizePartName, ValidationContext } from "../../lib/validation-context";
import { extractPackageParts } from "../../lib/zip-path";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "./checks";
import { XML_SYNTAX_CHECK_ID } from "./checks/xml";
import { OPC_EQUIVALENT_PART_NAMES } from "./checks/opc-equivalent-part-names";
import { contextFromArchive, contextFromDirectory } from "./package-context";
import { BaseSchemaValidator, walkFiles } from "./validators/base";
import { DOCXSchemaValidator } from "./validators/docx";
import { buildRepairPlanIssues, collectDocxSemanticInventory, compareDocxSemanticInventories } from "./validators/docx-diagnostics";
import { PPTXSchemaValidator } from "./validators/pptx";
import { validateRedlining } from "./validators/redlining";
import { isWordBlockingIssue } from "./validators/word-blocking";
import { explainWordError } from "./validators/word-error-explanations";

const SUPPORTED_SUFFIXES = new Set([".docm", ".docx", ".pptx", ".xlsx"]);

/**
 * Side-effect boundary for real-Word verification. Defaults to
 * `verifyInWord` (a temporary copy, opened through the injectable
 * `WordTransport`); tests substitute their own so no lane opens Word.
 */
export type WordVerifier = (bytes: Uint8Array) => Promise<WordVerification>;

export interface ValidateOptions {
    /** Path to the original packed Office file. Required for DOCX redlining. */
    original?: string;
    /** Run auto-repair before validating (mirrors `--auto-repair`). */
    autoRepair?: boolean;
    /** Author name to use for redlining validation. */
    author?: string;
    /** Verbose mode — passes through to subclasses' verbose flags. */
    verbose?: boolean;
    /** Override XSD root (defaults to the path used by `BaseSchemaValidator`). */
    schemasDir?: string;
    /**
     * Validation profile. Defaults per format: `"word-valid"` for
     * `.docx`/`.docm` (Microsoft Word openability, no XSD walk) and
     * `"lenient"` for every other format, which matches real-world
     * Microsoft Office output. Pass `"strict"` for spec-purist behaviour
     * (flags BOM-prefixed parts and other tolerated-but-non-canonical
     * constructs). See {@link Profile}.
     */
    profile?: Profile;
    /** `--strict`: shorthand for `profile: "strict"`. */
    strict?: boolean;
    /** `--quick`: run the common-defect preset and disclose the omissions. */
    quick?: boolean;
    /** `--checks`: run these check ids or group names, plus their prerequisites. */
    checks?: readonly string[];
    /**
     * `--verify-word`: after the portable mode passes, ask Microsoft Word
     * about the exact bytes. Requested-but-unavailable or inconclusive
     * verification is NOT satisfied, so the run reports `valid: false`.
     */
    verifyWord?: boolean;
    /** Override the real-Word verifier. Defaults to {@link verifyInWord}. */
    wordVerifier?: WordVerifier;
    /**
     * Override the XSD engine for every schema validator this run creates.
     * When omitted, `BaseSchemaValidator` lazy-creates the default
     * libxml2-wasm validator.
     */
    xsdValidator?: XsdValidator;
}

/** Package format a dispatch suffix maps to. `.docm` validates exactly like `.docx`. */
export function formatForSuffix(suffix: string): PackageFormat {
    if (suffix === ".pptx") return "pptx";
    if (suffix === ".xlsx") return "xlsx";
    return "docx";
}

export interface ValidateRunResult extends ValidationResult {
    /** Number of issues auto-repaired (0 unless `autoRepair` is true). */
    repairs: number;
    /** Detected suffix used for dispatch (e.g. ".docx"). */
    suffix: string;
    /**
     * XSD engine trace. Attached here, after `mergeResults`, because the
     * merge keeps only `valid` and `issues`.
     */
    xsdWalk: XsdWalkTrace;
}

/**
 * Public entry point — mirrors `main()` in
 * `src/docx-validate/scripts/office/validate.py` minus the argv parsing. Use
 * this from other TS code; the CLI shim at the bottom of the file just turns
 * argv into options and exits with the right code.
 */
export async function validate(target: string, opts: ValidateOptions = {}): Promise<ValidateRunResult> {
    const author = opts.author;
    const verbose = opts.verbose ?? false;
    const autoRepair = opts.autoRepair ?? false;
    const original = opts.original;

    // Resolve the mode before touching the filesystem: a contradictory
    // selection is a configuration error, not a verdict about a document.
    const dispatchSuffix = path.extname(original ?? target).toLowerCase();
    const format = formatForSuffix(dispatchSuffix);
    const mode = resolveMode(format, opts);
    const profile: Profile = mode.profile;
    const registry = createDefaultCheckRegistry();
    const selected = mode.select === undefined ? undefined : selectionUnderProfile(registry, mode.select, format, profile);

    await assertExists(target, `Error: ${target} does not exist`);

    let originalFile: string | null = null;
    if (original) {
        await assertIsFile(original, `Error: ${original} is not a file`);
        const ext = path.extname(original).toLowerCase();
        if (!SUPPORTED_SUFFIXES.has(ext)) {
            throw new Error(`Error: ${original} must be a .docx, .docm, .pptx, or .xlsx file`);
        }
        originalFile = path.resolve(original);
    }

    if (!SUPPORTED_SUFFIXES.has(dispatchSuffix)) {
        throw new Error(`Error: Cannot determine file type from ${target}. Use --original or provide a .docx/.docm/.pptx/.xlsx file.`);
    }

    const targetStat = await fs.stat(target);
    const targetSuffix = path.extname(target).toLowerCase();
    const targetIsPackedFile = targetStat.isFile() && SUPPORTED_SUFFIXES.has(targetSuffix);

    if (!targetIsPackedFile && !targetStat.isDirectory()) {
        throw new Error(`Error: ${target} is not a directory or Office file`);
    }

    // Read the exact bytes once, before any validation work, so Word is asked
    // about the same file the portable mode inspected.
    const packedBytes = opts.verifyWord === true && targetIsPackedFile ? await fs.readFile(target) : undefined;
    const finish = async (result: ValidateRunResult): Promise<ValidateRunResult> =>
        opts.verifyWord === true ? withWordVerification(result, packedBytes, opts.wordVerifier) : result;

    if (selected !== undefined) {
        return finish(await runSelectedChecks(target, targetIsPackedFile, dispatchSuffix, format, mode, registry, selected));
    }

    const runWithUnpacked = async (unpackedDir: string, extra = EMPTY_ENVELOPE_RUN): Promise<ValidateRunResult> => {
        const subclassResult = await runValidators(unpackedDir, {
            originalFile,
            suffix: dispatchSuffix,
            author,
            verbose,
            autoRepair,
            schemasDir: opts.schemasDir,
            xsdValidator: opts.xsdValidator,
            profile,
        });
        return {
            ...subclassResult,
            valid: subclassResult.valid && extra.issues.every((issue) => issue.severity !== "error"),
            issues: [...subclassResult.issues, ...extra.issues],
            suffix: dispatchSuffix,
            metadata: {
                profile,
                // Only the checks this run really executed (#319).
                executedChecks: orderedExecutedChecks(format, [...subclassResult.executedChecks, ...extra.executed]),
                omittedChecks: [],
                schemaCompleted: schemaWalkCompleted(subclassResult.xsdWalk, subclassResult.engineFailures),
            },
        };
    };

    // A packed input, or a directory holding a stored name that differs from
    // its part name (`word\document.xml`), is judged through the package view
    // staged under part names, so the path-keyed validators see one entry per
    // part (#371, #374). `sources` maps each staged part to the directory file
    // it came from, so `--auto-repair` keeps writing to the stored names.
    const runStaged = async (
        zip: JSZip,
        envelope: { issues: ValidationResult["issues"]; executed: readonly string[] },
        sources?: ReadonlyMap<string, string>,
    ): Promise<ValidateRunResult> =>
        withTempDir(async (tempDir) => {
            try {
                await extractPackageParts(zip, tempDir);
            } catch (err) {
                const detail = err instanceof Error ? err.message : String(err);
                return {
                    valid: false,
                    issues: [
                        ...envelope.issues,
                        {
                            severity: "error" as const,
                            message: `Cannot extract ${target}: ${detail}`,
                            code: "package-extract-failed",
                        },
                    ],
                    repairs: 0,
                    suffix: dispatchSuffix,
                    xsdWalk: xsdWalkNotStarted(),
                    // The envelope was read from the bytes before extraction failed,
                    // so its checks ran (CodeRabbit on #335).
                    metadata: {
                        profile,
                        executedChecks: orderedExecutedChecks(format, envelope.executed),
                        omittedChecks: [],
                        schemaCompleted: false,
                    },
                };
            }
            const before = new Map<string, Buffer>();
            if (autoRepair && sources) {
                for (const part of sources.keys()) before.set(part, await fs.readFile(path.join(tempDir, part)));
            }
            const unpacked = await runWithUnpacked(tempDir, envelope);
            // A staged directory still repairs its source, under stored names.
            // Bytes are compared so untouched files are never rewritten, and a
            // dropped duplicate entry has no write-back destination at all.
            if (autoRepair && sources) {
                for (const [part, file] of sources) {
                    let repaired: Buffer;
                    try {
                        repaired = await fs.readFile(path.join(tempDir, part));
                    } catch (err) {
                        if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
                        await fs.rm(file);
                        continue;
                    }
                    if (!repaired.equals(before.get(part)!)) await fs.writeFile(file, repaired);
                }
            }
            return unpacked;
        });

    if (targetIsPackedFile) {
        const buf = await fs.readFile(target);
        const envelope = runsWordPackageChecks(profile) ? await zipEnvelopeRun(buf, dispatchSuffix) : EMPTY_ENVELOPE_RUN;
        let zip: JSZip;
        try {
            zip = await JSZip.loadAsync(buf);
        } catch (err) {
            // The file is not a readable OPC (zip) package: encrypted OOXML
            // (a CFB/OLE compound file), or a truncated/corrupted archive.
            // Word refuses to open all of these, so this is an error — but
            // `validate()` must NEVER throw here (a thrown validator silently
            // crashes callers that shell out, e.g. python-jubarte, instead of
            // reporting the file invalid). Surface it as a structured issue.
            // Kept blocking under every profile: an unopenable package is not
            // something Word tolerates.
            const detail = err instanceof Error ? err.message : String(err);
            return finish({
                valid: false,
                issues: [...envelope.issues, packageOpenFailedIssue(target, detail)],
                repairs: 0,
                suffix: dispatchSuffix,
                xsdWalk: xsdWalkNotStarted(),
                // Only the envelope checks could read these bytes (#332).
                metadata: {
                    profile,
                    executedChecks: orderedExecutedChecks(format, envelope.executed),
                    omittedChecks: [],
                    schemaCompleted: false,
                },
            });
        }
        // Diagnose the archive's own names before extraction can reject
        // case/normalization collisions on the host filesystem (#351).
        const names = Object.values(zip.files)
            .filter((entry) => !entry.dir)
            .map((entry) => [entry.unsafeOriginalName ?? entry.name, ""] as const);
        const diagnostic = await equivalentNamesRun(ValidationContext.fromEntries(names), format, profile);
        return finish(
            await runStaged(zip, {
                issues: [...envelope.issues, ...diagnostic.issues],
                executed: [...envelope.executed, ...diagnostic.executed],
            }),
        );
    }

    const dir = path.resolve(target);
    const diagnostic = await equivalentNamesRun(contextFromDirectory(dir), format, profile);
    const entries = walkFiles(dir).map((file) => [path.relative(dir, file).split(path.sep).join("/"), file] as const);
    if (entries.some(([name]) => name !== normalizePartName(name))) {
        // JSZip is only an in-memory entry collection here: no ZIP is written.
        const zip = new JSZip();
        for (const [name, file] of entries) zip.file(name, await fs.readFile(file), { createFolders: false });
        return finish(await runStaged(zip, diagnostic, indexByPartName(entries)));
    }
    return finish(await runWithUnpacked(dir, diagnostic));
}

/**
 * Run the requested portable mode's verdict through real Microsoft Word when
 * the caller asked for it. Verification that was requested but could not run
 * (no bytes, another platform, Word not installed) or that stayed
 * inconclusive (timeout, unknown dialog) is NOT satisfied, so the run reports
 * `valid: false` — an unanswered question is not a pass.
 */
async function withWordVerification(
    result: ValidateRunResult,
    packedBytes: Uint8Array | undefined,
    injected: WordVerifier | undefined,
): Promise<ValidateRunResult> {
    const verify = injected ?? verifyInWord;
    if (packedBytes === undefined) {
        return applyWordVerification(result, {
            requested: true,
            outcome: "unavailable",
            satisfied: false,
            fileSha256: "",
            platform: process.platform,
            // Wall time of this answer, never the epoch (#329).
            observedAt: new Date().toISOString(),
            observationMs: 0,
            details: "Word verification needs a packed Office file; the validated target is an unpacked directory",
        });
    }
    try {
        return applyWordVerification(result, await verify(packedBytes));
    } catch (err) {
        // A verifier that throws could not run: report that, never reject (#333).
        return applyWordVerification(result, {
            requested: true,
            outcome: "unavailable",
            satisfied: false,
            fileSha256: createHash("sha256").update(packedBytes).digest("hex"),
            platform: process.platform,
            observedAt: new Date().toISOString(),
            observationMs: 0,
            details: `Word verification failed to run: ${err instanceof Error ? err.message : String(err)}`,
        });
    }
}

function applyWordVerification(result: ValidateRunResult, verification: WordVerification): ValidateRunResult {
    return {
        ...result,
        valid: result.valid && verification.satisfied,
        metadata: { ...(result.metadata ?? {}), wordVerification: { ...verification } },
    };
}

/**
 * The named checks a run executed, in registry order.
 *
 * The registry lists every check for a format, but a profile only reaches
 * some of them — the Word rules run under `word-valid` and `strict` only —
 * so the disclosure lists what ran, not what exists (#319).
 */
function orderedExecutedChecks(format: PackageFormat, executed: readonly string[]): readonly string[] {
    const ran = new Set(executed);
    return createDefaultCheckRegistry()
        .list(format)
        .map((def) => def.id)
        .filter((id) => ran.has(id));
}

/**
 * The named checks a full run under `profile` never reaches. Lenient runs
 * neither the ZIP-envelope checks ({@link runsWordPackageChecks}) nor the Word
 * rules (`WORD_RULE_CHECKS`), so a partial lenient run leaves them out too and
 * is never stricter than the full run (#381).
 */
function checksOutsideProfile(registry: CheckRegistry, profile: Profile): ReadonlySet<string> {
    if (profile !== "lenient") return new Set();
    const envelope = registry.list().filter((definition) => definition.archiveOnly === true);
    return new Set([...WORD_RULE_CHECKS, ...envelope].map((definition) => definition.id));
}

/**
 * The ids a partial run selects under its profile: each named check, and each
 * member of a named group, that the full run under that profile reaches.
 * Naming a check the profile never runs, or only groups with no check it runs,
 * asks for something the profile does not do (#381).
 *
 * @throws {ModeConflictError}
 */
function selectionUnderProfile(
    registry: CheckRegistry,
    select: readonly string[],
    format: PackageFormat,
    profile: Profile,
): readonly string[] {
    // Unknown and inapplicable names fail as they always have.
    registry.resolve(select, format);
    const outside = checksOutsideProfile(registry, profile);
    const refused = select.filter((name) => outside.has(name));
    if (refused.length > 0) {
        throw new ModeConflictError(
            `Cannot run ${refused.join(", ")} under profile '${profile}': it runs neither the Word rules nor the ZIP-envelope checks. Use --profile word-valid.`,
        );
    }
    const members = registry
        .list(format)
        .filter((definition) => select.includes(definition.id) || definition.groups.some((group) => select.includes(group)))
        .map((definition) => definition.id);
    const ids = members.filter((id) => !outside.has(id));
    // A recognized group with no check applicable to this format still runs
    // nothing and omits everything, under every profile.
    if (members.length === 0) return select;
    if (ids.length === 0) {
        throw new ModeConflictError(
            `No check in ${select.join(", ")} runs under profile '${profile}': it runs neither the Word rules nor the ZIP-envelope checks. Use --profile word-valid.`,
        );
    }
    return ids;
}

/**
 * A partial run's findings as the full run reports them. Under `word-valid`, an
 * error from a DOCX check that Word tolerates is a warning
 * ({@link isWordBlockingIssue}), as `DOCXSchemaValidator` reports it. A full run
 * also reads the ZIP envelope outside that validator, so the envelope checks
 * keep their severity (#381).
 */
function findingsUnderProfile(run: CheckRun, registry: CheckRegistry, format: PackageFormat, profile: Profile): ValidationIssue[] {
    // Equivalent stored names are Word-blocking for DOCX/DOCM (#351). Lenient
    // keeps the earlier policy and runs no Word rule, and no PowerPoint or
    // Excel behaviour has been observed, so those only warn.
    const advisory = (issue: ValidationIssue): ValidationIssue =>
        issue.code === OPC_EQUIVALENT_PART_NAMES && (profile === "lenient" || (profile === "word-valid" && format !== "docx"))
            ? { ...issue, severity: "warning" }
            : issue;
    if (profile !== "word-valid" || format !== "docx") return run.issues.map(advisory);
    return run.outcomes
        .flatMap((outcome) =>
            registry.get(outcome.id)?.archiveOnly === true
                ? outcome.issues
                : outcome.issues.map((issue): ValidationIssue =>
                      issue.severity === "error" && !isWordBlockingIssue(issue) ? { ...issue, severity: "warning" } : issue,
                  ),
        )
        .map(advisory);
}

/**
 * A partial run: only the selected named checks, against a context built from
 * the package. No schema validator is constructed, so nothing compiles an XSD
 * and no rule outside the selection runs. Everything the selection left out is
 * reported as an omission — `valid` here means "the requested checks passed".
 */
async function runSelectedChecks(
    target: string,
    packed: boolean,
    suffix: string,
    format: PackageFormat,
    mode: ResolvedMode,
    registry: CheckRegistry,
    selected: readonly string[],
): Promise<ValidateRunResult> {
    const reportsParseFailures = registry.resolve(selected, format).some((id) => registry.get(id)?.reportsParseFailures);
    // A companion, not a prerequisite: these checks still read every part
    // with their own parser, even when xml-syntax rejects it (#385).
    const select = reportsParseFailures ? [...selected, XML_SYNTAX_CHECK_ID] : selected;
    const opened = packed ? await contextFromArchive(await fs.readFile(target)) : { context: contextFromDirectory(path.resolve(target)) };
    if (opened.openError !== undefined) {
        // An archive that does not open has no parts: only the selected
        // envelope checks can read its bytes, and nothing else ran (#332).
        const archiveOnly = new Set(
            registry
                .list(format)
                .filter((definition) => definition.archiveOnly === true)
                .map((definition) => definition.id),
        );
        const resolved = registry.resolve(select, format);
        const envelopeSelect = resolved.filter((id) => archiveOnly.has(id));
        const envelope =
            envelopeSelect.length > 0 ? await runChecks(opened.context, registry, { format, select: envelopeSelect }) : undefined;
        const executed = envelope?.executed ?? [];
        return {
            valid: false,
            issues: [...(envelope?.issues ?? []), packageOpenFailedIssue(target, opened.openError)],
            repairs: 0,
            suffix,
            xsdWalk: xsdWalkNotStarted(),
            metadata: {
                profile: mode.profile,
                executedChecks: [...executed],
                // Omitted means "not selected", as for a package that opens; a
                // selected check the archive kept from running is not an
                // omission, and the error says why it did not run (Jules on #337).
                omittedChecks: registry
                    .list(format)
                    .map((definition) => definition.id)
                    .filter((id) => !resolved.includes(id)),
                schemaCompleted: false,
            },
        };
    }
    // A full lenient run stops at its first malformed part, so a lenient
    // partial run stops where xml-syntax fails (#391).
    const run = await runChecks(opened.context, registry, {
        format,
        select,
        ...(mode.profile === "lenient" ? { haltAfterFailing: XML_SYNTAX_CHECK_ID } : {}),
    });
    const malformed = new Set(
        run.outcomes
            .filter((outcome) => outcome.id === XML_SYNTAX_CHECK_ID)
            .flatMap((outcome) => outcome.issues)
            .flatMap((issue) => (issue.severity === "error" && issue.path !== undefined ? [issue.path] : [])),
    );
    const outcomes = run.outcomes.map((outcome) =>
        registry.get(outcome.id)?.reportsParseFailures
            ? { ...outcome, issues: withoutConsequentialParseFailures(outcome.issues, malformed) }
            : outcome,
    );
    const issues = findingsUnderProfile(
        { ...run, outcomes, issues: outcomes.flatMap((outcome) => outcome.issues) },
        registry,
        format,
        mode.profile,
    );
    return {
        valid: issues.every((issue) => issue.severity !== "error"),
        issues,
        repairs: 0,
        suffix,
        xsdWalk: xsdWalkNotStarted(),
        metadata: {
            profile: mode.profile,
            // What ran, not what was selected: an archive-only check is
            // unavailable for an unpacked directory (CodeRabbit on #322).
            executedChecks: [...run.executed],
            omittedChecks: [...run.omitted],
            schemaCompleted: false,
        },
    };
}

/**
 * The finding for a packed file that is not a readable OPC (zip) package:
 * encrypted OOXML, or a truncated or corrupted archive. Word refuses to open
 * all of these, so it is blocking under every profile.
 */
function packageOpenFailedIssue(target: string, detail: string): ValidationIssue {
    return {
        severity: "error",
        message: `Cannot open ${target} as an OPC package (not a valid .docx/zip — encrypted, corrupted, or truncated): ${detail}`,
        code: "package-open-failed",
    };
}

/**
 * A schema walk "completed" when it started, offered every eligible part to
 * the engine, and the engine answered for each one (no schema failed to load).
 */
function schemaWalkCompleted(walk: XsdWalkTrace, engineFailures: number): boolean {
    return walk.started && walk.parts.length === walk.eligible && engineFailures === 0;
}

/**
 * ZIP-envelope checks (bytes after the EOCD record, EOCD comment overrun)
 * describe packages Word refuses or repairs. `strict` is a superset of
 * `word-valid`, so both run them; explicit `lenient` keeps its historical
 * behaviour for callers that have not migrated.
 */
function runsWordPackageChecks(profile: Profile): boolean {
    return profile === "word-valid" || profile === "strict";
}

/** Findings of the archive-only `package` checks on the packed file's bytes. */
async function zipEnvelopeRun(
    archiveBytes: Uint8Array,
    suffix: string,
): Promise<{ issues: ValidationResult["issues"]; executed: readonly string[] }> {
    const format: PackageFormat = suffix === ".pptx" ? "pptx" : suffix === ".xlsx" ? "xlsx" : "docx";
    const context = ValidationContext.fromEntries({}, { source: "archive", archiveBytes });
    const registry = createDefaultCheckRegistry();
    // Every archive-only package check; the stored-name diagnosis runs separately.
    const select = registry
        .list()
        .filter((definition) => definition.archiveOnly === true)
        .map((definition) => definition.id);
    const run = await runChecks(context, registry, { format, select });
    return { issues: [...run.issues], executed: [...run.executed] };
}

/** Stored-name diagnosis for every full profile, before any XML/extraction gate. */
async function equivalentNamesRun(
    context: ValidationContext,
    format: PackageFormat,
    profile: Profile,
): Promise<{ issues: ValidationIssue[]; executed: readonly string[] }> {
    const registry = createDefaultCheckRegistry();
    const run = await runChecks(context, registry, { format, select: [OPC_EQUIVALENT_PART_NAMES] });
    return { issues: findingsUnderProfile(run, registry, format, profile), executed: run.executed };
}

/** An envelope run that never happened: no findings, no executed checks. */
const EMPTY_ENVELOPE_RUN: { issues: ValidationResult["issues"]; executed: readonly string[] } = { issues: [], executed: [] };

interface ValidatorRunner {
    repair(): Promise<number>;
    validate(): Promise<ValidationResult>;
}

interface RunValidatorsOptions {
    originalFile: string | null;
    suffix: string;
    /** Required iff `originalFile` is set (DOCX redlining cross-check). */
    author: string | undefined;
    verbose: boolean;
    autoRepair: boolean;
    schemasDir?: string;
    xsdValidator?: XsdValidator;
    profile: Profile;
}

async function runValidators(
    unpackedDir: string,
    opts: RunValidatorsOptions,
): Promise<ValidationResult & { repairs: number; xsdWalk: XsdWalkTrace; engineFailures: number; executedChecks: readonly string[] }> {
    const validators: ValidatorRunner[] = [];
    let schemaValidator: BaseSchemaValidator | undefined;

    if (opts.suffix === ".docx" || opts.suffix === ".docm") {
        const docx = new DOCXSchemaValidator({
            unpackedDir,
            originalFile: opts.originalFile ?? undefined,
            verbose: opts.verbose,
            schemasDir: opts.schemasDir,
            xsdValidator: opts.xsdValidator,
            profile: opts.profile,
        });
        schemaValidator = docx;
        validators.push(docx);
        if (opts.originalFile) {
            if (!opts.author) {
                throw new Error(
                    "validate(): `author` is required when `original` is provided (used to identify whose tracked changes to verify).",
                );
            }
            const originalDocx = opts.originalFile;
            const author = opts.author;
            const verbose = opts.verbose;
            validators.push({
                repair: async () => 0,
                validate: () =>
                    validateRedlining({
                        unpackedDir,
                        originalDocx,
                        author,
                        verbose,
                    }),
            });
        }
    } else if (opts.suffix === ".pptx") {
        const pptx = new PPTXSchemaValidator({
            unpackedDir,
            originalFile: opts.originalFile ?? undefined,
            verbose: opts.verbose,
            schemasDir: opts.schemasDir,
            xsdValidator: opts.xsdValidator,
            profile: opts.profile,
        });
        schemaValidator = pptx;
        validators.push(pptx);
    } else {
        // Library code is silent (CLAUDE.md: "Validator results are
        // structured, not printed"). The CLI shim renders this issue via
        // its general issue-printing path at the bottom of
        // runValidateFromArgv, so callers driving validate() directly get
        // the structured error without stderr noise.
        return {
            valid: false,
            issues: [
                {
                    severity: "error" as const,
                    message: `Unsupported file type: ${opts.suffix}`,
                    code: "unsupported-file-type",
                },
            ],
            repairs: 0,
            xsdWalk: xsdWalkNotStarted(),
            engineFailures: 0,
            // Nothing ran: an unsupported file type never reaches a validator.
            executedChecks: [],
        };
    }

    let repairs = 0;
    let repairDiagnostics: ValidationResult = { valid: true, issues: [] };
    if (opts.autoRepair) {
        const beforeInventory =
            opts.suffix === ".docx" || opts.suffix === ".docm" ? await collectDocxSemanticInventory(unpackedDir, opts.profile) : null;
        const beforeResults = await Promise.all(validators.map((v) => v.validate()));
        const repairPlanIssues = buildRepairPlanIssues(mergeResults(...beforeResults).issues);
        for (const v of validators) {
            repairs += await v.repair();
        }
        const contentIssues =
            beforeInventory && repairs > 0
                ? compareDocxSemanticInventories(beforeInventory, await collectDocxSemanticInventory(unpackedDir, opts.profile))
                : [];
        repairDiagnostics = mergeResults({
            valid: [...repairPlanIssues, ...contentIssues].every((i) => i.severity !== "error"),
            issues: [...repairPlanIssues, ...contentIssues],
        });
    }

    const results = await Promise.all(validators.map((v) => v.validate()));
    const merged = mergeResults(...results, repairDiagnostics);
    // mergeResults drops xsdWalk. Copy the trace off the schema validator
    // after the walk, including the word-valid path that never started XSD.
    return {
        ...merged,
        repairs,
        xsdWalk: schemaValidator?.xsdWalk ?? xsdWalkNotStarted(),
        engineFailures: schemaValidator?.xsdEngineFailures ?? 0,
        // The named checks the validator really ran (#319).
        executedChecks: schemaValidator?.executedChecks ?? [],
    };
}

async function assertExists(p: string, message: string): Promise<void> {
    try {
        await fs.access(p);
    } catch {
        throw new Error(message);
    }
}

async function assertIsFile(p: string, message: string): Promise<void> {
    let stat: Awaited<ReturnType<typeof fs.stat>>;
    try {
        stat = await fs.stat(p);
    } catch {
        throw new Error(message);
    }
    if (!stat.isFile()) {
        throw new Error(message);
    }
}

interface CliOptions {
    original?: string;
    verbose: boolean;
    autoRepair: boolean;
    author?: string;
    profile?: Profile;
    quick?: boolean;
    strict?: boolean;
    verifyWord?: boolean;
    json?: boolean;
    /** Comma-separated check ids or group names. */
    checks?: string;
    listChecks?: boolean;
}

export function buildValidateCommand(): Command {
    const cmd = new Command();
    cmd.name("validate")
        .description("Validate an Office document: XML structure, tracked changes, and Microsoft Word compatibility")
        .argument("[path]", "Path to unpacked directory or packed Office file (.docx/.docm/.pptx/.xlsx)")
        .option(
            "--original <file>",
            "Path to original file (.docx/.docm/.pptx/.xlsx). If omitted, all XSD errors are reported and redlining validation is skipped.",
        )
        .option("-v, --verbose", "Enable verbose output", false)
        .option("--auto-repair", "Automatically repair common issues (hex IDs, whitespace preservation)", false)
        .option("--author <name>", "Author name for redlining validation (required when --original is provided)")
        .option(
            "--profile <profile>",
            "Validation profile: 'word-valid' (default for .docx/.docm), 'lenient' (default for every other format), or 'strict'",
        )
        .option("--quick", "Run the common-defect check preset and report the checks it omits", false)
        .option("--strict", "Run the complete Word check set and require schema validation to complete", false)
        .option("--verify-word", "Also ask Microsoft Word to open the exact file; unavailable or inconclusive is not a pass", false)
        .option("--checks <ids-or-groups,...>", "Run only these named checks (or groups) plus their prerequisites")
        .option("--list-checks", "List every named check with its id, groups, formats and description, then exit", false)
        .option("--json", "Print one JSON document (every issue plus execution metadata) on stdout, with no human text", false);
    return cmd;
}

/** Split `--checks a,b , c` into trimmed, non-empty names. */
const parseCheckList = (raw: string | undefined): string[] =>
    (raw ?? "")
        .split(",")
        .map((name) => name.trim())
        .filter((name) => name !== "");

/**
 * Human-readable catalogue of the named checks, one per line:
 * `id — groups: … | formats: … — description`.
 */
export function renderCheckList(): string {
    const lines = ["Named checks (pass ids or groups to --checks):"];
    for (const def of createDefaultCheckRegistry().list()) {
        lines.push(`  ${def.id} — groups: ${def.groups.join(", ")} | formats: ${def.formats.join(", ")} — ${def.description}`);
    }
    return `${lines.join("\n")}\n`;
}

export async function runValidateFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildValidateCommand();
    cmd.exitOverride();
    // Commander throws CommanderError under exitOverride() for missing args,
    // invalid options, and --help. Catch here so they bypass the validator
    // try/catch below (which would otherwise format them as validate() errors).
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [target] = cmd.args;

    // Discovery validates nothing: it neither reads a package nor initialises
    // the XML or XSD engines.
    if (opts.listChecks === true) {
        process.stdout.write(renderCheckList());
        return 0;
    }

    if (target === undefined) {
        process.stderr.write("Error: missing required argument 'path' (or use --list-checks).\n");
        return 1;
    }

    if (opts.profile !== undefined && opts.profile !== "lenient" && opts.profile !== "strict" && opts.profile !== "word-valid") {
        const bad = String(opts.profile);
        process.stderr.write(`Invalid --profile: ${bad}. Must be 'lenient', 'strict', or 'word-valid'.\n`);
        return 1;
    }

    const checks = parseCheckList(opts.checks);

    // Reject contradictory selectors before any validation work, so the caller
    // gets one clear line instead of a verdict on a run they did not ask for.
    try {
        assertModeCoherent({ profile: opts.profile, strict: opts.strict, quick: opts.quick, checks });
    } catch (err) {
        process.stderr.write(`${err instanceof Error ? err.message : String(err)}\n`);
        return 1;
    }

    // Fail loudly at startup if the wasm-backed XSD engine is broken — otherwise
    // the per-file pipeline silently turns the same condition into per-file
    // "xsd-validation-failed" errors that look like document corruption.
    //
    // Only a run that walks schemas needs the engine. The default DOCX run is
    // `word-valid`, which judges the package with structural Word checks and
    // compiles no schema, so it must not fail because the engine is broken
    // (#318). A PPTX walks its schemas under every profile, so it always
    // probes; an XLSX reaches no validator, so it never does (#341). A partial
    // run never touches the engine either, so it skips the probe (and so does
    // any invocation that exits before validating).
    const dispatchSuffix = path.extname(opts.original ?? target).toLowerCase();
    const format = formatForSuffix(dispatchSuffix);
    const profile = resolveMode(format, {
        profile: opts.profile,
        strict: opts.strict,
        quick: opts.quick,
        checks,
    }).profile;
    if (!isPartialRequest({ quick: opts.quick, checks }) && profileWalksSchemas(format, profile)) {
        try {
            await BaseSchemaValidator.assertXsdValidationAvailable();
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            process.stderr.write(`${message}\n`);
            return 1;
        }
    }

    let result: ValidateRunResult;
    try {
        result = await validate(target, {
            original: opts.original,
            autoRepair: opts.autoRepair,
            author: opts.author,
            verbose: opts.verbose,
            profile: opts.profile,
            strict: opts.strict,
            quick: opts.quick,
            checks,
            verifyWord: opts.verifyWord,
        });
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        process.stderr.write(`${message}\n`);
        return 1;
    }

    // --json owns stdout: one document, every issue (warnings and their
    // locations included) plus the execution metadata, and nothing else. The
    // exit code still carries the verdict.
    if (opts.json === true) {
        process.stdout.write(`${JSON.stringify({ valid: result.valid, issues: result.issues, metadata: result.metadata ?? {} })}\n`);
        return result.valid ? 0 : 1;
    }

    if (opts.autoRepair && result.repairs > 0) {
        process.stdout.write(`Auto-repaired ${result.repairs} issue(s)\n`);
    }

    // A partial run is not a verdict on the whole document, so it discloses
    // what it left out and never claims that every validation passed. The
    // omitted registry ids are only part of that: the format validator's own
    // rules and the schema walk have no id, so they are named separately. A
    // PPTX run omits no id at all, yet skips both (#339).
    const partial = isPartialRequest({ quick: opts.quick, checks });
    const omitted = result.metadata?.omittedChecks ?? [];
    if (partial) {
        if (omitted.length > 0) {
            process.stdout.write(`Omitted checks (not run): ${omitted.join(", ")}\n`);
        }
        if (opts.quick === true) {
            process.stdout.write(`Quick preset groups: ${QUICK_CHECK_GROUPS.join(", ")}\n`);
        }
        process.stdout.write(`${partialRunGaps(dispatchSuffix, profile, opts.original !== undefined)}\n`);
    }

    if (result.valid) {
        if (!partial) {
            process.stdout.write("All validations PASSED!\n");
        } else {
            process.stdout.write(
                `Selected checks passed (${result.metadata?.executedChecks?.length ?? 0} run, ${omitted.length} omitted).\n`,
            );
        }
    } else {
        for (const issue of result.issues) {
            if (issue.severity !== "error") continue;
            const where = issue.path ? ` [${issue.path}]` : "";
            process.stderr.write(`${issue.severity.toUpperCase()}${where}: ${issue.message}\n`);
            if (result.metadata?.profile === "word-valid") {
                const explanation = explainWordError(issue.code);
                if (explanation) process.stderr.write(`  ↳ ${explanation}\n`);
            }
        }
        // An unsatisfied Word observation carries no issue of its own — say
        // so, so a non-zero exit is never silent.
        const observed = result.metadata?.wordVerification;
        if (observed && !observed.satisfied) {
            process.stderr.write(`WORD VERIFICATION ${observed.outcome}: ${observed.details}\n`);
        }
    }

    return result.valid ? 0 : 1;
}

/**
 * What a partial run skipped beyond its omitted check ids: the rules of the
 * format validator a full run would reach, under a profile that walks
 * schemas the schema walk (#339), and for a DOCX given an original the
 * redlining validation against it (#353). A suffix with no format validator
 * (`.xlsx`) names none.
 */
function partialRunGaps(suffix: string, profile: Profile, hasOriginal: boolean): string {
    const validator = suffix === ".docx" || suffix === ".docm" ? "DOCX" : suffix === ".pptx" ? "PPTX" : undefined;
    if (validator === undefined) return "Also not run: everything outside the selected checks.";
    const gaps = [`the ${validator} validator's own rules`];
    if (profileWalksSchemas(formatForSuffix(suffix), profile)) gaps.push("its schema walk");
    if (validator === "DOCX" && hasOriginal) gaps.push("the redlining validation against the original");
    const last = gaps.pop();
    return `Also not run: ${gaps.length > 0 ? `${gaps.join(", ")} and ${last}` : last}.`;
}

runCli(import.meta.url, () => runValidateFromArgv(process.argv.slice(2)));
