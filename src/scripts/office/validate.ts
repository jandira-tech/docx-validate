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
import { promises as fs, type Stats } from "node:fs";
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
    opcNameIssueSeverity,
    promoteFullCheckIssue,
    profileWalksSchemas,
    QUICK_CHECK_GROUPS,
    resolveMode,
    wordObservedIssueSeverity,
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
import {
    archiveStoredNames,
    assertPlaceableNames,
    indexByPartName,
    normalizePartName,
    ValidationContext,
} from "../../lib/validation-context";
import { extractPackageParts } from "../../lib/zip-path";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "./checks";
import { XML_SYNTAX_CHECK_ID } from "./checks/xml";
import { OPC_EQUIVALENT_PART_NAMES } from "./checks/opc-equivalent-part-names";
import { OPC_PART_NAME_WORD_REFUSED } from "./checks/opc-part-name-word-refused";
import { CT_OVERRIDE_PARTNAME_WORD_REFUSED } from "./checks/ct-override-partname-word-refused";
import { OPC_PART_NAME_ESCAPED_SLASH } from "./checks/opc-part-name-escaped-slash";
import { OPC_PART_NAME_ESCAPED_UNRESERVED } from "./checks/opc-part-name-escaped-unreserved";
import { CT_OVERRIDE_PARTNAME_INVALID } from "./checks/ct-override-partname";
import { CT_MARKUP_COMPATIBILITY } from "./checks/ct-markup-compatibility";
import { isMarkupCompatibilityJudgedPart } from "./checks/word-markup-compatibility";
import { WORD_PART_CONTENT_TYPE_GENERIC } from "./checks/word-part-content-type-generic";
import { contextFromArchive, openDirectoryContext, packageExtractFailedIssue, storedNamesRun } from "./package-context";
import {
    BaseSchemaValidator,
    containRelationshipTarget,
    rootIgnorableIssues,
    walkDirectory,
    withoutRestatedIgnorablePrefixes,
} from "./validators/base";
import { DOCXSchemaValidator } from "./validators/docx";
import { buildRepairPlanIssues, collectDocxSemanticInventory, compareDocxSemanticInventories } from "./validators/docx-diagnostics";
import { PPTXSchemaValidator } from "./validators/pptx";
import { validateRedlining } from "./validators/redlining";
import { propertiesParts } from "../../lib/properties-parts";
import { wordReadParts, wordStoryRootParts } from "../../lib/word-read-parts";
import { CT_DUPLICATE_DECLARATION } from "../../lib/content-types";
import { isWordBlockingIssue } from "./validators/word-blocking";
import { explainWordError } from "./validators/word-error-explanations";
import { loadCentralNamedArchive } from "../../lib/zip-central-names";
import { WORD_PART_CONTENT_TYPE_MISMATCH } from "./checks/word-part-content-type-mismatch";
import { WORD_PART_CONTENT_TYPE_INEXACT } from "./checks/word-part-content-type-inexact";

const SUPPORTED_SUFFIXES = new Set([".docm", ".docx", ".pptx", ".xlsx"]);

/**
 * The package-name findings whose severity follows the profile and the
 * format ({@link opcNameIssueSeverity}): equivalent stored names, escaped
 * unreserved characters and separators, duplicate content-type declarations,
 * malformed Override PartNames, and a Word-read part typed only by the
 * generic Default (#432) or carrying an inexact spelling of its Word
 * content type — a MIME parameter or padding (#540, #550) — or a mapped
 * content type belonging to another part (#600).
 */
const ADVISORY_CODES: ReadonlySet<string> = new Set([
    OPC_EQUIVALENT_PART_NAMES,
    OPC_PART_NAME_ESCAPED_UNRESERVED,
    OPC_PART_NAME_ESCAPED_SLASH,
    CT_DUPLICATE_DECLARATION,
    CT_OVERRIDE_PARTNAME_INVALID,
    WORD_PART_CONTENT_TYPE_GENERIC,
    WORD_PART_CONTENT_TYPE_INEXACT,
    WORD_PART_CONTENT_TYPE_MISMATCH,
]);

/**
 * The findings for spec-valid markup Word refuses, whose severity is
 * {@link wordObservedIssueSeverity}: Word-refused part-name spellings (#451,
 * #408, #443) and Markup Compatibility in the content-types stream (#459).
 */
const WORD_OBSERVED_CODES: ReadonlySet<string> = new Set([
    OPC_PART_NAME_WORD_REFUSED,
    CT_OVERRIDE_PARTNAME_WORD_REFUSED,
    CT_MARKUP_COMPATIBILITY,
]);

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

    // Read the packed file's bytes once, before any validation work: a packed
    // input that cannot be read is an archive that does not open — `fs.stat`
    // still answers for a mode-000 file, so the read is the first thing that
    // can fail (#422). A readable one gives Word verification the same bytes
    // the portable mode inspected.
    const packedRead: { bytes: Uint8Array } | { error: unknown } | undefined = targetIsPackedFile
        ? await fs.readFile(target).then(
              (bytes) => ({ bytes }),
              (error: unknown) => ({ error }),
          )
        : undefined;
    const packedBytes = packedRead !== undefined && "bytes" in packedRead ? packedRead.bytes : undefined;
    const finish = async (result: ValidateRunResult): Promise<ValidateRunResult> =>
        opts.verifyWord === true
            ? withWordVerification(
                  result,
                  packedBytes,
                  opts.wordVerifier,
                  packedRead !== undefined && "error" in packedRead
                      ? `Word verification needs the packed file's bytes; ${target} cannot be read: ${packedRead.error instanceof Error ? packedRead.error.message : String(packedRead.error)}`
                      : undefined,
              )
            : result;

    if (packedRead !== undefined && "error" in packedRead) {
        // An unreadable input is an archive that cannot be opened: the same
        // finding in every mode, and no check ran (#422).
        return finish({
            valid: false,
            issues: [
                packageOpenFailedIssue(
                    target,
                    packedRead.error instanceof Error ? packedRead.error.message : String(packedRead.error),
                    "read",
                ),
            ],
            repairs: 0,
            suffix: dispatchSuffix,
            xsdWalk: xsdWalkNotStarted(),
            metadata: { profile, executedChecks: [], omittedChecks: [], schemaCompleted: false },
        });
    }

    if (selected !== undefined) {
        return finish(await runSelectedChecks(target, packedBytes, dispatchSuffix, format, mode, registry, selected));
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

    // The package view could not be staged (a ZIP that failed extraction, or
    // a directory file that could not be read into it): a structured issue,
    // never a throw — `validate()` must NEVER throw here.
    //
    // `extra` holds the findings the archive still offers when the staging
    // failed after the archive opened (#454): they are merged into the
    // failure, so a partial run of the same package reports nothing the full
    // run does not (#381).
    const stagingFailed = (
        envelope: { issues: ValidationResult["issues"]; executed: readonly string[] },
        err: unknown,
        extra: { issues: readonly ValidationIssue[]; executed: readonly string[] } = EMPTY_FINDINGS,
    ): ValidateRunResult => {
        const detail = err instanceof Error ? err.message : String(err);
        return {
            valid: false,
            issues: [...envelope.issues, ...extra.issues, packageExtractFailedIssue(target, detail)],
            repairs: 0,
            suffix: dispatchSuffix,
            xsdWalk: xsdWalkNotStarted(),
            // The envelope was read from the bytes before extraction failed,
            // so its checks ran (CodeRabbit on #335).
            metadata: {
                profile,
                executedChecks: orderedExecutedChecks(format, [...envelope.executed, ...extra.executed]),
                omittedChecks: [],
                schemaCompleted: false,
            },
        };
    };

    // A packed input, or a directory holding a stored name that differs from
    // its part name (`word\document.xml`) or a symbolic link — a stored entry
    // or a directory segment — is judged through the package view staged
    // under part names, so the path-keyed validators see one entry per part
    // (#371, #374) and `--auto-repair` reaches its sources only through the
    // guarded write-back, never a link's target (#425). `sources` maps each
    // staged part to the directory file it came from, so `--auto-repair`
    // keeps writing to the stored names.
    const runStaged = async (
        zip: JSZip,
        envelope: { issues: ValidationResult["issues"]; executed: readonly string[] },
        sources?: ReadonlyMap<string, string>,
        // The findings to merge into a staging failure: what the archive
        // still holds after the extraction failed (#454).
        onExtractFailure?: (err: unknown) => Promise<{ issues: readonly ValidationIssue[]; executed: readonly string[] }>,
        // The input directory a staged directory's write-back may modify;
        // its destinations are guarded against symlinks and hard links
        // (#425, #545).
        writeBackRoot?: string,
    ): Promise<ValidateRunResult> =>
        withTempDir(async (tempDir) => {
            try {
                await extractPackageParts(zip, tempDir);
            } catch (err) {
                return stagingFailed(envelope, err, onExtractFailure ? await onExtractFailure(err) : undefined);
            }
            // Write-back findings; a staged-directory repair the source cannot
            // absorb is an error — the caller asked for a repair it did not
            // get — while the rest of the write-back keeps its best-effort
            // discipline (the validators swallow their own write failures).
            const writeBackIssues: ValidationIssue[] = [];
            const writeBackFailed = (part: string, file: string, err: unknown): void => {
                writeBackIssues.push({
                    severity: "error",
                    message: `Cannot write the auto-repaired ${JSON.stringify(part)} back to ${file}: ${err instanceof Error ? err.message : String(err)}`,
                    path: part,
                    code: "repair-write-back-failed",
                });
            };
            const before = new Map<string, Buffer | null>();
            if (autoRepair && sources) {
                for (const [part, file] of sources) {
                    try {
                        before.set(part, await fs.readFile(path.join(tempDir, part)));
                    } catch (err) {
                        // Without the pre-repair bytes a byte comparison cannot
                        // prove what changed, so this part's write-back is skipped.
                        before.set(part, null);
                        writeBackFailed(part, file, err);
                    }
                }
            }
            // Inspect every source before writing any of them. A symlinked
            // directory segment makes two stored paths name one physical file;
            // that physical file and all of its aliases are refused together,
            // so the link's target is untouched and no inode is written twice
            // (#544).
            const writeBackRefusals = new Map<string, Error>();
            const writeBackIdentities = new Map<string, string>();
            if (autoRepair && sources && writeBackRoot !== undefined) {
                const symlinkedPhysicalFiles = new Set<string>();
                for (const [part, file] of sources) {
                    const inspection = await inspectWriteBackDestination(writeBackRoot, file);
                    if (inspection.refusal !== undefined) writeBackRefusals.set(part, inspection.refusal);
                    if (inspection.identity === undefined) continue;
                    writeBackIdentities.set(part, inspection.identity);
                    if (inspection.throughSymbolicLink) symlinkedPhysicalFiles.add(inspection.identity);
                }
                for (const [part, file] of sources) {
                    const identity = writeBackIdentities.get(part);
                    if (identity === undefined || !symlinkedPhysicalFiles.has(identity) || writeBackRefusals.has(part)) continue;
                    writeBackRefusals.set(
                        part,
                        new Error(
                            `${file} is the same physical file as a symbolic-link path; the write-back never modifies a symbolic link's target`,
                        ),
                    );
                }
            }
            const unpacked = await runWithUnpacked(tempDir, envelope);
            // A staged directory still repairs its source, under stored names.
            // Bytes are compared so untouched files are never rewritten, and a
            // dropped duplicate entry has no write-back destination at all.
            if (autoRepair && sources) {
                const attemptedPhysicalFiles = new Set<string>();
                const claimPhysicalFile = (part: string, file: string): Error | undefined => {
                    const identity = writeBackIdentities.get(part);
                    if (identity === undefined) return undefined;
                    if (attemptedPhysicalFiles.has(identity)) {
                        return new Error(
                            `${file} is the same physical file as another write-back destination; one physical file is written at most once`,
                        );
                    }
                    attemptedPhysicalFiles.add(identity);
                    return undefined;
                };
                for (const [part, file] of sources) {
                    // #425, #544: a destination the input directory does not
                    // directly hold — a symbolic link at any path segment, or
                    // a path that resolves outside it — receives nothing.
                    const refusal = writeBackRefusals.get(part);
                    let repaired: Buffer;
                    try {
                        repaired = await fs.readFile(path.join(tempDir, part));
                    } catch (err) {
                        if ((err as NodeJS.ErrnoException).code === "ENOENT") {
                            if (refusal !== undefined) {
                                writeBackFailed(part, file, refusal);
                                continue;
                            }
                            // The repair removed the part: its source goes too.
                            // A source already removed mid-run is the desired
                            // end state, so only a different rm failure reports.
                            const duplicate = claimPhysicalFile(part, file);
                            if (duplicate !== undefined) {
                                writeBackFailed(part, file, duplicate);
                                continue;
                            }
                            try {
                                await fs.rm(file);
                            } catch (rmErr) {
                                if ((rmErr as NodeJS.ErrnoException).code !== "ENOENT") writeBackFailed(part, file, rmErr);
                            }
                        } else {
                            writeBackFailed(part, file, err);
                        }
                        continue;
                    }
                    const previous = before.get(part);
                    if (previous == null) continue;
                    try {
                        if (!repaired.equals(previous)) {
                            if (refusal !== undefined) writeBackFailed(part, file, refusal);
                            else {
                                const duplicate = claimPhysicalFile(part, file);
                                if (duplicate !== undefined) writeBackFailed(part, file, duplicate);
                                else await fs.writeFile(file, repaired);
                            }
                        }
                    } catch (err) {
                        writeBackFailed(part, file, err);
                    }
                }
            }
            return {
                ...unpacked,
                valid: unpacked.valid && writeBackIssues.every((issue) => issue.severity !== "error"),
                issues: [...unpacked.issues, ...writeBackIssues],
            };
        });

    if (targetIsPackedFile) {
        // Defined: the read above already reported an unreadable input.
        const buf = packedBytes!;
        const envelope = runsWordPackageChecks(profile) ? await zipEnvelopeRun(buf, dispatchSuffix, profile) : EMPTY_ENVELOPE_RUN;
        let zip: JSZip;
        try {
            zip = await loadCentralNamedArchive(buf);
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
        // case/normalization collisions on the host filesystem (#351). The
        // names are the central directory's: the loaded view already dropped
        // a duplicate or hid a `..` segment behind the entry it resolves to,
        // so the diagnosis and the refusal judge every stored name (#411).
        const stored = archiveStoredNames(
            buf,
            Object.values(zip.files).filter((entry) => !entry.dir),
        );
        const diagnostic = await storedNamesRun(ValidationContext.fromEntries(stored.map((name) => [name, ""] as const)), format, profile);
        const named = {
            issues: [...envelope.issues, ...diagnostic.issues],
            executed: [...envelope.executed, ...diagnostic.executed],
        };
        try {
            // The refusal runs on the stored names too: a `..` name JSZip
            // resolved onto a later entry names no part either way (#378).
            assertPlaceableNames(stored);
        } catch (err) {
            return finish(stagingFailed(named, err));
        }
        // The extraction can fail after the archive opened: the host
        // filesystem cannot hold two parts under one name (#373). The archive
        // still reads, so the findings a partial run computes on it belong in
        // the staging failure (#454).
        return finish(await runStaged(zip, named, undefined, () => archiveFindings(buf, format, profile, registry)));
    }

    const dir = path.resolve(target);
    const walked = walkDirectory(dir);
    const opened = openDirectoryContext(dir, walked.unreadable);
    if (opened.openError !== undefined) {
        // Refused stored names (#424) and unreadable walk entries (#557)
        // fail before any part is read. The stored-names diagnosis still
        // sees every stored name, readable or not.
        const diagnostic = await storedNamesRun(opened.context, format, profile);
        return finish(stagingFailed(diagnostic, new Error(opened.openError)));
    }
    const diagnostic = await storedNamesRun(opened.context, format, profile);
    const entries = walked.files.map((file) => [path.relative(dir, file).split(path.sep).join("/"), file] as const);
    // A symlink or hard link in the tree stages too (#425, #545): the stored
    // bytes are read into the package view, while every write-back
    // destination is guarded against following a link or modifying an
    // unnamed hard-link alias.
    if (entries.some(([name]) => name !== normalizePartName(name)) || walked.holdsSymlinks || walked.holdsHardLinks) {
        // JSZip is only an in-memory entry collection here: no ZIP is written.
        const zip = new JSZip();
        for (const [name, file] of entries) {
            let bytes: Buffer;
            try {
                bytes = await fs.readFile(file);
            } catch (err) {
                const detail = err instanceof Error ? err.message : String(err);
                // The staged package view cannot be built, so it fails like a
                // failed extraction — the error names the stored file (#412).
                return finish(stagingFailed(diagnostic, new Error(`Cannot read ${name}: ${detail}`, { cause: err })));
            }
            zip.file(name, bytes, { createFolders: false });
        }
        return finish(await runStaged(zip, diagnostic, indexByPartName(entries), undefined, dir));
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
    unavailableDetails = "Word verification needs a packed Office file; the validated target is an unpacked directory",
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
            details: unavailableDetails,
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
function findingsUnderProfile(
    run: CheckRun,
    registry: CheckRegistry,
    format: PackageFormat,
    profile: Profile,
    properties: ReadonlySet<string>,
    wordRead: ReadonlySet<string> = new Set(),
    context?: ValidationContext,
    storyRootParts: ReadonlySet<string> = new Set(),
): ValidationIssue[] {
    // Full DOCX validation diagnoses root Ignorable prefixes before dropping
    // their MC restatements. Apply those same two steps only when the selected
    // Word compatibility check ran; descendant-only findings remain distinct (#499).
    // The check judges the XML parts under word/ and every relationships part,
    // each compared as OPC compares part names (#636, #639), so the pairing
    // covers both, as the full run's validateNamespaces walk does.
    const outcomes = run.outcomes.map((outcome) => {
        if (outcome.id !== "word-markup-compatibility" || !context || !run.executed.includes(outcome.id)) return outcome;
        const roots = context.partNames.filter(isMarkupCompatibilityJudgedPart).flatMap((name) => {
            const parsed = context.xml(name);
            const root = parsed?.ok ? parsed.document.documentElement : undefined;
            return root ? rootIgnorableIssues(root, name) : [];
        });
        return { ...outcome, issues: withoutRestatedIgnorablePrefixes({ valid: run.valid, issues: [...roots, ...outcome.issues] }).issues };
    });
    // Equivalent names, escaped unreserved characters, escaped separators and
    // malformed Override PartNames are Word-blocking for DOCX/DOCM (#351, #408,
    // #436, #414), and so are two entries under the identical spelling, for
    // which Word offers recovery too (#411).
    // Lenient keeps the earlier policy and runs no Word rule. No PowerPoint or
    // Excel behaviour has been observed, so those only warn. Duplicate content-
    // type declarations follow the same policy (#431, Word 16.115 session 3).
    const advisory = (issue: ValidationIssue): ValidationIssue =>
        issue.code !== undefined && WORD_OBSERVED_CODES.has(issue.code)
            ? { ...issue, severity: wordObservedIssueSeverity(profile, format) }
            : issue.code !== undefined && ADVISORY_CODES.has(issue.code) && opcNameIssueSeverity(profile, format) === "warning"
              ? { ...issue, severity: "warning" }
              : promoteFullCheckIssue(issue, profile, format);
    if (profile !== "word-valid" || format !== "docx") return outcomes.flatMap((outcome) => outcome.issues).map(advisory);
    return outcomes
        .flatMap((outcome) =>
            registry.get(outcome.id)?.archiveOnly === true
                ? outcome.issues
                : outcome.issues.map((issue): ValidationIssue =>
                      issue.severity === "error" && !isWordBlockingIssue(issue, properties, wordRead, storyRootParts)
                          ? { ...issue, severity: "warning" }
                          : issue,
                  ),
        )
        .map(advisory);
}

/**
 * The findings a partial run of the same format reads from the archive, for
 * a full run whose extraction failed (#454).
 *
 * The archive opened, so its package view is readable in memory even when the
 * host filesystem cannot hold it: the checks a partial run runs read those
 * parts, so their findings belong in the staging failure. Without them the
 * partial run reports errors the full run never computed, and the partial
 * errors are not a subset of the full ones (#381).
 *
 * The machinery is the one a partial run uses: the registry's checks for the
 * format, the profile's policy on their findings, and lenient's stop at the
 * first malformed part.
 */
async function archiveFindings(
    bytes: Uint8Array,
    format: PackageFormat,
    profile: Profile,
    registry: CheckRegistry,
): Promise<{ issues: readonly ValidationIssue[]; executed: readonly string[] }> {
    const opened = await contextFromArchive(bytes);
    // An archive that does not open, or whose package view is refused, holds
    // no part: the full run reports that failure alone (#413).
    if (opened.openError !== undefined) return EMPTY_FINDINGS;
    const select = stagingFallbackSelection(registry, format, profile);
    if (select.length === 0) return EMPTY_FINDINGS;
    try {
        const run = await runChecks(opened.context, registry, {
            format,
            select,
            profile,
            // A full lenient run stops at its first malformed part, so the
            // merged run stops where xml-syntax fails (#391).
            ...(profile === "lenient" ? { haltAfterFailing: XML_SYNTAX_CHECK_ID } : {}),
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
        // The full run's reachability sets, so an XML error keeps the
        // severity the full run gives it (#437, #488).
        const [properties, read, storyRoots] =
            profile === "word-valid" && format === "docx"
                ? [
                      await propertiesParts(opened.context, profile),
                      await wordReadParts(opened.context, profile),
                      await wordStoryRootParts(opened.context, profile),
                  ]
                : [new Set<string>(), new Set<string>(), new Set<string>()];
        const issues = findingsUnderProfile(
            { ...run, outcomes, issues: outcomes.flatMap((outcome) => outcome.issues) },
            registry,
            format,
            profile,
            properties,
            read,
            opened.context,
            storyRoots,
        );
        return { issues, executed: [...run.executed] };
    } catch {
        // The staging failure is the verdict. A merged run that cannot start
        // adds nothing to it, and never throws out of `validate()`.
        return EMPTY_FINDINGS;
    }
}

/**
 * The checks a staging failure still runs: every check registered for the
 * format, minus the ones the full run already ran — the ZIP-envelope checks
 * and the stored-name checks — and the ones the profile never runs, so
 * nothing is reported twice.
 */
function stagingFallbackSelection(registry: CheckRegistry, format: PackageFormat, profile: Profile): readonly string[] {
    const outside = checksOutsideProfile(registry, profile);
    return registry
        .list(format)
        .filter((definition) => !outside.has(definition.id))
        .filter((definition) => definition.archiveOnly !== true)
        .filter((definition) => definition.storedNamesOnly !== true)
        .map((definition) => definition.id);
}

/**
 * A partial run: only the selected named checks, against a context built from
 * the package. No schema validator is constructed, so nothing compiles an XSD
 * and no rule outside the selection runs. Everything the selection left out is
 * reported as an omission — `valid` here means "the requested checks passed".
 */
async function runSelectedChecks(
    target: string,
    packedBytes: Uint8Array | undefined,
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
    const opened = packedBytes !== undefined ? await contextFromArchive(packedBytes) : openDirectoryContext(path.resolve(target));
    if (opened.openError !== undefined) {
        // An archive that does not open has no parts: only the selected
        // envelope checks can read its bytes, and nothing else ran (#332).
        // One that opens but whose package view cannot be read fails as the
        // full run's extraction does, not as an open failure (#413) — but its
        // context still carries the stored entry names, so the checks that
        // need only them run and report what the full run's storedNamesRun
        // reports (#420).
        const runnable = new Set(
            registry
                .list(format)
                .filter(
                    (definition) =>
                        definition.archiveOnly === true || (opened.extractFailed === true && definition.storedNamesOnly === true),
                )
                .map((definition) => definition.id),
        );
        const resolved = registry.resolve(select, format);
        const runnableSelect = resolved.filter((id) => runnable.has(id));
        const run =
            runnableSelect.length > 0
                ? await runChecks(opened.context, registry, { format, select: runnableSelect, profile: mode.profile })
                : undefined;
        // A refused view has no parts, so no properties part to block (#409).
        const issues = run === undefined ? [] : findingsUnderProfile(run, registry, format, mode.profile, new Set());
        const executed = run?.executed ?? [];
        return {
            valid: false,
            issues: [
                ...issues,
                opened.extractFailed === true
                    ? packageExtractFailedIssue(target, opened.openError)
                    : packageOpenFailedIssue(target, opened.openError),
            ],
            repairs: 0,
            suffix,
            xsdWalk: xsdWalkNotStarted(),
            metadata: {
                profile: mode.profile,
                executedChecks: [...executed],
                omittedChecks: registry
                    .list(format)
                    .map((definition) => definition.id)
                    // An archive that does not open keeps "omitted means not
                    // selected" (Jules on #337). A refused package view
                    // omits everything that did not run, so no selected id is
                    // in neither list (#420).
                    .filter((id) => (opened.extractFailed === true ? !executed.includes(id) : !resolved.includes(id))),
                schemaCompleted: false,
            },
        };
    }
    // A full lenient run stops at its first malformed part, so a lenient
    // partial run stops where xml-syntax fails (#391).
    const run = await runChecks(opened.context, registry, {
        format,
        select,
        profile: mode.profile,
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
    // #437: the word-valid reachability sets are computed once per partial
    // run, as the full DOCX run computes them, so both agree.
    const [properties, read, storyRoots] =
        mode.profile === "word-valid" && format === "docx"
            ? [
                  await propertiesParts(opened.context, mode.profile),
                  await wordReadParts(opened.context, mode.profile),
                  await wordStoryRootParts(opened.context, mode.profile),
              ]
            : [new Set<string>(), new Set<string>(), new Set<string>()];
    const issues = findingsUnderProfile(
        { ...run, outcomes, issues: outcomes.flatMap((outcome) => outcome.issues) },
        registry,
        format,
        mode.profile,
        properties,
        read,
        opened.context,
        storyRoots,
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

interface WriteBackDestinationInspection {
    /** The physical file's device/inode identity, when it can be read. */
    readonly identity: string | undefined;
    /** Why this destination cannot receive a write-back, when it cannot. */
    readonly refusal: Error | undefined;
    /** Whether the stored path reaches the file through a symbolic link. */
    readonly throughSymbolicLink: boolean;
}

/**
 * Inspect one staged auto-repair write-back destination (#425, #544, #545).
 * The stored file itself must be regular, no path segment below the input
 * root may be a symbolic link, its link count must be one, and its resolved
 * path must stay under the root. The physical identity lets the caller
 * refuse every alias of one file together. A destination that no longer
 * exists keeps the write-back's own ENOENT handling.
 */
async function inspectWriteBackDestination(root: string, file: string): Promise<WriteBackDestinationInspection> {
    let listed: Stats;
    try {
        listed = await fs.lstat(file);
    } catch {
        return { identity: undefined, refusal: undefined, throughSymbolicLink: false };
    }

    let followed: Stats | undefined;
    try {
        followed = await fs.stat(file);
    } catch {
        // lstat succeeded, so the entry exists; a failed stat cannot identify
        // its physical file, but the entry-kind and path checks still apply.
    }
    const identity = followed === undefined ? undefined : `${followed.dev}:${followed.ino}`;

    if (listed.isSymbolicLink()) {
        return {
            identity,
            refusal: new Error(`${file} is a symbolic link; the write-back never follows it`),
            throughSymbolicLink: true,
        };
    }
    if (!listed.isFile()) {
        return { identity, refusal: new Error(`${file} is not a regular file`), throughSymbolicLink: false };
    }

    const linkSegment = await symbolicLinkPathSegment(root, file);
    if (linkSegment !== undefined) {
        return {
            identity,
            refusal: new Error(`${file} passes through symbolic-link path segment ${linkSegment}; the write-back never follows it`),
            throughSymbolicLink: true,
        };
    }
    if (followed !== undefined && followed.nlink > 1) {
        return {
            identity,
            refusal: new Error(`${file} has link count ${followed.nlink}; the write-back refuses a hard-linked file`),
            throughSymbolicLink: false,
        };
    }
    if (containRelationshipTarget(root, file) === null) {
        return { identity, refusal: new Error(`${file} resolves outside the input directory`), throughSymbolicLink: false };
    }
    return { identity, refusal: undefined, throughSymbolicLink: false };
}

/**
 * The first symbolic-link path segment below `root` on the way to `file`, or
 * `undefined` when the path has none. The root itself is outside this
 * check: only segments between the input root and the stored destination
 * are write-back paths the validator must not traverse.
 */
async function symbolicLinkPathSegment(root: string, file: string): Promise<string | undefined> {
    const rootPath = path.resolve(root);
    const relative = path.relative(rootPath, path.resolve(file));
    if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) return undefined;

    const segments = relative.split(path.sep);
    segments.pop();
    let current = rootPath;
    for (const segment of segments) {
        current = path.join(current, segment);
        try {
            const listed = await fs.lstat(current);
            if (listed.isSymbolicLink()) return current;
        } catch {
            return undefined;
        }
    }
    return undefined;
}

/**
 * The finding for a packed file that is not a readable OPC (zip) package:
 * encrypted OOXML, or a truncated or corrupted archive. Word refuses to open
 * all of these, so it is blocking under every profile. A file the host cannot
 * read at all (#422) names the system error instead, since nothing about the
 * archive is known (#547).
 */
function packageOpenFailedIssue(target: string, detail: string, failure: "read" | "archive" = "archive"): ValidationIssue {
    return {
        severity: "error",
        message:
            failure === "read"
                ? `Cannot read ${target}: ${detail}`
                : `Cannot open ${target} as an OPC package (not a valid .docx/zip — encrypted, corrupted, or truncated): ${detail}`,
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
    profile: Profile,
): Promise<{ issues: ValidationResult["issues"]; executed: readonly string[] }> {
    const format: PackageFormat = suffix === ".pptx" ? "pptx" : suffix === ".xlsx" ? "xlsx" : "docx";
    const context = ValidationContext.fromEntries({}, { source: "archive", archiveBytes });
    const registry = createDefaultCheckRegistry();
    // Every archive-only package check; the stored-name diagnosis runs separately.
    const select = registry
        .list()
        .filter((definition) => definition.archiveOnly === true)
        .map((definition) => definition.id);
    const run = await runChecks(context, registry, { format, select, profile });
    return { issues: [...run.issues], executed: [...run.executed] };
}

/** An envelope run that never happened: no findings, no executed checks. */
const EMPTY_ENVELOPE_RUN: { issues: ValidationResult["issues"]; executed: readonly string[] } = { issues: [], executed: [] };

/** Findings a staging failure has none of: the archive offered nothing more. */
const EMPTY_FINDINGS: { issues: readonly ValidationIssue[]; executed: readonly string[] } = { issues: [], executed: [] };

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
    let beforeIssues: readonly ValidationIssue[] = [];
    let contentIssues: ValidationIssue[] = [];
    if (opts.autoRepair) {
        const beforeInventory =
            opts.suffix === ".docx" || opts.suffix === ".docm" ? await collectDocxSemanticInventory(unpackedDir, opts.profile) : null;
        const beforeResults = await Promise.all(validators.map((v) => v.validate()));
        beforeIssues = mergeResults(...beforeResults).issues;
        for (const v of validators) {
            repairs += await v.repair();
        }
        contentIssues =
            beforeInventory && repairs > 0
                ? compareDocxSemanticInventories(beforeInventory, await collectDocxSemanticInventory(unpackedDir, opts.profile))
                : [];
    }

    const results = await Promise.all(validators.map((v) => v.validate()));
    const after = mergeResults(...results);
    const repairPlanIssues = opts.autoRepair ? buildRepairPlanIssues(beforeIssues, after.issues) : [];
    const merged = mergeResults(after, {
        valid: contentIssues.every((i) => i.severity !== "error"),
        issues: [...repairPlanIssues, ...contentIssues],
    });
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
    } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code === "ENOENT" || code === "ENOTDIR") throw new Error(message);
        const detail = err instanceof Error ? err.message : String(err);
        throw new Error(`Error: Cannot access ${p}: ${detail}`, { cause: err });
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
