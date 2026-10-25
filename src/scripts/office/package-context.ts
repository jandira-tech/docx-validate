/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Builds a {@link ValidationContext} over an unpacked package directory. */

import { readFileSync } from "node:fs";
import path from "node:path";

import type { PackageFormat, Profile, ValidationIssue } from "../../lib/types";
import { runChecks } from "../../lib/check-scheduler";
import { opcNameIssueSeverity, wordObservedIssueSeverity } from "../../lib/mode-policy";
import { assertPlaceableNames, ValidationContext, PackageViewError, PackageOpenError } from "../../lib/validation-context";
import { createDefaultCheckRegistry } from "./checks";
import { OPC_PART_NAME_WORD_REFUSED } from "./checks/opc-part-name-word-refused";
import { walkDirectory, walkFiles, type UnreadableEntry } from "./validators/base";

/** Shared stored-name diagnosis before any XML/extraction gate (#572). */
export async function storedNamesRun(
    context: ValidationContext,
    format: PackageFormat,
    profile: Profile,
): Promise<{ issues: ValidationIssue[]; executed: readonly string[] }> {
    const registry = createDefaultCheckRegistry();
    const select = registry
        .list(format)
        .filter((definition) => definition.storedNamesOnly === true)
        .map((definition) => definition.id);
    const run = await runChecks(context, registry, { format, select, profile });
    return {
        issues: run.issues.map((issue) => ({
            ...issue,
            severity:
                issue.code === OPC_PART_NAME_WORD_REFUSED
                    ? wordObservedIssueSeverity(profile, format)
                    : opcNameIssueSeverity(profile, format),
        })),
        executed: run.executed,
    };
}

/**
 * The finding for an input whose package view cannot be read: an item name
 * places no part, an entry cannot be read, the host filesystem cannot hold
 * two parts' names, or a staged directory's stored file cannot be read
 * (#412). The package view fails before the format validator runs, so
 * partial runs of a packed file report the same failure (#413). Directly
 * constructed DOCX and PPTX validators report a refused directory view with
 * it too, so they match `validate()` (#548).
 */
export function packageExtractFailedIssue(target: string, detail: string): ValidationIssue {
    return {
        severity: "error",
        message: `Cannot extract ${target}: ${detail}`,
        code: "package-extract-failed",
    };
}

/**
 * A context over an unpacked package directory. Parts are read lazily and at
 * most once; archive-only checks report unavailable, since a directory has no
 * ZIP envelope.
 *
 * `unreadableStoredEntries` (#426) names stored entries that can never yield
 * part bytes — a broken symbolic link, a FIFO. They stay in the context as
 * names without content, so the stored-names checks still see every name;
 * the run that could not read them fails before any check reads a part. A
 * directory is no stored file entry (#573): it contributes no name to the
 * context, so the stored-name checks never judge a name no package holds.
 */
export function contextFromDirectory(dir: string, unreadableStoredEntries: readonly UnreadableEntry[] = []): ValidationContext {
    const entries = walkFiles(dir).map((file) => [path.relative(dir, file).split(path.sep).join("/"), () => readFileSync(file)] as const);
    const unreadable = unreadableStoredEntries.flatMap((entry) => {
        // An unreadable directory names every stored file under it, but the
        // input directory itself names no stored entry (#543): it stays out
        // of the context, while the run still fails on it. A directory is no
        // stored file entry at all (#573), so it names nothing here.
        if (entry.kind === "directory") return [];
        const name = path.relative(dir, entry.file).split(path.sep).join("/");
        return name === "" ? [] : ([[name, "" as const]] as const);
    });
    // Every stored name, readable or not, is judged for placeability (#424).
    const storedNames = [...entries, ...unreadable].map(([name]) => name);
    try {
        assertPlaceableNames(storedNames);
    } catch (err) {
        throw new PackageViewError(err instanceof Error ? err.message : String(err), { cause: err, storedNames });
    }
    return ValidationContext.fromEntries([...entries, ...unreadable], { source: "directory" });
}

/**
 * Open a directory with the same refused-view result as an archive (#424):
 * a stored name the package view refuses gives a part-free context that keeps
 * every stored name, with `extractFailed`. The same walk refuses unreadable
 * entries before any check, for full, partial and direct validators (#557).
 * Stored-name refusals take precedence over walk refusals (#424).
 */
export function openDirectoryContext(
    dir: string,
    unreadableStoredEntries: readonly UnreadableEntry[] = walkDirectory(dir).unreadable,
): {
    context: ValidationContext;
    openError?: string;
    extractFailed?: boolean;
} {
    try {
        const context = contextFromDirectory(dir, unreadableStoredEntries);
        if (unreadableStoredEntries.length === 0) return { context };
        return { context, openError: unreadableStoredDetail(dir, unreadableStoredEntries), extractFailed: true };
    } catch (err) {
        if (!(err instanceof PackageViewError)) throw err;
        return {
            context: ValidationContext.fromStoredNames(err.storedNames, { source: "directory" }),
            openError: err.message,
            extractFailed: true,
        };
    }
}

/**
 * Why a directory input's package view cannot be read: one clause per
 * stored entry that can never yield part bytes, each naming the stored file
 * (#426) — or per directory the walk cannot list, naming the directory
 * (#543), with `.` standing for the input directory itself.
 */
function unreadableStoredDetail(dir: string, unreadable: readonly UnreadableEntry[]): string {
    return unreadable
        .map((entry) => {
            const name = path.relative(dir, entry.file).split(path.sep).join("/");
            return `Cannot read ${name === "" ? "." : name}: ${entry.reason}`;
        })
        .join("; ");
}

/**
 * A context over packed archive bytes. When the bytes are not a readable ZIP
 * (encrypted OOXML, truncated or corrupted archive) the context still carries
 * them, with no parts, so the ZIP-envelope checks can explain the damage; the
 * reason the archive could not be opened is returned beside it. An archive
 * that opens but whose package view cannot be read ({@link PackageViewError}
 * — a refused item name or an unreadable entry) is reported with
 * `extractFailed`, since extraction fails it for the same reason (#413); its
 * context carries no parts but keeps the stored entry names (#420). Any other
 * error is internal — a bug, not a defect in the file — and propagates
 * unwrapped rather than wearing a file-defect code (#427).
 */
export async function contextFromArchive(
    bytes: Uint8Array,
): Promise<{ context: ValidationContext; openError?: string; extractFailed?: boolean }> {
    try {
        return { context: await ValidationContext.fromArchive(bytes) };
    } catch (err) {
        // Only the two classified failures are the input's: the archive that
        // did not open and the package view extraction refuses. Anything
        // else — a TypeError out of our own post-load code — is an internal
        // error and propagates, never a file-defect finding (#427).
        if (!(err instanceof PackageOpenError) && !(err instanceof PackageViewError)) {
            throw err;
        }
        return {
            // A refused package view still carries the stored entry names, so the
            // checks that need only them can run on it (#420).
            context: ValidationContext.fromStoredNames(err instanceof PackageViewError ? err.storedNames : [], {
                source: "archive",
                archiveBytes: bytes,
            }),
            openError: err.message,
            extractFailed: err instanceof PackageViewError || undefined,
        };
    }
}
