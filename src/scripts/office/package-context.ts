/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Builds a {@link ValidationContext} over an unpacked package directory. */

import { readFileSync } from "node:fs";
import path from "node:path";

import { ValidationContext } from "../../lib/validation-context";
import { walkFiles } from "./validators/base";

/**
 * A context over an unpacked package directory. Parts are read lazily and at
 * most once; archive-only checks report unavailable, since a directory has no
 * ZIP envelope.
 */
export function contextFromDirectory(dir: string): ValidationContext {
    const entries = walkFiles(dir).map((file) => [path.relative(dir, file).split(path.sep).join("/"), () => readFileSync(file)] as const);
    return ValidationContext.fromEntries(entries, { source: "directory" });
}

/**
 * A context over packed archive bytes. When the bytes are not a readable ZIP
 * (encrypted OOXML, truncated or corrupted archive) the context still carries
 * them, with no parts, so the ZIP-envelope checks can explain the damage; the
 * reason the archive could not be opened is returned beside it.
 */
export async function contextFromArchive(bytes: Uint8Array): Promise<{ context: ValidationContext; openError?: string }> {
    try {
        return { context: await ValidationContext.fromArchive(bytes) };
    } catch (err) {
        return {
            context: ValidationContext.fromEntries({}, { source: "archive", archiveBytes: bytes }),
            openError: err instanceof Error ? err.message : String(err),
        };
    }
}
