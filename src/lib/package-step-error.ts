/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import type { ValidationIssue } from "./types";

/** The step of a bytes-in class that can refuse a package. */
export type PackageStep = "unpack" | "pack";

/**
 * A step `Repair` or `Normalize` runs refused the package: `unpack` could not
 * extract it, or `pack`'s validation failed and wrote nothing. Before #1157
 * both classes read the output anyway and threw a bare `ENOENT` naming their
 * temp directory.
 */
export class PackageStepError extends Error {
    public constructor(
        public readonly step: PackageStep,
        message: string,
        /** `pack`'s validation lines, when the pack step refused. */
        public readonly validationLog?: string,
        /** The issues `pack`'s validation failed on, when the pack step refused. */
        public readonly issues: readonly ValidationIssue[] = [],
    ) {
        super(message);
        this.name = "PackageStepError";
    }
}

/**
 * Throw a {@link PackageStepError} when `result` reports a refusal, with the
 * temp directory the step ran in elided from its message. The log and issues
 * pass through: neither names the directory, as validators name parts
 * relative to the package.
 */
export const assertStepOk = (
    owner: string,
    step: PackageStep,
    result: { ok: boolean; message: string; validationLog?: string; issues?: readonly ValidationIssue[] },
    tempDir: string,
): void => {
    if (result.ok) {
        return;
    }
    throw new PackageStepError(
        step,
        `${owner}: ${step} refused the package: ${result.message.replaceAll(tempDir, "<temp>")}`,
        result.validationLog,
        result.issues ?? [],
    );
};
