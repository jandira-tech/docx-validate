/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * `Repair` — bytes-in, repaired-bytes-out top-level class.
 *
 * PR C Task C.2 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §2.
 *
 * Wraps today's repair surface: `BaseSchemaValidator.repair()` +
 * `DOCXSchemaValidator.repair()` (whitespace `xml:space="preserve"` injection,
 * structural fixes, durable-id / para-id / mc:Ignorable repairs).
 *
 * The repair-count return value is what the existing `validate.ts` CLI shim
 * reports today; this class lifts that surface to a typed bytes-in API.
 */

import { writeFile, readFile } from "node:fs/promises";
import path from "node:path";
import type { ValidationIssue } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { DOCXSchemaValidator } from "./scripts/office/validators/docx";
import { unpack } from "./scripts/office/unpack";
import { pack } from "./scripts/office/pack";
import { withTempDir } from "./lib/run-cli";

export type RepairOptions = {
    /**
     * Override the XSD engine used internally (PR B's injection mechanism).
     * Defaults to the legacy libxmljs2 path until the cutover lands.
     */
    xsdValidator?: XsdValidator;
    /** Override the bundled XSD schemas directory. */
    schemasDir?: string;
    /** Verbose mode passed through to the validator. */
    verbose?: boolean;
};

export type RepairResult = {
    bytes: Uint8Array;
    repairs: number;
    diagnostics: ValidationIssue[];
};

export class Repair {
    public constructor(private readonly opts: RepairOptions = {}) {}

    public run(bytes: Uint8Array): Promise<RepairResult> {
        return withTempDir(async (dir) => {
            const inputPath = path.join(dir, "input.docx");
            await writeFile(inputPath, bytes);

            const unpackedDir = path.join(dir, "unpacked");
            await unpack(inputPath, unpackedDir);

            const validator = new DOCXSchemaValidator({
                unpackedDir,
                verbose: this.opts.verbose ?? false,
                schemasDir: this.opts.schemasDir,
                xsdValidator: this.opts.xsdValidator,
            });
            const repairs = await validator.repair();

            const outPath = path.join(dir, "out.docx");
            await pack(unpackedDir, outPath);
            const out = await readFile(outPath);

            return {
                bytes: new Uint8Array(out),
                repairs,
                diagnostics: [],
            };
        });
    }
}
