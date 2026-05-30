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
 * `Validate` — bytes-in / `ValidationResult`-out top-level class.
 *
 * PR C Task C.1 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §2.
 *
 * This is the platform-agnostic surface that the eventual
 * `jubarte.validate(bytes)` namespace helper composes (PR C Task C.5).
 *
 * Today the implementation writes the bytes to a tempfile and delegates
 * to the existing path-taking `validate()` function. Once Task C.7 lands
 * the in-memory ZIP core (`packBytes` / `unpackBytes`), this class will
 * skip the disk hop entirely and become fully byte-native.
 *
 * The Node-only convenience `validateFile(path)` wrapper lives in
 * `src/node/validate-file.ts` (PR C Task C.6).
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";
import type { ValidationResult } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { validate as runValidate } from "./scripts/office/validate";
import { withTempDir } from "./lib/run-cli";

export type ValidateOptions = {
    /**
     * Override the XSD engine. When omitted, BaseSchemaValidator uses its
     * legacy libxmljs2 path (the dual-path injection mechanism landed in
     * PR B). PR C wires this option end-to-end so consumers can opt into
     * the wasm validator without modifying BaseSchemaValidator construction.
     */
    xsdValidator?: XsdValidator;
    /**
     * Override the bundled XSD schemas directory. Defaults to the
     * `src/scripts/office/schemas/` tree shipped with the package.
     */
    schemasDir?: string;
    /**
     * `"lenient"` (default — matches real-world Microsoft Office output)
     * or `"strict"` (spec-purist). See {@link "./lib/types"} for the union.
     */
    profile?: "lenient" | "strict" | "word-valid";
    /**
     * File-extension override. Detected from the DOCX magic bytes when
     * omitted; specify when the bytes are known to be a different
     * supported suffix (e.g. `.pptx`, `.docm`).
     */
    suffix?: string;
};

export class Validate {
    public constructor(private readonly opts: ValidateOptions = {}) {}

    public run(bytes: Uint8Array): Promise<ValidationResult> {
        return withTempDir(async (dir) => {
            const suffix = this.opts.suffix ?? ".docx";
            const docxPath = path.join(dir, `input${suffix}`);
            await writeFile(docxPath, bytes);
            const result = await runValidate(docxPath, {
                schemasDir: this.opts.schemasDir,
                profile: this.opts.profile,
            });
            return { valid: result.valid, issues: result.issues };
        });
    }
}
