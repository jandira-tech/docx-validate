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
 * `Validate` — bytes-in / `ValidationResult`-out top-level class.
 *
 * PR C Task C.1 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §2.
 *
 * This is the platform-agnostic surface that the `jubarte.validate(bytes)`
 * namespace helper composes (PR C Task C.5).
 *
 * Today the implementation writes the bytes to a tempfile and delegates
 * to the existing path-taking `validate()` function. Once Task C.7 lands
 * the in-memory ZIP core (`packBytes` / `unpackBytes`), this class will
 * skip the disk hop entirely and become fully byte-native.
 *
 * Every option the path entrypoint understands is forwarded here verbatim
 * (plan task 11), so bytes and paths return the same issues and the same
 * execution metadata.
 *
 * The Node-only convenience `validateFile(path)` wrapper lives in
 * `src/node/validate-file.ts` (PR C Task C.6).
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";
import type { Profile, ValidationResult } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { validate as runValidate, type ValidateOptions as PathValidateOptions, type WordVerifier } from "./scripts/office/validate";
import { withTempDir } from "./lib/run-cli";

export type ValidateOptions = {
    /**
     * Override the XSD engine. When omitted, BaseSchemaValidator uses its
     * default wasm-backed validator (the libxmljs2 legacy path was retired
     * with the libxml2-wasm cutover — the native addon cannot build on
     * Node 26).
     */
    xsdValidator?: XsdValidator;
    /**
     * Override the bundled XSD schemas directory. Defaults to the
     * `src/scripts/office/schemas/` tree shipped with the package.
     */
    schemasDir?: string;
    /**
     * Validation profile. Omitted means the per-format default
     * (`"word-valid"` for `.docx`/`.docm`, `"lenient"` otherwise). See
     * {@link "./lib/mode-policy"} for how a request becomes a mode.
     */
    profile?: Profile;
    /** Shorthand for `profile: "strict"`. */
    strict?: boolean;
    /** Run the common-defect check preset and disclose the checks it omits. */
    quick?: boolean;
    /** Run these check ids or group names, plus their prerequisites. */
    checks?: readonly string[];
    /** Also ask Microsoft Word to open the exact bytes. */
    verifyWord?: boolean;
    /** Override the real-Word verifier. Defaults to `verifyInWord`. */
    wordVerifier?: WordVerifier;
    /**
     * File-extension override. Detected from the DOCX magic bytes when
     * omitted; specify when the bytes are known to be a different
     * supported suffix (e.g. `.pptx`, `.docm`).
     */
    suffix?: string;
};

/** The subset of {@link ValidateOptions} the path entrypoint consumes. */
const forwardable = (opts: ValidateOptions): PathValidateOptions => ({
    schemasDir: opts.schemasDir,
    xsdValidator: opts.xsdValidator,
    profile: opts.profile,
    strict: opts.strict,
    quick: opts.quick,
    checks: opts.checks,
    verifyWord: opts.verifyWord,
    wordVerifier: opts.wordVerifier,
});

export class Validate {
    public constructor(private readonly opts: ValidateOptions = {}) {}

    public run(bytes: Uint8Array): Promise<ValidationResult> {
        const options = forwardable(this.opts);
        return withTempDir(async (dir) => {
            const suffix = this.opts.suffix ?? ".docx";
            const docxPath = path.join(dir, `input${suffix}`);
            await writeFile(docxPath, bytes);
            const result = await runValidate(docxPath, options);
            // `xsdWalk` and `repairs` stay on the richer path-entrypoint result;
            // this class reports the shared `ValidationResult`, metadata included.
            return { valid: result.valid, issues: result.issues, metadata: result.metadata };
        });
    }
}
