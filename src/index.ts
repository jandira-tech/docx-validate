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
 * Public surface for `docx-validate` (github.com/jandira-tech/docx-validate).
 *
 * Three usage paths are supported:
 *
 *   1. Programmatic — `import { validate } from "docx-validate"` and call it
 *      directly. Returns a `ValidateRunResult` with `valid: boolean` and a
 *      structured `issues[]` list. This is the primary API.
 *   2. Validator classes — `DOCXSchemaValidator`, `PPTXSchemaValidator`,
 *      `BaseSchemaValidator` for callers that want to compose individual
 *      checks (e.g. only XSD validation, only redlining, custom pipelines).
 *   3. CLI — `bunx tsx src/scripts/office/validate.ts <path>` for one-off
 *      shell invocations. The CLI just turns argv into options and exits
 *      with the right code; the rest of the surface is identical.
 */

export * from "./lib/types.ts";
export * from "./lib/xml-helpers.ts";
export * from "./lib/run-cli.ts";

// Programmatic validate() entry point and shapes.
export { validate, buildCommand as buildValidateCommand, runFromArgv as runValidateFromArgv } from "./scripts/office/validate.ts";
export type { ValidateOptions, ValidateRunResult } from "./scripts/office/validate.ts";

// Validator classes for callers that want to compose individual checks.
export { BaseSchemaValidator, defaultSchemasDir } from "./scripts/office/validators/base.ts";
export type { BaseSchemaValidatorOptions } from "./scripts/office/validators/base.ts";
export {
    DOCXSchemaValidator,
    WORD_2006_NAMESPACE,
    WORD_STRICT_NAMESPACE,
    WORD_PARAGRAPH_NAMESPACES,
} from "./scripts/office/validators/docx.ts";
export { PPTXSchemaValidator, PRESENTATIONML_NAMESPACE } from "./scripts/office/validators/pptx.ts";
export { validateRedlining, RedliningValidator } from "./scripts/office/validators/redlining.ts";
