// cspell:words vbaProject macroEnabled
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
 * Active-content finding. A validator never executes a package, so a VBA
 * macro project cannot run here — but "the weapon is still in the package"
 * is a fact the caller must hear named, not a silent pass. A package
 * carrying a `vbaProject.bin` part (or a macro-enabled content type in its
 * content-types stream) is reported as a warning under every profile: Word
 * opens a macro-enabled document only after a security prompt, so the
 * package is not refused, but the embedded code is named with its part.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { findContentTypesStream } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";

export const WORD_MACRO_PROJECT = "word-macro-project";

/** A stored part that is a VBA macro project, by its leaf name. */
const VBA_PROJECT_PART = /(?:^|\/)vbaProject\.bin$/i;

/** Content-type spellings that mark a macro-enabled package. */
const MACRO_CONTENT_TYPES = [/vbaProject/i, /macroEnabled/i];

/** A `ContentType` attribute of a `Default` or `Override`, its value in group 2. */
const CONTENT_TYPE_ATTRIBUTE = /\bContentType\s*=\s*(["'])(.*?)\1/g;

export const wordMacroProjectCheck: CheckDefinition = {
    id: "word-macro-project",
    description: "Names embedded active content: a vbaProject.bin part or a macro-enabled content type is reported, never executed.",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        const vbaPart = ctx.partNames.find((name) => VBA_PROJECT_PART.test(name));
        if (vbaPart !== undefined) {
            issues.push({
                severity: "warning",
                code: WORD_MACRO_PROJECT,
                path: vbaPart,
                message: `Package embeds a VBA macro project in ${JSON.stringify(vbaPart)}. The validator never executes it; Word opens the document only after a macro security prompt.`,
            });
        }
        const stream = findContentTypesStream(ctx.partNames);
        const text = stream === undefined ? undefined : ctx.text(stream);
        if (text !== undefined) {
            // Only the declared media types count: a PartName, an Extension
            // or a comment that merely spells "macroEnabled" is no macro
            // content type (PR #1281 review).
            const declared = [...text.matchAll(CONTENT_TYPE_ATTRIBUTE)].map((match) => match[2]!);
            const hit = MACRO_CONTENT_TYPES.find((pattern) => declared.some((contentType) => pattern.test(contentType)));
            if (hit !== undefined) {
                issues.push({
                    severity: "warning",
                    code: WORD_MACRO_PROJECT,
                    path: stream,
                    message: `Content-types stream ${JSON.stringify(stream)} declares a macro-enabled package (${hit.source}). The validator never executes it; Word opens the document only after a macro security prompt.`,
                });
            }
        }
        return issues;
    },
};
