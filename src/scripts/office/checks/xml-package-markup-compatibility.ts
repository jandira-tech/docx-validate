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

import type { CheckDefinition } from "../../../lib/check-registry";
import { isContentTypesStreamName } from "../../../lib/content-types";
import { xmlPartNames } from "./xml";
import { isMarkupCompatibilityJudgedPart } from "./word-markup-compatibility";

/**
 * Complement word-markup-compatibility outside its folder/sidecar scope.
 * The content-types stream has ct-markup-compatibility. Reuse the processor:
 * ancestor scope, opaque extension slots and unselected content stay correct.
 * Keep mce-* codes so existing root-Ignorable deduplication still works.
 */
export const xmlPackageMarkupCompatibilityCheck: CheckDefinition = {
    id: "xml-package-markup-compatibility",
    description: "Other package XML parts have well-formed markup compatibility directives.",
    groups: ["xml"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx, profile) =>
        [...xmlPartNames(ctx.partNames, (part) => ctx.bytes(part))]
            .filter((part) => !isMarkupCompatibilityJudgedPart(part) && !isContentTypesStreamName(part))
            .flatMap((part) =>
                ctx
                    .compatibilityIssues(part)
                    .filter((issue) => issue.severity === "error")
                    .map((issue) => ({
                        ...issue,
                        severity: profile === "strict" ? ("error" as const) : ("warning" as const),
                    })),
            ),
};
