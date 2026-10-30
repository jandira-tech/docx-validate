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
import { CONTENT_TYPES_NAMESPACE, findContentTypesStream, findEquivalentPart, partNameViolation } from "../../../lib/content-types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

export const CT_STALE_OVERRIDE = "ct-stale-override";

/**
 * An `Override` whose well-formed PartName names no stored part (#1212). A
 * warning under every profile: Word 16.115 opened such an Override clean, the
 * plain `/word/nothere.xml` with `invalid-x` and with `application/xml` types
 * (docs/plans/evidence/word-probes-2026-10-24.md, sessions 3c, 4, 5 and 7,
 * #414). Names that are not part names stay `ct-override-partname-invalid`'s.
 */
export const ctStaleOverrideCheck: CheckDefinition = {
    id: CT_STALE_OVERRIDE,
    description: "Content-type Overrides name stored parts.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx) => {
        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) return [];
        const parsed = ctx.effectiveXml(stream);
        if (!parsed?.ok) return [];
        return getElementsByTagNameNSAll(parsed.document, CONTENT_TYPES_NAMESPACE, "Override").flatMap((element) => {
            const name = element.getAttribute("PartName");
            // Malformed or absent names already have ct-override-partname-invalid.
            if (name === null || partNameViolation(name) !== undefined || findEquivalentPart(ctx.partNames, name) !== undefined) return [];
            return [
                {
                    severity: "warning" as const,
                    code: CT_STALE_OVERRIDE,
                    path: stream,
                    message: `Override PartName ${JSON.stringify(name)} names no stored part; Word opened such an Override clean (sessions 3c to 7, #414).`,
                },
            ];
        });
    },
};
