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
 * Layout-bomb ceilings: the combinatorial shapes that make a renderer
 * explode — an oversized `w:gridSpan` (Word's table grid tops out at 63
 * columns) and an unbounded field count (`w:fldChar`/`w:instrText`/
 * `w:fldSimple`, the fields behind bibliography and cross-reference bombs).
 * Each is a named ceiling from the run's resource ceilings, reported as an
 * error so the package is refused rather than preserved.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordReadStoryParts } from "../../../lib/word-read-parts";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

export const WORD_GRID_SPAN_CEILING = "word-grid-span-ceiling";
export const WORD_FIELD_COUNT_CEILING = "word-field-count-ceiling";

const FIELD_MARKER_ELEMENTS = ["fldChar", "instrText", "fldSimple"] as const;

export const wordLayoutCeilingsCheck: CheckDefinition = {
    id: "word-layout-ceilings",
    description:
        "No w:gridSpan exceeds the grid-span ceiling (63, Word's table column limit) and no story part carries more field markers than the field ceiling.",
    groups: ["story", "tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        for (const part of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) continue;
            const document = parsed.document;
            let overSpan = 0;
            let worstSpan = 0;
            let fieldMarkers = 0;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const gridSpan of getElementsByTagNameNSAll(document, ns, "gridSpan")) {
                    const value = Number(gridSpan.getAttributeNS(ns, "val"));
                    if (Number.isFinite(value) && value > ctx.ceilings.maxGridSpan) {
                        overSpan += 1;
                        worstSpan = Math.max(worstSpan, value);
                    }
                }
                for (const element of FIELD_MARKER_ELEMENTS) {
                    fieldMarkers += getElementsByTagNameNSAll(document, ns, element).length;
                }
            }
            if (overSpan > 0) {
                issues.push({
                    severity: "error",
                    code: WORD_GRID_SPAN_CEILING,
                    path: part,
                    message: `${overSpan} w:gridSpan value(s) in this part exceed the grid-span ceiling of ${ctx.ceilings.maxGridSpan} (maxGridSpan, Word's table column limit); worst is ${worstSpan}.`,
                });
            }
            if (fieldMarkers > ctx.ceilings.maxFieldMarkers) {
                issues.push({
                    severity: "error",
                    code: WORD_FIELD_COUNT_CEILING,
                    path: part,
                    message: `This part carries ${fieldMarkers} field markers (w:fldChar/w:instrText/w:fldSimple), over the field ceiling of ${ctx.ceilings.maxFieldMarkers} (maxFieldMarkers). A bibliography or cross-reference bomb expands combinatorially at layout time.`,
                });
            }
        }
        return issues;
    },
};
