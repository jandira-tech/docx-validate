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
import { EXTENSION_NAMESPACES } from "../../../lib/markup-compatibility";
import type { ValidationIssue } from "../../../lib/types";
import { wordReadStoryParts } from "../../../lib/word-read-parts";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

export const WORD_FIELD_UNBALANCED = "word-field-unbalanced";
const WORD = new Set<string>(WORD_PARAGRAPH_NAMESPACES);
// A field in one note/comment/textbox/docPart cannot close one in another.
const SCOPES = new Set(["footnote", "endnote", "comment", "txbxContent", "docPartBody"]);

/** Pure ordered scan, with independent field stacks for embedded stories. */
export function fieldBalanceProblems(root: Element): string[] {
    const problems: string[] = [];
    type Field = { separated: boolean };
    type Visit = { element: Element; fields: Field[] } | { finish: Field[] };
    const fields: Field[] = [];
    const pending: Visit[] = [{ finish: fields }, { element: root, fields }];
    while (pending.length > 0) {
        const step = pending.pop()!;
        if ("finish" in step) {
            if (step.finish.length > 0) problems.push(`${step.finish.length} begin marker(s) have no end in this story.`);
            continue;
        }
        const { element } = step;
        if (element.localName === "ext" && EXTENSION_NAMESPACES.has(element.namespaceURI ?? "")) continue;
        const isWord = WORD.has(element.namespaceURI ?? "");
        const current: Field[] = isWord && SCOPES.has(element.localName) ? [] : step.fields;
        if (current !== step.fields) pending.push({ finish: current });
        if (isWord && element.localName === "fldChar") {
            const type = element.getAttributeNS(element.namespaceURI!, "fldCharType");
            if (type === "begin") current.push({ separated: false });
            else if (type === "separate") {
                const open = current.at(-1);
                if (open === undefined) problems.push("separate marker has no preceding begin in this story.");
                else if (open.separated) problems.push("A field has more than one separate marker.");
                else open.separated = true;
            } else if (type === "end" && current.pop() === undefined) problems.push("end marker has no preceding begin in this story.");
            // Missing/invalid fldCharType is a schema/attribute concern, not balance.
        }
        for (let child = element.lastChild; child; child = child.previousSibling) {
            if (child.nodeType === 1) pending.push({ element: child as Element, fields: current });
        }
    }
    return problems;
}

export const wordFieldBalanceCheck: CheckDefinition = {
    id: WORD_FIELD_UNBALANCED,
    description: "Complex-field begin, separate and end markers balance within each read story.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        for (const part of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok || !parsed.document.documentElement) continue;
            for (const problem of fieldBalanceProblems(parsed.document.documentElement))
                issues.push({
                    severity: profile === "strict" ? "error" : "warning",
                    code: WORD_FIELD_UNBALANCED,
                    path: part,
                    message: `Unbalanced complex field: ${problem}`,
                });
        }
        return issues;
    },
};
