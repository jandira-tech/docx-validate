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

export const WORD_MOVE_RANGE_UNPAIRED = "word-move-range-unpaired";
const WORD = new Set<string>(WORD_PARAGRAPH_NAMESPACES);
const SCOPES = new Set(["footnote", "endnote", "comment", "txbxContent", "docPartBody"]);
type Scope = {
    active: Map<string, number>;
    names: Map<string, { name: string; from: number; to: number }>;
};
const newScope = (): Scope => ({ active: new Map(), names: new Map() });
// Compare xsd:integer values without lossy Number conversion.
const decimalKey = (value: string): string => (/^[+-]?[0-9]+$/.test(value.trim()) ? BigInt(value.trim()).toString() : value);

/** Pair ranges by kind/id and move counterparts by name, in each story. */
export function moveRangeProblems(root: Element): string[] {
    const problems: string[] = [];
    type Visit = { element: Element; scope: Scope } | { finish: Scope };
    const scope = newScope();
    const pending: Visit[] = [{ finish: scope }, { element: root, scope }];
    while (pending.length > 0) {
        const step = pending.pop()!;
        if ("finish" in step) {
            for (const [key, count] of step.finish.active) {
                if (count > 0)
                    problems.push(`${count} range start(s) ${JSON.stringify(key.split("\u0000").slice(1))} have no end in this story.`);
            }
            for (const { name, from, to } of step.finish.names.values()) {
                if (from !== to)
                    problems.push(`Move ${JSON.stringify(name)} has ${from} moveFrom and ${to} moveTo start(s) in this story.`);
            }
            continue;
        }
        const { element } = step;
        const ns = element.namespaceURI ?? "";
        if (element.localName === "ext" && EXTENSION_NAMESPACES.has(ns)) continue;
        const current = WORD.has(ns) && SCOPES.has(element.localName) ? newScope() : step.scope;
        if (current !== step.scope) pending.push({ finish: current });
        const marker = WORD.has(ns) ? /^(moveFrom|moveTo)Range(Start|End)$/.exec(element.localName) : null;
        if (marker) {
            const kind = marker[1]!;
            const id = element.getAttributeNS(ns, "id");
            if (id !== null) {
                const key = `${ns}\u0000${kind}\u0000${decimalKey(id)}`;
                const count = current.active.get(key) ?? 0;
                if (marker[2] === "Start") current.active.set(key, count + 1);
                else if (count === 0) problems.push(`${kind}RangeEnd id ${JSON.stringify(id)} has no preceding start in this story.`);
                else current.active.set(key, count - 1);
            }
            if (marker[2] === "Start") {
                const name = element.getAttributeNS(ns, "name");
                if (name !== null) {
                    const key = `${ns}\u0000${name}`;
                    const counts = current.names.get(key) ?? { name, from: 0, to: 0 };
                    if (kind === "moveFrom") counts.from += 1;
                    else counts.to += 1;
                    current.names.set(key, counts);
                }
            }
        }
        for (let child = element.lastChild; child; child = child.previousSibling) {
            if (child.nodeType === 1) pending.push({ element: child as Element, scope: current });
        }
    }
    return problems;
}

export const wordMoveRangePairingCheck: CheckDefinition = {
    id: WORD_MOVE_RANGE_UNPAIRED,
    description: "Move ranges pair by kind/id and have corresponding moveFrom/moveTo names in each read story.",
    groups: ["tracked-changes"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        for (const part of await wordReadStoryParts(ctx, profile)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok || !parsed.document.documentElement) continue;
            for (const problem of moveRangeProblems(parsed.document.documentElement))
                issues.push({
                    severity: profile === "strict" ? "error" : "warning",
                    code: WORD_MOVE_RANGE_UNPAIRED,
                    path: part,
                    message: problem,
                });
        }
        return issues;
    },
};
