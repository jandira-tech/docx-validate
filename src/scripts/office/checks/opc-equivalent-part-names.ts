/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName, findContentTypesStream, findEquivalentPart, isContentTypesStreamName } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import { indexByPartName, normalizePartName } from "../../../lib/validation-context";

export const OPC_EQUIVALENT_PART_NAMES = "opc-equivalent-part-names";

/**
 * Diagnose aliases before the package view drops them. No XML or content is
 * read. The stream is not an OPC part: percent-encoded brackets must never
 * merge with its literal spelling. Profile policy makes this spec diagnosis
 * advisory outside strict until Word's behaviour has been observed.
 */
export const opcEquivalentPartNamesCheck: CheckDefinition = {
    id: OPC_EQUIVALENT_PART_NAMES,
    description: "Stored part names and content-types stream spellings are unique under OPC [M1.12] equivalence.",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    run: (ctx): ValidationIssue[] => {
        const groups = new Map<string, string[]>();
        for (const name of ctx.storedNames) {
            const normalized = normalizePartName(name);
            const key = isContentTypesStreamName(normalized) ? "stream" : `part:${canonicalPartName(normalized)}`;
            const group = groups.get(key);
            if (group) group.push(name);
            else groups.set(key, [name]);
        }
        const issues: ValidationIssue[] = [];
        for (const [key, names] of groups) {
            if (names.length < 2) continue;
            // Match the existing normalized index, then the canonical lookup.
            // An exact canonical spelling wins; otherwise code-unit order wins.
            const parts = [...indexByPartName(names.map((name) => [name, name] as const)).keys()].sort();
            const kept = key === "stream" ? findContentTypesStream(parts) : findEquivalentPart(parts, key.slice("part:".length));
            issues.push({
                code: OPC_EQUIVALENT_PART_NAMES,
                severity: "error",
                path: kept,
                message: `OPC [M1.12] forbids equivalent stored names: ${names.map((name) => JSON.stringify(name)).join(", ")}.`,
            });
        }
        return issues.sort((a, b) => (a.path! < b.path! ? -1 : a.path! > b.path! ? 1 : 0));
    },
};
