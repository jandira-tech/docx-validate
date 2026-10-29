/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import type { CheckDefinition } from "../../../lib/check-registry";

export const OPC_PART_NAME_WORD_REFUSED = "opc-part-name-word-refused";

/**
 * Word 16.115 offers recovery for these spec-valid spellings in a DOCX
 * package (#408, #443, #451; docs/plans/evidence/word-probes-2026-10-24.md,
 * sessions 6–12) — the only format probed, so the message names that scope
 * and claims nothing about a PPTX/XLSX package Word cannot open (#473).
 * Consume whole escape triplets: `%2521` contains `%25`, not `%21`.
 * Unreserved escapes belong to the existing OPC [M1.8] checks.
 */
export function wordRefusedNameReasons(name: string, includeLiteral = true): readonly string[] {
    const reasons = new Set<string>();
    for (let index = 0; index < name.length; index += 1) {
        const character = name[index]!;
        if (character === "%" && /^[0-9a-f]{2}$/i.test(name.slice(index + 1, index + 3))) {
            const octet = Number.parseInt(name.slice(index + 1, index + 3), 16);
            if ([0x21, 0x27, 0x28, 0x29, 0x2a].includes(octet)) reasons.add("percent-escaped RFC 2396 marks (!, ', (, ), *)");
            if (octet >= 0x80) reasons.add("percent-escaped octets at or above 0x80");
            index += 2;
        } else if (includeLiteral && character.charCodeAt(0) > 0x7f) {
            reasons.add("literal non-ASCII characters");
        }
    }
    return [...reasons];
}

/** Read every stored file-entry spelling, including names a refused package view drops. */
export const opcPartNameWordRefusedCheck: CheckDefinition = {
    id: OPC_PART_NAME_WORD_REFUSED,
    description: "Stored names contain no escaped marks, escaped high octets or literal non-ASCII characters that Word refuses.",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    storedNamesOnly: true,
    run: (ctx) =>
        [...new Set(ctx.storedNames)].flatMap((name) => {
            const reasons = wordRefusedNameReasons(name);
            return reasons.length === 0
                ? []
                : [
                      {
                          code: OPC_PART_NAME_WORD_REFUSED,
                          severity: "error" as const,
                          path: name,
                          message: `Stored name ${JSON.stringify(name)} contains ${reasons.join("; ")}; Word 16.115 offers recovery for this spelling in a DOCX package.`,
                      },
                  ];
        }),
};
