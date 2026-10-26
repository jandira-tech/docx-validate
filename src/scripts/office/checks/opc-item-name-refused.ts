/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import type { CheckDefinition } from "../../../lib/check-registry";
import { entryNameRefusal } from "../../../lib/validation-context";

export const OPC_ITEM_NAME_REFUSED = "opc-item-name-refused";

/**
 * Every stored item name the package view refuses (#629): a leading slash, a
 * drive, or an empty or dot-only segment (#372, #378). The view still refuses
 * to stage such a package, so `package-extract-failed` stays; this check names
 * each refused entry, where the refusal stops at the first. Word 16.115 offers
 * recovery for a `..` name in the central directory (#411; session 3 in
 * docs/plans/evidence/word-probes-2026-10-24.md).
 */
export const opcItemNameRefusedCheck: CheckDefinition = {
    id: OPC_ITEM_NAME_REFUSED,
    description: "Stored item names place a package part: no leading slash, no drive, and no empty or dot-only segment.",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    storedNamesOnly: true,
    run: (ctx) =>
        [...new Set(ctx.storedNames)].flatMap((name) => {
            const refusal = entryNameRefusal(name);
            return refusal === undefined ? [] : [{ code: OPC_ITEM_NAME_REFUSED, severity: "error" as const, path: name, message: refusal }];
        }),
};
