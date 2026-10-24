/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A w:tc whose last element child is w:tbl makes Word offer
 * recovery (fixture 0792). One such cell is enough. Inserting
 * an empty w:p after that table opens clean. Removing the
 * trailing w:p from a cell Word opened clean makes it warn.
 * A cell that already ends with w:p is not this defect.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

export const wordCellEndsWithTableCheck: CheckDefinition = {
    id: "word-cell-ends-with-table",
    description: "Cells ending with a nested table have a following paragraph.",
    groups: ["tables"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const cells = getElementsByTagNameNSAll(dom, ns, "tc");
                for (const cell of cells) {
                    let last: Element | null = null;
                    for (const child of Array.from(cell.childNodes)) {
                        if (!child || child.nodeType !== 1) {
                            continue;
                        }
                        last = child as Element;
                    }
                    if (!last) {
                        continue;
                    }
                    const local = last.localName || last.tagName.split(":").pop() || "";
                    if (local !== "tbl" || last.namespaceURI !== ns) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            "A table cell ends with w:tbl and has no following w:p. " +
                            "Microsoft Word offers to recover the file. Add a w:p after that table.",
                        path: xmlFile,
                        code: "word-cell-ends-with-table",
                    });
                }
            }
        }

        return issues;
    },
};
