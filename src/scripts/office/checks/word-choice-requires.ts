/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * mc:Choice @Requires names a prefix that is not declared on
 * that element or an ancestor (fixture 0508). Word offers
 * recovery. A declaration on a descendant, or on the fallback
 * sibling, does not put the prefix in scope.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const prefixDeclaredInScope = (elem: Element, prefix: string): boolean => {
    // XML Namespaces predefines this binding, even without xmlns:xml.
    if (prefix === "xml") {
        return true;
    }
    const attrName = `xmlns:${prefix}`;
    let current: Node | null = elem;
    while (current && current.nodeType === 1) {
        const attrs = (current as Element).attributes;
        for (const attr of attrs) {
            if (attr && attr.name === attrName) {
                return true;
            }
        }
        current = current.parentNode;
    }
    return false;
};

export const wordChoiceRequiresUndeclaredCheck: CheckDefinition = {
    id: "word-choice-requires-undeclared",
    description: "Choice Requires prefixes are declared on the choice or its ancestors.",
    groups: ["xml"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of ctx.storyParts()) {
            const parsed = ctx.xml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const choices = getElementsByTagNameNSAll(dom, MC_NAMESPACE, "Choice");
            for (const choice of choices) {
                const requires = choice.getAttribute("Requires");
                if (requires === null) {
                    continue;
                }
                const missing: string[] = [];
                for (const prefix of requires.split(/\s+/)) {
                    if (prefix.length === 0) {
                        continue;
                    }
                    if (prefixDeclaredInScope(choice, prefix)) {
                        continue;
                    }
                    missing.push(prefix);
                }
                if (missing.length === 0) {
                    continue;
                }
                issues.push({
                    severity: "error",
                    message:
                        `mc:Choice Requires lists undeclared prefix ${missing.join(", ")}. ` +
                        "Microsoft Word offers to recover the file. " +
                        "Declare the prefix on the Choice element or an ancestor.",
                    path: xmlFile,
                    code: "word-choice-requires-undeclared",
                });
            }
        }

        return issues;
    },
};
