/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Body-level m:oMathPara containing m:sPre makes Word refuse to open the file.
 * Fixture: tests/fixtures/external/superdoc/behavior/math-spre-tests.docx; paragraph-wrapped math is tolerated.
 * Unparseable parts retain word-math-parse when the check is called directly; xml-syntax gates registry runs.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

export const wordMathParseCheck: CheckDefinition = {
    id: "word-math-parse",
    description: "The main Word document part parses before its math is inspected.",
    groups: ["math"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    reportsParseFailures: true,
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        const documentXml = ctx.partNames.find((part) => path.posix.basename(part) === "document.xml" && part.startsWith("word/"));
        if (!documentXml) {
            return issues;
        }
        const parsed = ctx.effectiveXml(documentXml);
        if (parsed && !parsed.ok) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${parsed.message}`,
                path: documentXml,
                code: "word-math-parse",
            });
        }
        return issues;
    },
};

export const wordMathSpreBodyCheck: CheckDefinition = {
    id: "word-math-spre-body",
    description: "Display math containing a prescript is wrapped in a paragraph.",
    groups: ["math"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        const documentXml = ctx.partNames.find((part) => path.posix.basename(part) === "document.xml" && part.startsWith("word/"));
        if (!documentXml) {
            return issues;
        }
        const parsed = ctx.effectiveXml(documentXml);
        if (!parsed?.ok) {
            return issues;
        }
        const dom = parsed.document;
        let body: Element | null = null;
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const [node] = getElementsByTagNameNSAll(dom, ns, "body");
            if (node) {
                body = node;
                break;
            }
        }
        if (!body) {
            return issues;
        }
        for (const node of Array.from(body.childNodes)) {
            if (!node || node.nodeType !== 1) {
                continue;
            }
            const elem = node as Element;
            const tagName = elem.localName || elem.tagName.split(":").pop() || elem.tagName;
            if (elem.namespaceURI !== MATH_NAMESPACE || tagName !== "oMathPara") {
                continue;
            }
            if (getElementsByTagNameNSAll(elem, MATH_NAMESPACE, "sPre").length === 0) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    "Microsoft Word refuses to open documents with a body-level <m:oMathPara> " +
                    "that contains <m:sPre>; wrap it in a paragraph or avoid m:sPre in display math.",
                path: documentXml,
                code: "word-math-spre-body",
            });
        }

        return issues;
    },
};
