/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Markup Compatibility in the content-types stream (#459).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { findContentTypesStream } from "../../../lib/content-types";
import { NS, type ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";

export const CT_MARKUP_COMPATIBILITY = "ct-markup-compatibility";

const tokens = (value: string): string[] => value.trim().split(/\s+/).filter(Boolean);

const isXmlnsAttribute = (attr: Attr): boolean => attr.name === "xmlns" || attr.name.startsWith("xmlns:");

/**
 * Report one finding for a content-types stream that uses Markup
 * Compatibility: an attribute in the MC namespace (`mc:Ignorable`,
 * `mc:ProcessContent`, `mc:MustUnderstand`), an element in that namespace
 * (`mc:AlternateContent` and its branches), or an element or attribute in a
 * namespace an `mc:Ignorable` lists. A bare `xmlns:mc` declaration that
 * nothing uses is not a finding. The message lists every such element and
 * attribute once per occurrence, in document order (#495).
 *
 * Microsoft Word 16.115 does not process Markup Compatibility in this stream;
 * it offers recovery for every shape probed (session 14a in
 * docs/plans/evidence/word-probes-2026-10-24.md), so the finding is
 * Word-blocking for DOCX/DOCM. Strict reports it as an error — the superset
 * contract, and its XSD walk already refuses the same stream — while lenient
 * and the unobserved PPTX/XLSX warn. The other content-types checks judge the
 * stream AFTER markup compatibility is applied (OPC [M1.19]); this check reads
 * the raw stored stream, which is the only place the markup can still be seen.
 *
 * A content-types stream that does not parse is reported by the XML syntax
 * gate; nothing is reported from it here.
 */
export const ctMarkupCompatibilityCheck: CheckDefinition = {
    id: CT_MARKUP_COMPATIBILITY,
    description: "The content-types stream uses no Markup Compatibility markup (#459).",
    groups: ["content-types"],
    formats: ["docx", "pptx", "xlsx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) return [];
        const parsed = ctx.xml(stream);
        if (!parsed?.ok || !parsed.document.documentElement) return [];

        // One entry per occurrence, in document order: the message lists
        // every Markup Compatibility element and attribute the stream holds
        // (#495), not one mention per tag kind. Document order puts an
        // element before its own attributes (WHATWG DOM: an attribute is
        // FOLLOWING | CONTAINED_BY, 20, against the element owning it), so
        // the element's entry comes first (#526).
        const found: string[] = [];
        const visit = (element: Element, inherited: ReadonlySet<string>): void => {
            // mc:Ignorable applies to the element it sits on and its
            // descendants, as applyMarkupCompatibility merges it.
            const ignorable = new Set(inherited);
            for (const prefix of tokens(element.getAttributeNS(NS.MC, "Ignorable") ?? "")) {
                const uri = prefix === "xml" ? NS.XML : element.lookupNamespaceURI(prefix);
                if (uri) ignorable.add(uri);
            }
            if (element.namespaceURI === NS.MC) {
                found.push(`<${element.tagName}> element`);
            } else if (element.namespaceURI && ignorable.has(element.namespaceURI)) {
                found.push(`element <${element.tagName}> in the ignorable namespace ${JSON.stringify(element.namespaceURI)}`);
            }
            for (const attr of Array.from(element.attributes)) {
                if (isXmlnsAttribute(attr)) continue;
                if (attr.namespaceURI === NS.MC) {
                    found.push(`${attr.name} attribute on <${element.tagName}>`);
                } else if (attr.namespaceURI && attr.namespaceURI !== NS.XML && ignorable.has(attr.namespaceURI)) {
                    found.push(
                        `${attr.name} attribute on <${element.tagName}> in the ignorable namespace ${JSON.stringify(attr.namespaceURI)}`,
                    );
                }
            }
            for (let child = element.firstChild; child; child = child.nextSibling) {
                if (child.nodeType === 1 /* ELEMENT_NODE */) visit(child as Element, ignorable);
            }
        };
        visit(parsed.document.documentElement, new Set());
        if (found.length === 0) return [];
        return [
            {
                severity: "error",
                code: CT_MARKUP_COMPATIBILITY,
                path: stream,
                message: `Markup Compatibility markup in the content-types stream: ${[...found].join("; ")}`,
            },
        ];
    },
};
