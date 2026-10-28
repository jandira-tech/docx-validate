/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { CONTENT_TYPES_NAMESPACE, isContentTypesStreamName, isRelationshipsPartName } from "./content-types";
import { NS, type ValidationIssue } from "./types";

export const UNDERSTOOD_NAMESPACES: ReadonlySet<string> = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://www.w3.org/XML/1998/namespace",
    NS.MC,
]);

// ISO/IEC 29500 Strict moves each Transitional OOXML namespace under
// purl.oclc.org/ooxml/ (the targetNamespace of each schemas/iso-strict/*.xsd);
// markup compatibility and xml: keep one URI in both conformance classes.
// Transitional URIs stay understood so mixed content reaches the Strict XSD
// instead of being dropped as ignorable.
export const STRICT_UNDERSTOOD_NAMESPACES: ReadonlySet<string> = new Set([
    ...UNDERSTOOD_NAMESPACES,
    "http://purl.oclc.org/ooxml/officeDocument/math",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
    "http://purl.oclc.org/ooxml/schemaLibrary/main",
    "http://purl.oclc.org/ooxml/drawingml/main",
    "http://purl.oclc.org/ooxml/drawingml/chart",
    "http://purl.oclc.org/ooxml/drawingml/chartDrawing",
    "http://purl.oclc.org/ooxml/drawingml/diagram",
    "http://purl.oclc.org/ooxml/drawingml/picture",
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
    "http://purl.oclc.org/ooxml/officeDocument/sharedTypes",
]);

// The namespaces Microsoft Word declares on the root of the document parts
// it saves (Word for Mac 16.x, e.g. working/package.word-anchor-open-repaired
// .docx), on top of both conformance classes. Word selects the
// AlternateContent branches and keeps the Ignorable content of these
// namespaces, so the Word rules judge the view this set produces.
export const WORD_UNDERSTOOD_NAMESPACES: ReadonlySet<string> = new Set([
    ...STRICT_UNDERSTOOD_NAMESPACES,
    "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "http://schemas.microsoft.com/office/drawing/2010/main",
    "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/ink",
    "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "http://schemas.microsoft.com/office/2019/extlst",
    "urn:schemas-microsoft-com:office:office",
    "urn:schemas-microsoft-com:vml",
    "urn:schemas-microsoft-com:office:word",
    "http://schemas.microsoft.com/office/word/2006/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordml",
    "http://schemas.microsoft.com/office/word/2012/wordml",
    "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "http://schemas.microsoft.com/office/word/2018/wordml",
    "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
]);

// The processor that applies OPC's Part 2 rules to a package stream
// understands that stream's own vocabulary: an `mc:Ignorable` naming the
// content-types namespace cannot un-declare every part's content type, and
// one naming the package-relationships namespace cannot hide the
// relationships from the checks that read `.rels` parts (#493). Word parts
// keep the Word-understood set alone.
export const CONTENT_TYPES_UNDERSTOOD_NAMESPACES: ReadonlySet<string> = new Set([...WORD_UNDERSTOOD_NAMESPACES, CONTENT_TYPES_NAMESPACE]);
export const RELATIONSHIPS_UNDERSTOOD_NAMESPACES: ReadonlySet<string> = new Set([...WORD_UNDERSTOOD_NAMESPACES, NS.PR]);

/** The understood namespaces for the processed view of one part: the stream's own vocabulary for the content-types and relationships streams, the Word set otherwise. */
export const understoodNamespacesForPart = (part: string): ReadonlySet<string> => {
    if (isContentTypesStreamName(part)) {
        return CONTENT_TYPES_UNDERSTOOD_NAMESPACES;
    }
    return isRelationshipsPartName(part) ? RELATIONSHIPS_UNDERSTOOD_NAMESPACES : WORD_UNDERSTOOD_NAMESPACES;
};

// ISO-IEC29500-4_2016/dml-chart.xsd:30 and sml.xsd:232 CT_Extension
// require one lax xsd:any child. dml-main.xsd:183 CT_OfficeArtExtension
// and pml.xsd:761 CT_Extension allow zero or more lax xsd:any children.
// In these slots preserve the foreign payload for the schema's lax verdict.
// The Strict schemas declare the same slots (iso-strict/dml-chart.xsd:29,
// sml.xsd:224, dml-main.xsd:181, pml.xsd:759).
const EXTENSION_NAMESPACES = new Set([
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    NS.A,
    NS.P,
    NS.S,
    "http://purl.oclc.org/ooxml/drawingml/chart",
    "http://purl.oclc.org/ooxml/drawingml/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
]);
type Context = {
    ignorable: ReadonlySet<string>;
    processContent: ReadonlySet<string>;
};
const tokens = (value: string): string[] => value.trim().split(/\s+/).filter(Boolean);
// XML 1.0 NCName: NameStartChar/NameChar without the namespace separator.
const NCNAME_START =
    "A-Z_a-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\u{10000}-\u{EFFFF}";
const NCNAME = new RegExp(`^[${NCNAME_START}][${NCNAME_START}0-9.\\-\u00B7\u0300-\u036F\u203F-\u2040]*$`, "u");
const children = (element: Element): Element[] => Array.from(element.childNodes).filter((node) => node.nodeType === 1) as Element[];
// Word 16.115 offered recovery for character data directly inside an
// mc:AlternateContent — before the Choice, between the Choice and the
// Fallback, after the Fallback, after a Choice with no Fallback and between
// two Choices, in a used comments part's stored root (sessions 65 s65c and
// 67 s67b) and in a run-level body AlternateContent (session 67 s67b) alike
// — while opening the same roots clean with the text inside a Choice or the
// Fallback (session 65 s65b) and with whitespace only between the branches
// (session 67 s67b). Whitespace is note 8's [ \t\n\r], which pretty-printed
// parts carry; mc.xsd declares none of the three elements mixed, so it
// cannot draw Word's line. CDATA sections are unprobed.
const holdsCharacterData = (element: Element): boolean =>
    Array.from(element.childNodes).some((node) => node.nodeType === 3 && (node.nodeValue ?? "").match(/[^ \t\n\r]/) !== null);
const namespaceUri = (element: Element, prefix: string): string | null => (prefix === "xml" ? NS.XML : element.lookupNamespaceURI(prefix));
const unboundPrefixIssue = (directive: string, prefix: string): ValidationIssue => ({
    severity: "error",
    code: "mce-unbound-prefix",
    message: `Unbound prefix in mc:${directive}: ${prefix}`,
});
/** Build a validation-only view; source XML and unselected branches remain untouched. */
export const applyMarkupCompatibility = (
    source: Document,
    opts: { understood: ReadonlySet<string> },
): { document: Document; issues: ValidationIssue[]; rootUnwrapElementCount?: number; rootUnwrapIssue?: ValidationIssue } => {
    const document = source.cloneNode(true) as Document;
    const issues: ValidationIssue[] = [];
    let rootUnwrapElementCount: number | undefined;
    let rootUnwrapIssue: ValidationIssue | undefined;
    const unwrap = (element: Element): void => {
        const parent = element.parentNode!;
        if (parent.nodeType === 9) {
            const count = children(element).length;
            if (count !== 1) {
                // #640: validate the recursively processed children before moving
                // any into a Document; a second element throws in xmldom, and
                // zero leaves no root. Return the stored root so its existing
                // rules keep the verdict: Word 16.115 offered recovery for a
                // _rels/.rels root AlternateContent whose Fallback holds two
                // Relationship elements (session 36, s36c, in
                // docs/plans/evidence/word-probes-2026-10-24.md), and for one
                // that unwraps to none, its Relationships element in an
                // unselected Choice (session 38, s38a). One whose Fallback
                // holds the Relationships element unwraps to a single root and
                // opened clean in the same chunk (#646); other zero- and
                // multiple-element root shapes are unobserved (#658).
                rootUnwrapElementCount = count;
                rootUnwrapIssue = {
                    severity: "error",
                    code: "mce-malformed",
                    message: `Root unwrapping requires exactly one element; processed content has ${count}. The part is left unprocessed.`,
                };
                issues.push(rootUnwrapIssue);
                return;
            }
        }
        const next = element.nextSibling;
        // Detach first: a Document cannot contain both the old and selected roots.
        parent.removeChild(element);
        while (element.firstChild) {
            parent.insertBefore(element.firstChild, next);
        }
    };
    const undeclaredExtensions = new Map<string, Set<string>>();
    let extensionAttributeCount = 0;
    const visit = (element: Element, inherited: Context, attributesOnly = false): void => {
        // Application extension slots are opaque to MCE, including their directives.
        if (element.localName === "ext" && EXTENSION_NAMESPACES.has(element.namespaceURI || "")) {
            return;
        }
        const directiveIssues: ValidationIssue[] = [];
        const ignorable = new Set(inherited.ignorable);
        const processContent = new Set(inherited.processContent);
        for (const prefix of tokens(element.getAttributeNS(NS.MC, "Ignorable") || "")) {
            const uri = namespaceUri(element, prefix);
            if (uri) {
                ignorable.add(uri);
            } else {
                directiveIssues.push(unboundPrefixIssue("Ignorable", prefix));
            }
        }
        for (const name of tokens(element.getAttributeNS(NS.MC, "ProcessContent") || "")) {
            const [prefix, local, ...rest] = name.split(":");
            if (rest.length > 0 || !NCNAME.test(prefix) || (local !== "*" && !NCNAME.test(local ?? ""))) {
                directiveIssues.push({
                    severity: "error",
                    code: "mce-invalid-token",
                    message: `Malformed QName in mc:ProcessContent: ${name}`,
                });
                continue;
            }
            const uri = namespaceUri(element, prefix);
            if (uri) {
                processContent.add(`${uri}#${local}`);
            } else {
                directiveIssues.push(unboundPrefixIssue("ProcessContent", prefix));
            }
        }
        const elementNamespace = element.namespaceURI || "";
        const ignored = ignorable.has(elementNamespace) && !opts.understood.has(elementNamespace);
        if (ignored && !processContent.has(`${elementNamespace}#${element.localName}`) && !processContent.has(`${elementNamespace}#*`)) {
            element.parentNode!.removeChild(element);
            return;
        }
        issues.push(...directiveIssues);
        for (const prefix of tokens(element.getAttributeNS(NS.MC, "MustUnderstand") || "")) {
            const uri = namespaceUri(element, prefix);
            if (!uri) {
                issues.push(unboundPrefixIssue("MustUnderstand", prefix));
            } else if (!opts.understood.has(uri)) {
                issues.push({
                    severity: "error",
                    code: "mce-must-understand",
                    message: `Required namespace is not understood: ${prefix}`,
                });
            }
        }
        if (attributesOnly) {
            return;
        }
        const context = {
            ignorable,
            processContent,
        };
        for (const attr of Array.from(element.attributes)) {
            // Word understands Office extension attributes even when the part
            // omits Ignorable. The ISO schemas cannot validate these attributes;
            // remove them only from this clone and diagnose the missing directive.
            // A namespace the consumer understands (a standalone extension
            // part's own schema) keeps its attributes.
            if (
                attr.namespaceURI?.startsWith("http://schemas.microsoft.com/office/") &&
                !ignorable.has(attr.namespaceURI) &&
                !opts.understood.has(attr.namespaceURI)
            ) {
                const names = undeclaredExtensions.get(attr.namespaceURI) ?? new Set<string>();
                names.add(attr.name);
                undeclaredExtensions.set(attr.namespaceURI, names);
                extensionAttributeCount += 1;
                element.removeAttributeNode(attr);
                continue;
            }
            if (
                attr.namespaceURI === NS.MC ||
                (attr.namespaceURI && ignorable.has(attr.namespaceURI) && !opts.understood.has(attr.namespaceURI))
            ) {
                element.removeAttributeNode(attr);
            }
        }
        if (element.namespaceURI === NS.MC && element.localName === "AlternateContent") {
            for (const child of children(element)) {
                const uri = child.namespaceURI || "";
                if (ignorable.has(uri) && !opts.understood.has(uri)) {
                    element.removeChild(child);
                }
            }
            const branches = children(element);
            if (holdsCharacterData(element)) {
                issues.push({
                    severity: "error",
                    code: "mce-malformed",
                    message:
                        "AlternateContent holds character data directly inside it; Word 16.115 offered recovery for every position probed (sessions 65 and 67).",
                });
            }
            let choices = 0;
            let fallbacks = 0;
            let malformed = false;
            for (const branch of branches) {
                if (branch.namespaceURI !== NS.MC) {
                    malformed = true;
                    continue;
                }
                if (branch.localName === "Choice") {
                    choices += 1;
                    const required = tokens(branch.getAttribute("Requires") || "");
                    if (fallbacks > 0 || required.length === 0 || required.some((prefix) => !namespaceUri(branch, prefix))) {
                        malformed = true;
                    }
                } else if (branch.localName === "Fallback") {
                    fallbacks += 1;
                    if (fallbacks > 1) {
                        malformed = true;
                    }
                } else {
                    malformed = true;
                }
            }
            if (choices === 0 || malformed) {
                issues.push({
                    severity: "error",
                    code: "mce-malformed",
                    message: "AlternateContent requires Choices with declared Requires prefixes, followed by at most one Fallback.",
                });
            }
            // mc.xsd declares Requires use="required": a Choice without a
            // non-empty Requires is malformed (flagged above) and must never
            // be selected — an empty token list would otherwise be vacuously
            // understood and win over the well-formed Fallback (#494).
            const selectableChoice = (node: Element): boolean => {
                const required = tokens(node.getAttribute("Requires") || "");
                return (
                    node.namespaceURI === NS.MC &&
                    node.localName === "Choice" &&
                    required.length > 0 &&
                    required.every((prefix) => opts.understood.has(namespaceUri(node, prefix) || ""))
                );
            };
            const selected =
                branches.find(selectableChoice) ?? branches.find((node) => node.namespaceURI === NS.MC && node.localName === "Fallback");
            for (const branch of branches) {
                if (
                    branch !== selected &&
                    branch.namespaceURI === NS.MC &&
                    (branch.localName === "Choice" || branch.localName === "Fallback")
                ) {
                    // Branch attributes always apply; unselected contents do not exist.
                    visit(branch, context, true);
                }
            }
            if (selected) {
                visit(selected, context);
                unwrap(selected);
            }
            // Only selected content enters the validation view.
            for (const branch of children(element)) {
                if (branches.includes(branch)) {
                    element.removeChild(branch);
                }
            }
            unwrap(element);
            return;
        }
        if (ignored) {
            for (const child of children(element)) {
                visit(child, context);
            }
            unwrap(element);
            return;
        }
        for (const child of children(element)) {
            visit(child, context);
        }
    };
    if (document.documentElement) {
        visit(document.documentElement, {
            ignorable: new Set(),
            processContent: new Set(),
        });
    }
    if (source.documentElement && !document.documentElement) {
        // The whole part is ignorable content: there is nothing left for a
        // schema to judge, so say so instead of validating an empty view.
        issues.push({
            severity: "info",
            code: "mce-root-ignored",
            message: `Root element ${source.documentElement.tagName} is in an ignorable namespace that is not understood; the part has no content to schema-validate.`,
        });
    }
    if (extensionAttributeCount > 0) {
        const namespaces = Array.from(undeclaredExtensions, ([uri, names]) => `${uri}: ${Array.from(names).sort().join(", ")}`);
        issues.push({
            severity: "warning",
            code: "mce-extension-not-ignorable",
            message: `Removed ${extensionAttributeCount} attributes from the validation view in non-ignorable Office extension namespaces (${namespaces.join("; ")}). Declare the corresponding prefixes in mc:Ignorable.`,
        });
    }
    return rootUnwrapElementCount === undefined
        ? { document, issues }
        : { document: source.cloneNode(true) as Document, issues, rootUnwrapElementCount, rootUnwrapIssue };
};
