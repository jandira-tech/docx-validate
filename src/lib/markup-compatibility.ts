/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

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

// ISO-IEC29500-4_2016/dml-chart.xsd:30 and sml.xsd:232 CT_Extension
// require one lax xsd:any child. dml-main.xsd:183 CT_OfficeArtExtension
// and pml.xsd:761 CT_Extension allow zero or more lax xsd:any children.
// In these slots preserve the foreign payload for the schema's lax verdict.
const EXTENSION_NAMESPACES = new Set(["http://schemas.openxmlformats.org/drawingml/2006/chart", NS.A, NS.P, NS.S]);
type Context = { ignorable: ReadonlySet<string>; processContent: ReadonlySet<string>; extensionPayload: boolean };
const tokens = (value: string): string[] => value.trim().split(/\s+/).filter(Boolean);
const children = (element: Element): Element[] => Array.from(element.childNodes).filter((node) => node.nodeType === 1) as Element[];
const namespaceUri = (element: Element, prefix: string): string | null => (prefix === "xml" ? NS.XML : element.lookupNamespaceURI(prefix));
const unwrap = (element: Element): void => {
    const parent = element.parentNode!;
    while (element.firstChild) {
        parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
};

/** Build a validation-only view; source XML and unselected branches remain untouched. */
export const applyMarkupCompatibility = (
    source: Document,
    opts: { understood: ReadonlySet<string> },
): { document: Document; issues: ValidationIssue[] } => {
    const document = source.cloneNode(true) as Document;
    const issues: ValidationIssue[] = [];
    const visit = (element: Element, inherited: Context): void => {
        const ignorable = new Set(inherited.ignorable);
        const processContent = new Set(inherited.processContent);
        for (const prefix of tokens(element.getAttributeNS(NS.MC, "Ignorable") || "")) {
            const uri = namespaceUri(element, prefix);
            if (uri) {
                ignorable.add(uri);
            }
        }
        for (const name of tokens(element.getAttributeNS(NS.MC, "ProcessContent") || "")) {
            const [prefix, local] = name.split(":");
            const uri = namespaceUri(element, prefix);
            if (uri && local) {
                processContent.add(`${uri}#${local}`);
            }
        }
        for (const prefix of tokens(element.getAttributeNS(NS.MC, "MustUnderstand") || "")) {
            const uri = namespaceUri(element, prefix);
            if (!uri || !opts.understood.has(uri)) {
                issues.push({ severity: "error", code: "mce-must-understand", message: `Required namespace is not understood: ${prefix}` });
            }
        }
        const context = {
            ignorable,
            processContent,
            extensionPayload:
                inherited.extensionPayload || (element.localName === "ext" && EXTENSION_NAMESPACES.has(element.namespaceURI || "")),
        };
        for (const attr of Array.from(element.attributes)) {
            if (
                attr.namespaceURI === NS.MC ||
                (!inherited.extensionPayload &&
                    attr.namespaceURI &&
                    ignorable.has(attr.namespaceURI) &&
                    !opts.understood.has(attr.namespaceURI))
            ) {
                element.removeAttributeNode(attr);
            }
        }
        if (element.namespaceURI === NS.MC && element.localName === "AlternateContent") {
            const branches = children(element);
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
            const selected =
                branches.find(
                    (node) =>
                        node.namespaceURI === NS.MC &&
                        node.localName === "Choice" &&
                        tokens(node.getAttribute("Requires") || "").every((prefix) =>
                            opts.understood.has(namespaceUri(node, prefix) || ""),
                        ),
                ) ?? branches.find((node) => node.namespaceURI === NS.MC && node.localName === "Fallback");
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
        const elementNamespace = element.namespaceURI || "";
        if (!inherited.extensionPayload && ignorable.has(elementNamespace) && !opts.understood.has(elementNamespace)) {
            if (processContent.has(`${elementNamespace}#${element.localName}`) || processContent.has(`${elementNamespace}#*`)) {
                for (const child of children(element)) {
                    visit(child, context);
                }
                unwrap(element);
            } else {
                element.parentNode!.removeChild(element);
            }
            return;
        }
        for (const child of children(element)) {
            visit(child, context);
        }
    };
    if (document.documentElement) {
        visit(document.documentElement, { ignorable: new Set(), processContent: new Set(), extensionPayload: false });
    }
    return { document, issues };
};
