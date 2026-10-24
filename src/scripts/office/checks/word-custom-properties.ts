/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const CUSTOM_PROPERTIES_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const VARIANT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/** Empty custom-property filetimes make Word offer recovery. */
export const wordCustomPropertyFiletimeEmptyCheck: CheckDefinition = {
    id: "word-custom-property-filetime-empty",
    description: "Custom-property vt:filetime scalars contain a date and time.",
    groups: ["settings"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => {
        const parsed = ctx.effectiveXml("docProps/custom.xml");
        if (!parsed?.ok) {
            return [];
        }
        const issues: ValidationIssue[] = [];
        for (const filetime of getElementsByTagNameNSAll(parsed.document, VARIANT_TYPES_NAMESPACE, "filetime")) {
            const parent = filetime.parentNode;
            if (parent?.nodeType !== 1) {
                continue;
            }
            const property = parent as Element;
            if (property.namespaceURI !== CUSTOM_PROPERTIES_NAMESPACE || property.localName !== "property") {
                continue;
            }
            if (!/^[ \t\r\n]*$/.test(filetime.textContent ?? "")) {
                continue;
            }
            issues.push({
                severity: "error",
                code: "word-custom-property-filetime-empty",
                path: "docProps/custom.xml",
                message:
                    "vt:filetime inside custom-property property must contain a non-empty dateTime value. Microsoft Word offers recovery for an empty value.",
            });
        }
        return issues;
    },
};
