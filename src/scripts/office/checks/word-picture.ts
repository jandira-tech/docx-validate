/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";

// CT_Picture requires nvPicPr, blipFill, spPr in sequence (dml-picture.xsd).
const missingPictureChildCheck = (id: string, childName: string, position: string): CheckDefinition => ({
    id,
    description: `Pictures have direct pic:${childName} properties.`,
    groups: ["drawings"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of ctx.storyParts()) {
            const parsed = ctx.xml(part);
            if (!parsed?.ok) {
                continue;
            }
            for (const pic of getElementsByTagNameNSAll(parsed.document, PICTURE_NAMESPACE, "pic")) {
                const hasProperties = Array.from(pic.childNodes).some(
                    (child) =>
                        child.nodeType === 1 &&
                        (child as Element).namespaceURI === PICTURE_NAMESPACE &&
                        (child as Element).localName === childName,
                );
                if (hasProperties) {
                    continue;
                }
                issues.push({
                    severity: "error",
                    code: id,
                    path: part,
                    message: `pic:${childName} is missing from parent pic:pic; expected a direct pic:${childName} ${position} pic:blipFill (CT_Picture).`,
                });
            }
        }
        return issues;
    },
});

export const wordPictureNvpicprMissingCheck = missingPictureChildCheck("word-picture-nvpicpr-missing", "nvPicPr", "before");
export const wordPictureSpprMissingCheck = missingPictureChildCheck("word-picture-sppr-missing", "spPr", "after");
