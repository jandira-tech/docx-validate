/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Markup compatibility as Word processes it. Word selects AlternateContent
 * branches and honours mc:MustUnderstand with the namespaces it understands
 * (`WORD_UNDERSTOOD_NAMESPACES`), so the errors of that pass are judged on
 * the context's Word view, without the schema walk. The judged parts are the
 * XML parts under `word/` and every relationships part: the schema walk
 * applied markup compatibility to the same folder, and Word 16.115 applies
 * it to relationships parts too — it opened session 33's footer shapes
 * (s33d/s33e, #632) and refused mc:MustUnderstand on the `_rels/.rels` root
 * and an undeclared mc:Choice Requires prefix in
 * `word/_rels/document.xml.rels` (session 34, s34c/s34d, #636). Which
 * findings stay errors under word-valid is keyed by part in
 * `validators/word-blocking.ts`; unobserved relationships parts keep the
 * profile's demotion there. Parts outside this scope keep whatever
 * markup-compatibility findings the schema walk makes under lenient and
 * strict.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { isRelationshipsPartName } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";

const WORD_XML_PART = /^word\/.+\.xml$/i;

/**
 * The parts this check judges: the XML parts under `word/` and every
 * relationships part, as OPC compares part names. Shared with the partial
 * runs' root-Ignorable pairing (#499) so full and named runs agree (#381).
 */
export const isMarkupCompatibilityJudgedPart = (partName: string): boolean =>
    WORD_XML_PART.test(partName) || isRelationshipsPartName(partName);

export const wordMarkupCompatibilityCheck: CheckDefinition = {
    id: "word-markup-compatibility",
    description: "Markup compatibility in Word parts and relationships parts is well formed and requires only namespaces Word understands.",
    groups: ["xml"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] =>
        ctx.partNames
            .filter(isMarkupCompatibilityJudgedPart)
            .flatMap((name) => ctx.compatibilityIssues(name).filter((issue) => issue.severity === "error")),
};
