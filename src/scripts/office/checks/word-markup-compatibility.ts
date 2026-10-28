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
 * `validators/word-blocking.ts`, which compares canonical part names too;
 * unobserved relationships parts keep the profile's demotion there. Parts
 * outside this scope keep whatever markup-compatibility findings the schema
 * walk makes under lenient and strict. The judged names compare as OPC
 * compares part names — the canonical form, ASCII case folded and with
 * escapes of unreserved characters decoded — so `word/FOO.XML` and
 * `word/foo.%78ml` are judged as `word/foo.xml` is (#639), on both halves.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { citesWordSession } from "../../../lib/conformance-class";
import { canonicalPartName, isRelationshipsPartName } from "../../../lib/content-types";
import { isAlternateContentCharacterData, withWordSessionCitation } from "../../../lib/markup-compatibility";
import type { ValidationIssue } from "../../../lib/types";
import { xmlPartNames } from "./xml";

const WORD_XML_PART = /^word\/.+\.xml$/;

/**
 * The parts this check judges: the XML parts under `word/` and every
 * relationships part, as OPC compares part names — the canonical name
 * carries the judgement on both halves, as `isRelationshipsPartName` and
 * `isXmlPartName` already compare theirs (#639). Shared with the partial
 * runs' root-Ignorable pairing (#499) so full and named runs agree (#381).
 */
export const isMarkupCompatibilityJudgedPart = (partName: string): boolean =>
    WORD_XML_PART.test(canonicalPartName(partName)) || isRelationshipsPartName(partName);

export const wordMarkupCompatibilityCheck: CheckDefinition = {
    id: "word-markup-compatibility",
    description: "Markup compatibility in Word parts and relationships parts is well formed and requires only namespaces Word understands.",
    groups: ["xml"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues = ctx.partNames
            .filter(isMarkupCompatibilityJudgedPart)
            .flatMap((name) => ctx.compatibilityIssues(name).filter((issue) => issue.severity === "error"));
        // Sessions 65 and 67 stored Transitional packages (#1016).
        return issues.some(isAlternateContentCharacterData) &&
            citesWordSession(
                ctx,
                xmlPartNames(ctx.partNames, (name) => ctx.bytes(name)),
            )
            ? issues.map(withWordSessionCitation)
            : issues;
    },
};
