/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Markup compatibility as Word processes it. Word selects AlternateContent
 * branches and honours mc:MustUnderstand with the namespaces it understands
 * (`WORD_UNDERSTOOD_NAMESPACES`), so the errors of that pass are judged on
 * the context's Word view, without the schema walk. Only parts under `word/`
 * are judged: the schema walk applied markup compatibility to the same
 * folder, and no Word evidence covers the package-level parts.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";

const WORD_XML_PART = /^word\/.+\.xml$/i;

export const wordMarkupCompatibilityCheck: CheckDefinition = {
    id: "word-markup-compatibility",
    description: "Markup compatibility in Word parts is well formed and requires only namespaces Word understands.",
    groups: ["xml"],
    formats: ["docx"],
    run: (ctx): ValidationIssue[] =>
        ctx.partNames
            .filter((name) => WORD_XML_PART.test(name))
            .flatMap((name) => ctx.compatibilityIssues(name).filter((issue) => issue.severity === "error")),
};
