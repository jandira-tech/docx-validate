/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * WordprocessingML and math namespace constants. A leaf module so named
 * checks and the DOCX validator can share them without importing each other;
 * `validators/docx.ts` re-exports them as part of the public API. Transitional
 * namespaces come first, Strict (`purl.oclc.org`) second; checks iterate both
 * so a rule cannot go blind to one conformance class.
 */

export const WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const WORD_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
/** Python: `WORD_PARAGRAPH_NAMESPACES = (WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE)` */
export const WORD_PARAGRAPH_NAMESPACES: readonly [string, string] = [WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE];

export const MATH_2006_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const MATH_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/math";
/** Math namespaces, Transitional first, Strict second. */
export const MATH_NAMESPACES: readonly [string, string] = [MATH_2006_NAMESPACE, MATH_STRICT_NAMESPACE];
