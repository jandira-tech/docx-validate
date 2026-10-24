/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * WordprocessingML namespace constants. A leaf module so named checks and the
 * DOCX validator can share them without importing each other;
 * `validators/docx.ts` re-exports them as part of the public API.
 */

export const WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const WORD_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
/** Python: `WORD_PARAGRAPH_NAMESPACES = (WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE)` */
export const WORD_PARAGRAPH_NAMESPACES: readonly [string, string] = [WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE];
