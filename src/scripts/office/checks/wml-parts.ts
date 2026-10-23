/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { ValidationContext } from "../../../lib/validation-context";

export const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export type WmlPartDocument = { readonly part: string; readonly document: Document };

/**
 * Every package part whose root element is WordprocessingML, parsed once
 * through the shared context. Content-model placement rules apply to the
 * whole WordprocessingML surface — stories, styles, numbering, settings,
 * web settings, and the font table — not only to the parts that carry user
 * text, so a rule must not assume `ctx.storyParts()` is enough.
 *
 * Parts that fail to parse are skipped: XML syntax is reported by its own
 * check, and a DOM rule cannot say anything about a part it cannot read.
 * Parts rooted in another namespace (strict OOXML stories, DrawingML themes,
 * OPC and docProps parts) are skipped as well: their content models belong
 * to other schemas.
 */
export const wmlPartDocuments = (ctx: ValidationContext): readonly WmlPartDocument[] => {
    const out: WmlPartDocument[] = [];
    for (const part of ctx.partNames) {
        const parsed = ctx.xml(part);
        if (!parsed?.ok) {
            continue;
        }
        if (parsed.document.documentElement?.namespaceURI !== WORDPROCESSINGML_NAMESPACE) {
            continue;
        }
        out.push({ part, document: parsed.document });
    }
    return out;
};
