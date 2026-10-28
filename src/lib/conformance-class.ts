/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The ISO OOXML conformance class of a package's parts, and whether a
 * finding may cite the Word session that showed its shape (#1016).
 *
 * A Word session observed one stored package. Every session a runtime
 * citation names stored Transitional parts outside the parts the finding's
 * own shape names, so a package with an ISO Strict part anywhere else is not
 * the package the session stored: the finding stays, without the citation
 * (#861).
 */

import { canonicalPartName } from "./content-types";
import { NS } from "./types";
import type { ValidationContext } from "./validation-context";
import { getElementsByTagNameNSAll } from "./xml-helpers";

/**
 * URI prefix shared by all ISO OOXML Strict conformance class namespaces.
 * Strict documents use `purl.oclc.org/ooxml/` instead of the Transitional
 * `schemas.openxmlformats.org/…` prefix. Any root element namespace that
 * starts with this prefix belongs to the Strict conformance class.
 *
 * The bundled `iso-strict/` schemas cover these namespaces. Under
 * `profile: "strict"` `validateAgainstXsd` validates Strict parts against
 * them (namespace + part identity dispatch); lenient keeps the historical
 * info-level `xsd-strict-skipped` skip, and word-valid walks no schema at
 * all, so it emits no XSD issue.
 */
export const STRICT_OOXML_NS_PREFIX = "http://purl.oclc.org/ooxml/";

const strictParts = new WeakMap<ValidationContext, Map<string, boolean>>();

const isStrict = (namespace: string | null | undefined): boolean => namespace?.startsWith(STRICT_OOXML_NS_PREFIX) ?? false;

/**
 * Whether `part` belongs to the Strict conformance class: its stored root,
 * or its root after markup compatibility's branch selection, is in a Strict
 * namespace, or it is a relationships part holding a relationship of a
 * Strict type. The content-types stream, the relationships parts and the
 * core properties share one namespace across both classes, so a
 * relationships part shows its class by its relationship types alone. A part
 * that does not parse shows no class.
 */
const isStrictPart = (ctx: ValidationContext, part: string): boolean => {
    const stored = ctx.xml(part);
    if (!stored?.ok) {
        return false;
    }
    const root = stored.document.documentElement;
    if (isStrict(root?.namespaceURI)) {
        return true;
    }
    const effective = ctx.effectiveXml(part);
    if (effective?.ok && isStrict(effective.document.documentElement?.namespaceURI)) {
        return true;
    }
    return getElementsByTagNameNSAll(stored.document, NS.PR, "Relationship").some((relationship) =>
        isStrict(relationship.getAttribute("Type")),
    );
};

/**
 * Whether a finding may cite the Word session that showed its shape: no
 * part of `xmlParts` outside `shapeParts` belongs to the Strict conformance
 * class. `xmlParts` are the parts the XML passes judge (`xmlPartNames` in
 * `checks/xml.ts`); `shapeParts` are the parts the finding's own shape
 * names, whose class the session's specimens set row by row. Each part's
 * class is read once per context, and only when a finding asks.
 */
export const citesWordSession = (ctx: ValidationContext, xmlParts: Iterable<string>, shapeParts: readonly string[] = []): boolean => {
    let known = strictParts.get(ctx);
    if (known === undefined) {
        known = new Map();
        strictParts.set(ctx, known);
    }
    const shape = new Set(shapeParts.map(canonicalPartName));
    for (const part of xmlParts) {
        const canonical = canonicalPartName(part);
        if (shape.has(canonical)) {
            continue;
        }
        let strict = known.get(canonical);
        if (strict === undefined) {
            strict = isStrictPart(ctx, part);
            known.set(canonical, strict);
        }
        if (strict) {
            return false;
        }
    }
    return true;
};
