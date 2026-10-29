/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The global element declarations of a bundled XSD, and of every schema it
 * imports or includes, transitively — the set of roots that schema can judge
 * as a document root (#887). `_schemaMappingKey` routes a part stored
 * directly in `ppt/` under a name outside `PPT_DIRECT_PART_NAMES` by asking
 * whether the root it holds is one of these declarations, so a renamed main
 * presentation or a relocated slide keeps its pml.xsd check while a root no
 * bundled schema declares stays unmapped. The declarations are read from the
 * bundled files themselves, never hand-listed: pml.xsd's closure reaches
 * DrawingML roots (`a:tblStyleLst`, `a:theme`) through its dml-main.xsd
 * import that a name list would have to copy and keep current.
 */

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { parseXmlBytes } from "../../lib/xml-helpers";

/** XML Schema's own namespace: `xsd:schema`, `xsd:element`, `xsd:import`, `xsd:include`. */
const XSD_NAMESPACE = "http://www.w3.org/2001/XMLSchema";

/** A `schemaLocation` that names a remote resource (a URL), which no bundled schema resolves through. */
const URL_LOCATION = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;

/** Each entry schema's declarations, read once, keyed by the entry schema's path. */
const closures = new Map<string, ReadonlySet<string>>();

/**
 * The `{namespace}localName` spelling of a global element declaration's
 * identity — the same braces the engine's own findings spell QNames in.
 */
export const globalElementKey = (namespace: string, localName: string): string => `{${namespace}}${localName}`;

/**
 * The global element declarations of `schemaFile` and of every schema its
 * `xsd:import` and `xsd:include` elements reach, transitively, as
 * {@link globalElementKey} keys. A declaration belongs to its schema's
 * target namespace — an included schema that declares none (the chameleon
 * case) borrows the including one's. A `schemaLocation` that is a URL, or
 * names no local file, is skipped: only the bundled files decide. A file
 * that cannot be read or parsed contributes nothing, so the answer rests on
 * the schemas that are there, exactly as the engine's own load would see
 * them.
 */
export const schemaGlobals = (schemaFile: string): ReadonlySet<string> => {
    const cached = closures.get(schemaFile);
    if (cached !== undefined) return cached;
    const globals = new Set<string>();
    // A schema included from two parents with different target namespaces
    // declares its chameleon globals in each, so the walk keys a file by the
    // namespace it declares them in, not by its path alone.
    const visited = new Set<string>();
    const pending: { file: string; namespace: string | null }[] = [{ file: schemaFile, namespace: null }];
    while (pending.length > 0) {
        const { file, namespace } = pending.shift()!;
        const visit = `${file}\u0000${namespace ?? ""}`;
        if (visited.has(visit)) continue;
        visited.add(visit);
        let schema: Element;
        try {
            schema = parseXmlBytes(readFileSync(file)).documentElement;
        } catch {
            continue;
        }
        if (schema.namespaceURI !== XSD_NAMESPACE || schema.localName !== "schema") continue;
        const targetNamespace = schema.getAttribute("targetNamespace") || namespace;
        for (const child of Array.from(schema.childNodes)) {
            if (child.nodeType !== 1) continue;
            const declaration = child as Element;
            if (declaration.namespaceURI !== XSD_NAMESPACE) continue;
            if (declaration.localName === "element" && declaration.getAttribute("name")) {
                globals.add(globalElementKey(targetNamespace ?? "", declaration.getAttribute("name")!));
            }
            if (declaration.localName === "import" || declaration.localName === "include") {
                const location = declaration.getAttribute("schemaLocation");
                if (!location || URL_LOCATION.test(location)) continue;
                const resolved = path.resolve(path.dirname(file), location);
                if (!existsSync(resolved)) continue;
                pending.push({ file: resolved, namespace: declaration.localName === "include" ? targetNamespace : null });
            }
        }
    }
    closures.set(schemaFile, globals);
    return globals;
};
