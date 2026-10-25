/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { format } from "oxfmt";
import { runCli } from "../src/lib/run-cli";
import { getElementsByTagNameAll, parseXml } from "../src/lib/xml-helpers";

const XSD = "http://www.w3.org/2001/XMLSchema";
const RELATIONSHIPS = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);
const ATTRIBUTES = new Set(["id", "embed", "link"]);
const SCHEMAS = fileURLToPath(new URL("../src/scripts/office/schemas/", import.meta.url));
const OUTPUT = new URL("../src/lib/relationship-reference-elements.ts", import.meta.url);

type SchemaSource = { name: string; xml: string };
type AttributeResult = { attributes: Set<string>; uncertain: boolean };
export type RelationshipElements = {
    declared: Set<string>;
    references: Map<string, Set<string>>;
    uncertain: Set<string>;
    schemaFiles: number;
    elementSchemaFiles: number;
};

const children = (element: Element): Element[] => {
    const result: Element[] = [];
    for (let child = element.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 1 && (child as Element).namespaceURI === XSD) result.push(child as Element);
    }
    return result;
};
const key = (namespace: string, name: string): string => `{${namespace}}${name}`;
const qname = (element: Element, value: string): string => {
    const colon = value.indexOf(":");
    const prefix = colon < 0 ? null : value.slice(0, colon);
    const namespace = prefix === "xml" ? "http://www.w3.org/XML/1998/namespace" : element.lookupNamespaceURI(prefix ?? "");
    if (prefix !== null && namespace === null) throw new Error(`Unbound XSD prefix: ${value}`);
    return key(namespace ?? "", colon < 0 ? value : value.slice(colon + 1));
};
const merge = (target: AttributeResult, source: AttributeResult): void => {
    for (const attribute of source.attributes) target.attributes.add(attribute);
    target.uncertain ||= source.uncertain;
};

/** Read every bundled XSD, including supporting schemas with no elements. No network resolution. */
export const readBundledSchemas = (): SchemaSource[] => {
    const walk = (dir: string): string[] =>
        readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
            const file = path.join(dir, entry.name);
            return entry.isDirectory() ? walk(file) : entry.name.endsWith(".xsd") ? [file] : [];
        });
    return walk(SCHEMAS)
        .sort()
        .map((file) => ({ name: path.relative(SCHEMAS, file).split(path.sep).join("/"), xml: readFileSync(file, "utf8") }));
};

/**
 * Union all declarations of an expanded element name. Follow only attributes
 * of its own type, never attributes of child elements. Unknown type/ref chains,
 * cycles, xs:anyType and attribute wildcards stay explicitly uncertain, so they
 * cannot establish that an attribute is absent. This is not an XSD validator.
 */
export const collectRelationshipElements = (sources: readonly SchemaSource[]): RelationshipElements => {
    const schemas = sources.map((source) => ({ ...source, root: parseXml(source.xml).documentElement! }));
    const owners = new Map<Document, Element>();
    const globals = new Map<string, Element[]>();
    for (const { root } of schemas) {
        owners.set(root.ownerDocument!, root);
        const namespace = root.getAttribute("targetNamespace") ?? "";
        for (const declaration of children(root)) {
            const name = declaration.getAttribute("name");
            if (!name) continue;
            const id = `${declaration.localName}:${key(namespace, name)}`;
            globals.set(id, [...(globals.get(id) ?? []), declaration]);
        }
    }
    const cached = new Map<Element, AttributeResult>();
    const visiting = new Set<Element>();
    const resolveGlobal = (kind: string, id: string): AttributeResult => {
        if (id.startsWith(`{${XSD}}`)) return { attributes: new Set(), uncertain: id === key(XSD, "anyType") };
        const declarations = globals.get(`${kind}:${id}`) ?? (kind === "complexType" ? globals.get(`simpleType:${id}`) : undefined);
        const result: AttributeResult = { attributes: new Set(), uncertain: !declarations };
        for (const declaration of declarations ?? []) merge(result, resolve(declaration));
        return result;
    };
    const resolve = (element: Element): AttributeResult => {
        const previous = cached.get(element);
        if (previous) return previous;
        if (visiting.has(element)) return { attributes: new Set(), uncertain: true };
        visiting.add(element);
        const result: AttributeResult = { attributes: new Set(), uncertain: false };
        const kind = element.localName;
        if (kind === "element") {
            const ref = element.getAttribute("ref");
            const type = element.getAttribute("type");
            const inline = children(element).find((child) => child.localName === "complexType" || child.localName === "simpleType");
            if (ref) merge(result, resolveGlobal("element", qname(element, ref)));
            else if (type) merge(result, resolveGlobal("complexType", qname(element, type)));
            else if (inline) merge(result, resolve(inline));
            else {
                const head = element.getAttribute("substitutionGroup");
                if (head) merge(result, resolveGlobal("element", qname(element, head)));
                else result.uncertain = true; // The implicit type is xs:anyType.
            }
        } else if (kind !== "simpleType") {
            const base = element.getAttribute("base");
            // A restriction of anyType defines its own attribute uses; it does
            // not inherit anyType's open wildcard. Other bases are followed.
            if (base && !(kind === "restriction" && qname(element, base) === key(XSD, "anyType"))) {
                merge(result, resolveGlobal("complexType", qname(element, base)));
            }
            const ref = element.getAttribute("ref");
            if (kind === "attributeGroup" && ref) merge(result, resolveGlobal("attributeGroup", qname(element, ref)));
            for (const child of children(element)) {
                if (child.localName === "attribute") {
                    const ref = child.getAttribute("ref");
                    const root = owners.get(child.ownerDocument!)!;
                    const qualified =
                        child.parentNode === root ||
                        (child.getAttribute("form") || root.getAttribute("attributeFormDefault")) === "qualified";
                    const id = ref
                        ? qname(child, ref)
                        : key(qualified ? (root.getAttribute("targetNamespace") ?? "") : "", child.getAttribute("name") ?? "");
                    const end = id.indexOf("}");
                    const name = id.slice(end + 1);
                    if (RELATIONSHIPS.has(id.slice(1, end)) && ATTRIBUTES.has(name)) {
                        if (child.getAttribute("use") === "prohibited") result.attributes.delete(name);
                        else result.attributes.add(name);
                    }
                } else if (child.localName === "anyAttribute") {
                    result.uncertain = true;
                } else if (["complexContent", "simpleContent", "extension", "restriction", "attributeGroup"].includes(child.localName)) {
                    merge(result, resolve(child));
                }
            }
        }
        visiting.delete(element);
        cached.set(element, result);
        return result;
    };
    const result: RelationshipElements = {
        declared: new Set(),
        references: new Map(),
        uncertain: new Set(),
        schemaFiles: sources.length,
        elementSchemaFiles: 0,
    };
    for (const { root } of schemas) {
        const elements = getElementsByTagNameAll(root).filter((element) => element.namespaceURI === XSD && element.localName === "element");
        if (elements.length) result.elementSchemaFiles += 1;
        for (const element of elements) {
            const ref = element.getAttribute("ref");
            const qualified =
                element.parentNode === root || (element.getAttribute("form") || root.getAttribute("elementFormDefault")) === "qualified";
            const id = ref
                ? qname(element, ref)
                : key(qualified ? (root.getAttribute("targetNamespace") ?? "") : "", element.getAttribute("name") ?? "");
            result.declared.add(id);
            const attributes = resolve(element);
            if (attributes.attributes.size) {
                const union = result.references.get(id) ?? new Set();
                for (const attribute of attributes.attributes) union.add(attribute);
                result.references.set(id, union);
            }
            if (attributes.uncertain) result.uncertain.add(id);
        }
    }
    // Match the same Transitional/iso-strict schema filenames used by
    // STRICT_SCHEMA_MAPPINGS in validators/base.ts. Read targetNamespace from
    // the actual counterpart schemas (also covering hyphen/camel-case changes).
    // Union both directions so conformance classes classify identically even
    // where Strict removes a Transitional declaration (e.g. w:subDoc).
    for (const schema of schemas.filter((schema) => schema.name.startsWith("ISO-IEC29500-4_2016/"))) {
        const counterpart = schemas.find((candidate) => candidate.name === `iso-strict/${path.posix.basename(schema.name)}`);
        if (!counterpart) continue;
        const a = schema.root.getAttribute("targetNamespace") ?? "";
        const b = counterpart.root.getAttribute("targetNamespace") ?? "";
        if (a === b) continue;
        for (const id of [...result.declared]) {
            const end = id.indexOf("}");
            const namespace = id.slice(1, end);
            if (namespace !== a && namespace !== b) continue;
            const other = key(namespace === a ? b : a, id.slice(end + 1));
            result.declared.add(other);
            const union = new Set([...(result.references.get(id) ?? []), ...(result.references.get(other) ?? [])]);
            if (union.size) {
                result.references.set(id, union);
                result.references.set(other, union);
            }
            if (result.uncertain.has(id) || result.uncertain.has(other)) {
                result.uncertain.add(id);
                result.uncertain.add(other);
            }
        }
    }
    return result;
};

/** Deterministic, formatted output; used in memory by the sync test. */
export const renderRelationshipElements = async (data: RelationshipElements): Promise<string> => {
    const groups: Record<string, string> = {};
    for (const id of [...data.declared].sort()) {
        const end = id.indexOf("}");
        const ns = id.slice(1, end);
        groups[ns] = [groups[ns], id.slice(end + 1)].filter(Boolean).join(" ");
    }
    const references = Object.fromEntries(
        [...data.references].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)).map(([id, names]) => [id, [...names].sort().join(" ")]),
    );
    const source = `/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
// Generated by scripts/generate-relationship-reference-elements.ts; do not edit.
// ${data.schemaFiles} XSD files (${data.elementSchemaFiles} with elements), ${Object.keys(groups).length} namespaces,
// ${data.declared.size} element pairs, ${data.references.size} relationship-bearing pairs, ${data.uncertain.size} unresolved pairs.
// Transitional and ISO Strict counterparts are unioned conservatively.
export const SCHEMA_DECLARED_ELEMENTS: ReadonlySet<string> = new Set(
    Object.entries(${JSON.stringify(groups)}).flatMap(([namespace, names]) => names.split(" ").map((name) => \`{\${namespace}}\${name}\`)),
);
export const SCHEMA_REFERENCE_ELEMENTS: ReadonlyMap<string, ReadonlySet<string>> = new Map(
    Object.entries(${JSON.stringify(references)}).map(([element, names]) => [element, new Set(names.split(" "))]),
);
// An unresolved type cannot prove absence of a relationship attribute.
export const SCHEMA_UNRESOLVED_ELEMENTS: ReadonlySet<string> = new Set(${JSON.stringify([...data.uncertain].sort())});
`;
    const formatted = await format("relationship-reference-elements.ts", source, { tabWidth: 4, printWidth: 140 });
    if (formatted.errors.length) throw new Error(JSON.stringify(formatted.errors));
    return formatted.code;
};

runCli(import.meta.url, async () => {
    const data = collectRelationshipElements(readBundledSchemas());
    writeFileSync(OUTPUT, await renderRelationshipElements(data));
    console.log(
        `Generated ${data.declared.size} declared pairs, ${data.references.size} reference pairs, ${data.uncertain.size} unresolved pairs from ${data.schemaFiles} schemas.`,
    );
    return 0;
});
