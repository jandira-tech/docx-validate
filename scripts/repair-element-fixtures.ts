/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Repair every `word-element-fixtures/<element>.docx` to the absolute minimal
 * form that still passes validation while remaining representative of the
 * element's XML rule: the element is kept, placed at the schema-correct
 * position: shortest path across same-name declarations, earliest part root
 * (document first) within one declaration, through the bundled XSD
 * content-model graph; with exactly the required children and required
 * attributes the schema demands — nothing more.
 *
 * Elements that are roots of their own Word part (styles, settings, …) get
 * that part added to the fixture package with its content-type Override and
 * relationship; the main document stays a minimal valid document.
 *
 * Usage:
 *   tsx scripts/repair-element-fixtures.ts [--out word-element-fixtures] [--results <jsonl>] [--limit n] [--only a,b] [--resume]
 */

import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import JSZip from "jszip";

import { runCli } from "../src/lib/run-cli";
import { parseXmlBytes } from "../src/lib/xml-helpers";
import {
    compareCodeUnits,
    fixtureMatchesRow,
    limitFlagPassed,
    loadManifestRows,
    parseLimitOption,
    persistManifestRows,
    resolveInside,
} from "./element-fixture-cli";
import type { Profile } from "../src/lib/types";
import { Validate } from "../src/validate";

const XSD = "http://www.w3.org/2001/XMLSchema";
const XMLNS = "http://www.w3.org/XML/1998/namespace";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const SCHEMAS = path.join(REPO, "src", "scripts", "office", "schemas");
const WML_XSD = path.join(SCHEMAS, "ISO-IEC29500-4_2016", "wml.xsd");

const key = (ns: string, name: string): string => `{${ns}}${name}`;
const splitKey = (id: string): [string, string] => {
    const end = id.indexOf("}");
    return [id.slice(1, end), id.slice(end + 1)];
};

const children = (element: Element): Element[] => {
    const result: Element[] = [];
    for (let child = element.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 1 && (child as Element).namespaceURI === XSD) result.push(child as Element);
    }
    return result;
};

const qname = (element: Element, value: string): string => {
    const colon = value.indexOf(":");
    const prefix = colon < 0 ? "" : value.slice(0, colon);
    let namespace = prefix === "xml" ? XMLNS : (element.lookupNamespaceURI(prefix) ?? undefined);
    if (namespace === undefined && prefix === "") {
        const root = element.ownerDocument?.documentElement;
        namespace = root?.lookupNamespaceURI("") ?? root?.getAttribute("targetNamespace") ?? "";
    }
    if (namespace === undefined) throw new Error(`Unbound XSD prefix: ${value}`);
    return key(namespace, colon < 0 ? value : value.slice(colon + 1));
};

interface SchemaFile {
    name: string;
    root: Element;
    ns: string;
}

const loadClosure = (entry: string, ns: string | null, out: Map<string, SchemaFile>): void => {
    const file = path.resolve(entry);
    if (out.has(file)) return;
    const root = parseXmlBytes(readFileSync(file)).documentElement!;
    const ownNs = root.getAttribute("targetNamespace") || ns || "";
    out.set(file, { name: path.relative(SCHEMAS, file).split(path.sep).join("/"), root, ns: ownNs });
    for (const child of children(root)) {
        if (child.localName !== "import" && child.localName !== "include") continue;
        const location = child.getAttribute("schemaLocation");
        if (!location || /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(location)) continue;
        const resolved = path.resolve(path.dirname(file), location);
        if (!existsSync(resolved)) continue;
        loadClosure(resolved, child.localName === "include" ? ownNs : null, out);
    }
};

/** Effective namespace of an element declaration inside a schema file. */
const effectiveNs = (decl: Element, fileNs: string): string => {
    const root = decl.ownerDocument!.documentElement!;
    const form = decl.getAttribute("form") ?? root.getAttribute("elementFormDefault");
    return form === "qualified" ? (root.getAttribute("targetNamespace") ?? fileNs) : "";
};

class SchemaModel {
    /** Every `xsd:element` declaration (global and local), keyed by effective `{ns}name`. */
    public readonly elements = new Map<string, Element[]>();
    public readonly types = new Map<string, Element[]>();
    public readonly groups = new Map<string, Element[]>();
    public readonly attrGroups = new Map<string, Element[]>();
    public readonly files: SchemaFile[] = [];

    public constructor(entry: string) {
        const loaded = new Map<string, SchemaFile>();
        loadClosure(entry, null, loaded);
        this.files = [...loaded.values()];
        for (const { root, ns } of this.files) {
            const walk = (el: Element): void => {
                for (const child of children(el)) {
                    const name = child.getAttribute("name");
                    if (child.localName === "element" && name) {
                        const id = key(effectiveNs(child, ns), name);
                        this.elements.set(id, [...(this.elements.get(id) ?? []), child]);
                    } else if (name && (child.localName === "complexType" || child.localName === "simpleType")) {
                        const id = key(ns, name);
                        this.types.set(id, [...(this.types.get(id) ?? []), child]);
                    } else if (name && child.localName === "group") {
                        this.groups.set(key(ns, name), [...(this.groups.get(key(ns, name)) ?? []), child]);
                    } else if (name && child.localName === "attributeGroup") {
                        this.attrGroups.set(key(ns, name), [...(this.attrGroups.get(key(ns, name)) ?? []), child]);
                    }
                    walk(child);
                }
            };
            walk(root);
        }
    }

    public resolve(kind: "complexType" | "simpleType" | "group" | "attributeGroup" | "element", id: string): Element | undefined {
        const map =
            kind === "element" ? this.elements : kind === "group" ? this.groups : kind === "attributeGroup" ? this.attrGroups : this.types;
        return map.get(id)?.[0];
    }

    private readonly subCache = new Map<string, string[]>();

    /** Element ids whose substitutionGroup chain reaches `head`. */
    public substitutionMembers(head: string): string[] {
        const cached = this.subCache.get(head);
        if (cached) return cached;
        const members: string[] = [];
        for (const [id, decls] of this.elements) {
            for (const decl of decls) {
                let sub = decl.getAttribute("substitutionGroup");
                let hops = 0;
                while (sub && hops++ < 10) {
                    const subId = qname(decl, sub);
                    if (subId === head) {
                        members.push(id);
                        break;
                    }
                    sub =
                        (this.elements.get(subId) ?? [])
                            .find((d) => d.getAttribute("substitutionGroup"))
                            ?.getAttribute("substitutionGroup") ?? null;
                }
            }
        }
        this.subCache.set(head, members);
        return members;
    }
}

type Particle =
    | { kind: "element"; id: string; decl: Element; min: number; max: number }
    | { kind: "group"; id: string; min: number; max: number }
    | { kind: "model"; el: Element; min: number; max: number }
    | { kind: "any"; min: number; max: number };

const occurs = (el: Element): [number, number] => {
    const min = Number(el.getAttribute("minOccurs") ?? "1");
    const maxRaw = el.getAttribute("maxOccurs") ?? "1";
    return [min, maxRaw === "unbounded" ? Number.POSITIVE_INFINITY : Number(maxRaw)];
};

const parseModel = (model: SchemaModel, modelEl: Element): Particle[] => {
    const result: Particle[] = [];
    for (const child of children(modelEl)) {
        const [min, max] = occurs(child);
        if (child.localName === "element") {
            const ref = child.getAttribute("ref");
            if (ref) {
                const id = qname(child, ref);
                result.push({ kind: "element", id, decl: model.resolve("element", id) ?? child, min, max });
            } else {
                const name = child.getAttribute("name");
                if (name) {
                    const root = child.ownerDocument!.documentElement!;
                    const ns = effectiveNs(child, root.getAttribute("targetNamespace") ?? "");
                    result.push({ kind: "element", id: key(ns, name), decl: child, min, max });
                }
            }
        } else if (child.localName === "group") {
            const ref = child.getAttribute("ref");
            if (ref) result.push({ kind: "group", id: qname(child, ref), min, max });
        } else if (child.localName === "sequence" || child.localName === "choice" || child.localName === "all") {
            result.push({ kind: "model", el: child, min, max });
        } else if (child.localName === "any") {
            result.push({ kind: "any", min, max });
        }
    }
    return result;
};

/** The complexType/simpleType declaration behind an element declaration. */
const typeOf = (model: SchemaModel, el: Element): { complex?: Element; simple?: Element } => {
    const inline = children(el).find((c) => c.localName === "complexType" || c.localName === "simpleType");
    const typeAttr = el.getAttribute("type");
    if (typeAttr) {
        const id = qname(el, typeAttr);
        const t = model.resolve("complexType", id) ?? model.resolve("simpleType", id);
        if (t) return t.localName === "complexType" ? { complex: t } : { simple: t };
    }
    if (inline) return inline.localName === "complexType" ? { complex: inline } : { simple: inline };
    const sub = el.getAttribute("substitutionGroup");
    if (sub) {
        const head = model.resolve("element", qname(el, sub));
        if (head) return typeOf(model, head);
    }
    return {};
};

/** All particles of a complexType, following extension bases. */
const particlesOf = (model: SchemaModel, complex: Element, seen = new Set<Element>()): Particle[] => {
    if (seen.has(complex)) return [];
    seen.add(complex);
    const result: Particle[] = [];
    for (const child of children(complex)) {
        if (child.localName === "complexContent" || child.localName === "simpleContent") {
            for (const ext of children(child)) {
                if (ext.localName === "extension" || ext.localName === "restriction") {
                    const base = ext.getAttribute("base");
                    if (base && ext.localName === "extension") {
                        const baseDecl = model.resolve("complexType", qname(ext, base));
                        if (baseDecl) result.push(...particlesOf(model, baseDecl, seen));
                    }
                    const modelChild = children(ext).find(
                        (c) => c.localName === "sequence" || c.localName === "choice" || c.localName === "all" || c.localName === "group",
                    );
                    if (modelChild) {
                        const [min, max] = occurs(modelChild);
                        const ref = modelChild.getAttribute("ref");
                        if (modelChild.localName === "group" && ref) {
                            result.push({ kind: "group", id: qname(modelChild, ref), min, max });
                        } else {
                            result.push({ kind: "model", el: modelChild, min, max });
                        }
                    }
                }
            }
        } else if (
            child.localName === "sequence" ||
            child.localName === "choice" ||
            child.localName === "all" ||
            child.localName === "group"
        ) {
            const [min, max] = occurs(child);
            const ref = child.getAttribute("ref");
            if (child.localName === "group" && ref) {
                result.push({ kind: "group", id: qname(child, ref), min, max });
            } else {
                result.push({ kind: "model", el: child, min, max });
            }
        }
    }
    return result;
};

const expandGroup = (model: SchemaModel, id: string, seen = new Set<string>()): Particle[] => {
    if (seen.has(id)) return [];
    seen.add(id);
    const decl = model.resolve("group", id);
    if (!decl) return [];
    const modelChild = children(decl).find((c) => c.localName === "sequence" || c.localName === "choice" || c.localName === "all");
    return modelChild ? [{ kind: "model", el: modelChild, min: 1, max: 1 }] : [];
};

/** Flatten a particle tree to the element ids it can contain. */
const elementIdsOf = (model: SchemaModel, particles: Particle[], out: Set<string>, seenGroups = new Set<string>()): void => {
    for (const p of particles) {
        if (p.kind === "element") {
            out.add(p.id);
            for (const member of model.substitutionMembers(p.id)) out.add(member);
        } else if (p.kind === "group") {
            if (seenGroups.has(p.id)) continue;
            seenGroups.add(p.id);
            elementIdsOf(model, expandGroup(model, p.id), out, seenGroups);
        } else if (p.kind === "model") {
            elementIdsOf(model, parseModel(model, p.el), out, seenGroups);
        }
    }
};

const BUILTIN_VALUE: Record<string, string> = {
    string: "x",
    normalizedString: "x",
    token: "x",
    language: "en",
    Name: "x",
    NCName: "x",
    NMTOKEN: "x",
    ID: "x",
    IDREF: "x",
    anyURI: "x",
    QName: "w:x",
    boolean: "true",
    decimal: "0",
    integer: "0",
    nonNegativeInteger: "0",
    positiveInteger: "1",
    nonPositiveInteger: "0",
    negativeInteger: "-1",
    long: "0",
    int: "0",
    short: "0",
    byte: "0",
    unsignedLong: "0",
    unsignedInt: "0",
    unsignedShort: "0",
    unsignedByte: "0",
    float: "0",
    double: "0",
    hexBinary: "00",
    base64Binary: "AA==",
    date: "2000-01-01",
    dateTime: "2000-01-01T00:00:00Z",
    time: "00:00:00Z",
    duration: "PT0S",
    gYear: "2000",
    gYearMonth: "2000-01",
    gMonth: "01",
    gDay: "01",
    anySimpleType: "x",
};

const TYPE_NAME_VALUE: Record<string, string> = {
    ST_HexColor: "000000",
    ST_Guid: "00000000-0000-0000-0000-000000000000",
    ST_Lang: "en",
    ST_OnOff: "true",
    ST_DecimalNumber: "0",
    ST_UnsignedDecimalNumber: "0",
    ST_Merge: "continue",
    ST_HpsMeasure: "0",
    ST_TwipsMeasure: "0",
    ST_SignedTwipsMeasure: "0",
    ST_Percentage: "0",
    ST_FixedPercentage: "0",
    ST_DateTime: "2000-01-01T00:00:00Z",
    ST_String: "x",
    ST_ShortHexNumber: "00",
    ST_LongHexNumber: "00000000",
    ST_VerticalAlign: "baseline",
    ST_HorizontalAlign: "left",
};

const simpleValue = (model: SchemaModel, simple: Element | undefined, typeName: string, seen = new Set<Element>()): string => {
    if (!simple) return "x";
    if (seen.has(simple)) return "x";
    seen.add(simple);
    const [, name] = splitKey(typeName);
    if (name && TYPE_NAME_VALUE[name] !== undefined) return TYPE_NAME_VALUE[name];
    for (const child of children(simple)) {
        if (child.localName === "restriction") {
            const baseId = child.getAttribute("base");
            const enums = children(child)
                .filter((c) => c.localName === "enumeration")
                .map((c) => c.getAttribute("value") ?? "x");
            if (enums.length) return enums[0];
            const length = children(child)
                .find((c) => c.localName === "length")
                ?.getAttribute("value");
            const minLength = children(child)
                .find((c) => c.localName === "minLength")
                ?.getAttribute("value");
            const maxLength = children(child)
                .find((c) => c.localName === "maxLength")
                ?.getAttribute("value");
            const minInc = children(child)
                .find((c) => c.localName === "minInclusive")
                ?.getAttribute("value");
            let value: string;
            if (baseId) {
                const baseKey = qname(child, baseId);
                const base = model.resolve("simpleType", baseKey);
                value = base ? simpleValue(model, base, baseKey, seen) : (BUILTIN_VALUE[baseKey.slice(baseKey.indexOf("}") + 1)] ?? "x");
            } else {
                value = "x";
            }
            if (length !== undefined) {
                const n = Number(length);
                value =
                    baseId && baseId.includes("hexBinary")
                        ? "00".repeat(n)
                        : value.repeat(Math.ceil(n / Math.max(value.length, 1))).slice(0, n);
            }
            if (minLength !== undefined) value = value.padEnd(Number(minLength), "x");
            if (maxLength !== undefined) value = value.slice(0, Number(maxLength));
            if (minInc !== undefined && !Number.isNaN(Number(value)) && Number(value) < Number(minInc)) value = String(minInc);
            return value;
        }
        if (child.localName === "list") {
            const itemId = child.getAttribute("itemType");
            if (itemId) {
                const item = model.resolve("simpleType", qname(child, itemId));
                if (item) return simpleValue(model, item, qname(child, itemId), seen);
            }
            return "x";
        }
        if (child.localName === "union") {
            const members = (child.getAttribute("memberTypes") ?? "").split(/\s+/).filter(Boolean);
            const candidates: { value: string; base: string }[] = [];
            for (const member of members) {
                const id = qname(child, member);
                const m = model.resolve("simpleType", id);
                const value = m ? simpleValue(model, m, id, seen) : (BUILTIN_VALUE[id.slice(id.indexOf("}") + 1)] ?? "x");
                candidates.push({ value, base: baseBuiltin(model, id) });
            }
            const inline = children(child).find((c) => c.localName === "simpleType");
            if (inline) candidates.push({ value: simpleValue(model, inline, ""), base: "string" });
            const stringish = new Set([
                "string",
                "normalizedString",
                "token",
                "Name",
                "NCName",
                "NMTOKEN",
                "ID",
                "IDREF",
                "anyURI",
                "language",
                "QName",
                "anySimpleType",
                "union",
            ]);
            return candidates.find((c) => c.value && !stringish.has(c.base))?.value ?? candidates[0]?.value ?? "x";
        }
    }
    return "x";
};

const baseBuiltin = (model: SchemaModel, id: string, seen = new Set<string>()): string => {
    if (seen.has(id)) return "string";
    seen.add(id);
    if (id.startsWith(`{${XSD}}`)) return id.slice(id.indexOf("}") + 1);
    const t = model.resolve("simpleType", id);
    if (!t) return "string";
    for (const c of children(t)) {
        if (c.localName === "restriction") {
            const b = c.getAttribute("base");
            return b ? baseBuiltin(model, qname(c, b), seen) : "string";
        }
        if (c.localName === "list") {
            const it = c.getAttribute("itemType");
            return it ? baseBuiltin(model, qname(c, it), seen) : "string";
        }
        if (c.localName === "union") return "union";
    }
    return "string";
};

const valueForAttr = (model: SchemaModel, attr: Element): string => {
    const typeAttr = attr.getAttribute("type");
    if (typeAttr) {
        const id = qname(attr, typeAttr);
        const simple = model.resolve("simpleType", id);
        if (simple) return simpleValue(model, simple, id);
        if (id.startsWith(`{${XSD}}`)) return BUILTIN_VALUE[id.slice(id.indexOf("}") + 1)] ?? "x";
        return "x";
    }
    const inline = children(attr).find((c) => c.localName === "simpleType");
    if (inline) return simpleValue(model, inline, "");
    return "x";
};

interface AttrUse {
    ns: string;
    name: string;
    required: boolean;
    attr: Element;
}

const attributeUses = (model: SchemaModel, complex: Element, seen = new Set<Element>()): AttrUse[] => {
    if (seen.has(complex)) return [];
    seen.add(complex);
    const result: AttrUse[] = [];
    const add = (attr: Element): void => {
        const ref = attr.getAttribute("ref");
        if (ref) {
            if (attr.getAttribute("use") !== "prohibited") {
                const [ns, name] = splitKey(qname(attr, ref));
                result.push({ ns, name, required: attr.getAttribute("use") === "required", attr });
            }
            return;
        }
        const name = attr.getAttribute("name");
        if (!name || attr.getAttribute("use") === "prohibited") return;
        const root = attr.ownerDocument!.documentElement!;
        const qualified = (attr.getAttribute("form") ?? root.getAttribute("attributeFormDefault")) === "qualified";
        result.push({
            ns: qualified ? (root.getAttribute("targetNamespace") ?? "") : "",
            name,
            required: attr.getAttribute("use") === "required",
            attr,
        });
    };
    const addGroups = (el: Element): void => {
        for (const child of children(el)) {
            if (child.localName === "attribute") add(child);
            else if (child.localName === "attributeGroup") {
                const ref = child.getAttribute("ref");
                if (ref) {
                    const decl = model.resolve("attributeGroup", qname(child, ref));
                    if (decl) result.push(...attributeUses(model, decl, seen));
                }
            }
        }
    };
    for (const child of children(complex)) {
        if (child.localName === "complexContent" || child.localName === "simpleContent") {
            for (const ext of children(child)) {
                if (ext.localName === "extension" || ext.localName === "restriction") {
                    const base = ext.getAttribute("base");
                    if (base && ext.localName === "extension") {
                        const baseDecl = model.resolve("complexType", qname(ext, base));
                        if (baseDecl) result.push(...attributeUses(model, baseDecl, seen));
                    }
                    addGroups(ext);
                }
            }
        }
    }
    addGroups(complex);
    return result;
};

const PREFERRED_PREFIX: Record<string, string> = {
    [W]: "w",
    [R]: "r",
    "http://schemas.openxmlformats.org/officeDocument/2006/math": "m",
    "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing": "wp14",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:word": "w10",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "wps",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas": "wpc",
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram": "dgm",
    [XMLNS]: "xml",
};

class Prefixes {
    private readonly map = new Map<string, string>();
    private counter = 0;
    public of(ns: string): string {
        if (ns === "") return "";
        let prefix = this.map.get(ns);
        if (!prefix) {
            prefix = PREFERRED_PREFIX[ns] ?? `ns${this.counter++}`;
            this.map.set(ns, prefix);
        }
        return prefix;
    }
    public declarations(): string {
        return [...this.map].map(([ns, p]) => ` xmlns:${p}="${ns}"`).join("");
    }
}

const MAX_DEPTH = 16;

const escapeText = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (text: string): string => escapeText(text).replace(/"/g, "&quot;");

interface PathNode {
    id: string;
    decl: Element;
}

const buildElement = (model: SchemaModel, prefixes: Prefixes, path: PathNode[], pathIndex: number, depth: number): string => {
    if (depth > MAX_DEPTH) throw new Error(`depth cap at ${path[pathIndex].id}`);
    const { id, decl } = path[pathIndex];
    const [ns, local] = splitKey(id);
    const prefix = prefixes.of(ns);
    const tag = prefix ? `${prefix}:${local}` : local;
    const { complex, simple } = typeOf(model, decl);

    const attrs: string[] = [];
    let text = "";
    let childParticles: Particle[] = [];

    if (simple) {
        const typeAttr = decl.getAttribute("type");
        if (typeAttr) {
            const typeId = qname(decl, typeAttr);
            const st = model.resolve("simpleType", typeId);
            text = st ? simpleValue(model, st, typeId) : (BUILTIN_VALUE[typeId.slice(typeId.indexOf("}") + 1)] ?? "x");
        } else {
            text = simpleValue(model, simple, id);
        }
    } else if (complex) {
        for (const use of attributeUses(model, complex)) {
            if (!use.required) continue;
            const ap = prefixes.of(use.ns);
            const aname = ap ? `${ap}:${use.name}` : use.name;
            attrs.push(`${aname}="${escapeAttr(valueForAttr(model, use.attr))}"`);
        }
        let simpleContentBase: string | undefined;
        for (const child of children(complex)) {
            if (child.localName === "simpleContent") {
                for (const ext of children(child)) {
                    if (ext.localName === "extension" || ext.localName === "restriction") {
                        simpleContentBase = ext.getAttribute("base") ?? simpleContentBase;
                    }
                }
            }
        }
        if (simpleContentBase) {
            const baseId = qname(complex, simpleContentBase);
            const baseSimple = model.resolve("simpleType", baseId);
            text = baseSimple ? simpleValue(model, baseSimple, baseId) : (BUILTIN_VALUE[baseId.slice(baseId.indexOf("}") + 1)] ?? "x");
        } else {
            childParticles = particlesOf(model, complex);
        }
    }

    const must = path[pathIndex + 1]?.id;
    const childSpecs: PathNode[] = [];
    const contains = (p: Particle): boolean => {
        if (must === undefined) return false;
        const set = new Set<string>();
        elementIdsOf(model, [p], set, new Set());
        return set.has(must);
    };
    const collect = (particles: Particle[], groupSeen: Set<string>): void => {
        for (const p of particles) {
            if (p.kind === "element") {
                const ids = [p.id, ...model.substitutionMembers(p.id)];
                if (p.min > 0 || (must !== undefined && ids.includes(must))) {
                    const chosen = must !== undefined && ids.includes(must) ? must : p.id;
                    childSpecs.push({ id: chosen, decl: chosen === p.id ? p.decl : (model.resolve("element", chosen) ?? p.decl) });
                }
            } else if (p.kind === "group") {
                if (groupSeen.has(p.id) || (p.min === 0 && !contains(p))) continue;
                groupSeen.add(p.id);
                collect(expandGroup(model, p.id), groupSeen);
            } else if (p.kind === "model") {
                if (p.min === 0 && !contains(p)) continue;
                const subs = parseModel(model, p.el);
                if (p.el.localName === "choice") {
                    const hit = subs.find((s) => contains(s));
                    const chosen = hit ?? (p.min > 0 ? (subs.find((s) => s.min > 0) ?? subs[0]) : undefined);
                    if (chosen) collect([chosen], groupSeen);
                } else {
                    for (const s of subs) {
                        if (s.min > 0 || contains(s)) collect([s], groupSeen);
                    }
                }
            }
        }
    };
    collect(childParticles, new Set());

    const inner = childSpecs
        .map((c) => {
            const onPath = must !== undefined && c.id === must;
            const childPath = onPath ? path : [...path.slice(0, pathIndex + 1), c];
            return buildElement(model, prefixes, childPath, pathIndex + 1, depth + 1);
        })
        .join("");
    const body = text ? escapeText(text) : inner;
    const attrStr = attrs.length ? " " + attrs.join(" ") : "";
    return `<${tag}${attrStr}>${body}</${tag}>`;
};

/** Adjacency over declarations (not names): two elements with the same name but
 * different content models (e.g. block vs run `customXml`) stay distinct. */
const buildAdjacency = (model: SchemaModel): Map<Element, Set<PathNode>> => {
    const collectNodes = (particles: Particle[], out: Map<string, PathNode>, seenGroups: Set<string>): void => {
        for (const p of particles) {
            if (p.kind === "element") {
                if (!out.has(p.id)) out.set(p.id, { id: p.id, decl: p.decl });
                for (const member of model.substitutionMembers(p.id)) {
                    const decl = model.resolve("element", member);
                    if (decl && !out.has(member)) out.set(member, { id: member, decl });
                }
            } else if (p.kind === "group") {
                if (seenGroups.has(p.id)) continue;
                seenGroups.add(p.id);
                collectNodes(expandGroup(model, p.id), out, seenGroups);
            } else if (p.kind === "model") {
                collectNodes(parseModel(model, p.el), out, seenGroups);
            }
        }
    };
    const adj = new Map<Element, Set<PathNode>>();
    for (const [, decls] of model.elements) {
        for (const decl of decls) {
            const set = adj.get(decl) ?? new Set<PathNode>();
            const { complex } = typeOf(model, decl);
            if (complex) {
                const nodes = new Map<string, PathNode>();
                collectNodes(particlesOf(model, complex), nodes, new Set());
                for (const node of nodes.values()) set.add(node);
            }
            adj.set(decl, set);
        }
    }
    return adj;
};

/** Part root elements and the package part each one roots. `document` roots the main part. */
interface Rel {
    id: string;
    type: string;
    target: string;
    external?: boolean;
}

interface FixtureParts {
    files: Record<string, string | Uint8Array>;
    overrides: { part: string; contentType: string }[];
    defaults: { ext: string; contentType: string }[];
    /** Relationship sidecars, keyed by the part they belong to (e.g. "word/settings.xml"). */
    partRels: Record<string, Rel[]>;
}

const MAIN_DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/** Part root elements and the package part each one roots. `document` roots the main part. */
/** BFS per part root in ROOT_PARTS order (document first); across same-name declarations find() returns the shortest path. */
class PathIndex {
    private readonly parents = new Map<Element, { node: PathNode; root: string; parent: Element | null }>();

    public constructor(model: SchemaModel, adj: Map<Element, Set<PathNode>>) {
        // Breadth-first per part root, in ROOT_PARTS order with document first,
        // so primary-document paths win over equally short header/footer ones.
        for (const rootName of Object.keys(ROOT_PARTS)) {
            const decl = model.resolve("element", key(W, rootName));
            if (!decl || this.parents.has(decl)) continue;
            const rootNode = { id: key(W, rootName), decl };
            this.parents.set(decl, { node: rootNode, root: rootName, parent: null });
            const queue: { node: PathNode; root: string; parent: Element | null }[] = [{ node: rootNode, root: rootName, parent: null }];
            while (queue.length) {
                const { node, root, parent } = queue.shift()!;
                for (const child of adj.get(node.decl) ?? []) {
                    if (!this.parents.has(child.decl)) {
                        this.parents.set(child.decl, { node: child, root, parent: node.decl });
                        queue.push({ node: child, root, parent });
                    }
                }
            }
        }
    }

    public find(model: SchemaModel, targetId: string): { root: string; path: PathNode[] } | null {
        // Several declarations can share one qualified name (e.g. w:footnote
        // under both settings and footnotes); return the shortest path.
        let best: { root: string; path: PathNode[] } | null = null;
        for (const decl of model.elements.get(targetId) ?? []) {
            const entry = this.parents.get(decl);
            if (!entry) continue;
            const path: PathNode[] = [entry.node];
            let p = entry.parent;
            while (p) {
                const pe = this.parents.get(p)!;
                path.unshift(pe.node);
                p = pe.parent;
            }
            if (!best || path.length < best.path.length) best = { root: entry.root, path };
        }
        return best;
    }
}

const ROOT_PARTS: Record<string, { part: string; contentType: string; relType: string } | null> = {
    document: null,
    glossaryDocument: {
        part: "word/glossary/document.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
        relType: `${R}/glossaryDocument`,
    },
    styles: {
        part: "word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        relType: `${R}/styles`,
    },
    settings: {
        part: "word/settings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        relType: `${R}/settings`,
    },
    numbering: {
        part: "word/numbering.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        relType: `${R}/numbering`,
    },
    fonts: {
        part: "word/fontTable.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        relType: `${R}/fontTable`,
    },
    webSettings: {
        part: "word/webSettings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        relType: `${R}/webSettings`,
    },
    comments: {
        part: "word/comments.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        relType: `${R}/comments`,
    },
    footnotes: {
        part: "word/footnotes.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        relType: `${R}/footnotes`,
    },
    endnotes: {
        part: "word/endnotes.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        relType: `${R}/endnotes`,
    },
    hdr: {
        part: "word/header1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        relType: `${R}/header`,
    },
    ftr: {
        part: "word/footer1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        relType: `${R}/footer`,
    },
    recipients: {
        part: "word/recipients.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml",
        relType: `${R}/recipientData`,
    },
};

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const minimalDocument = (body = "<w:body/>"): string => {
    const prefixes = new Prefixes();
    prefixes.of(W);
    return XML_DECL + `<w:document${prefixes.declarations()}>${body}</w:document>\n`;
};

const renderChain = (model: SchemaModel, path: PathNode[]): string => {
    const prefixes = new Prefixes();
    const [rootNs, rootLocal] = splitKey(path[0].id);
    const rootPrefix = prefixes.of(rootNs);
    const rootTag = rootPrefix ? `${rootPrefix}:${rootLocal}` : rootLocal;
    return buildElement(model, prefixes, path, 0, 0).replace(`<${rootTag}`, `<${rootTag}${prefixes.declarations()}`);
};

interface RelSpec {
    type: string;
    part?: string;
    contentType?: string;
    binary?: boolean;
    /** Exact XML content for the target part (used when the part has an +xml content type). */
    content?: string;
}

const FONT_REL: RelSpec = {
    type: `${R}/font`,
    part: "word/fonts/font1.odttf",
    contentType: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
    binary: true,
};

/** Relationship spec for elements whose required r: attributes must resolve. */
const REL_FOR_ELEMENT: Record<string, RelSpec> = {
    footerReference: {
        type: `${R}/footer`,
        part: "word/footer1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    },
    headerReference: {
        type: `${R}/header`,
        part: "word/header1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    },
    embedBold: FONT_REL,
    embedBoldItalic: FONT_REL,
    embedItalic: FONT_REL,
    embedRegular: FONT_REL,
    embedSystemFonts: FONT_REL,
    embedTrueTypeFonts: FONT_REL,
    printerSettings: {
        type: `${R}/printerSettings`,
        part: "word/printerSettings.bin",
        binary: true,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings",
    },
    attachedTemplate: {
        type: `${R}/attachedTemplate`,
        part: "word/attachedTemplate.dotx",
        content: XML_DECL + `<w:document xmlns:w="${W}"><w:body/></w:document>\n`,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    },
    // ISO/IEC 29500: contentPart references its part via the customXml
    // relationship type; longDesc via hyperlink.
    contentPart: {
        type: `${R}/customXml`,
        part: "word/customXml/item1.xml",
        content:
            XML_DECL +
            `<ds:item xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{12345678-1234-1234-1234-123456789012}"/>\n`,
        contentType: "application/vnd.openxmlformats-officedocument.customXml+xml",
    },
    dataSource: { type: `${R}/mailMergeDataSource`, part: "word/dataSource.bin", binary: true, contentType: "application/octet-stream" },
    headerSource: {
        type: `${R}/mailMergeHeaderSource`,
        part: "word/headerSource.bin",
        binary: true,
        contentType: "application/octet-stream",
    },
    longDesc: { type: `${R}/hyperlink` },
    sourceFileName: { type: `${R}/sourceFileName`, part: "word/sourceFileName.html", binary: true, contentType: "text/html" },
    movie: { type: `${R}/video`, part: "word/movie.bin", binary: true, contentType: "video/mp4" },
    objectEmbed: { type: `${R}/oleObject`, part: "word/oleObject1.bin", binary: true, contentType: "application/octet-stream" },
    objectLink: { type: `${R}/oleObjectLink`, part: "word/oleObjectLink.bin", binary: true, contentType: "application/octet-stream" },
    src: { type: `${R}/image`, part: "word/image1.png", binary: true, contentType: "image/png" },
    subDoc: {
        type: `${R}/subDocument`,
        part: "word/subDoc.docx",
        content: XML_DECL + `<w:document xmlns:w="${W}"><w:body/></w:document>\n`,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    },
};

const COMMENT_MARKERS = `<w:p><w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;

const commentsPart = (): string =>
    XML_DECL + `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="x"><w:p/></w:comment></w:comments>\n`;

const COMMENT_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/** Elements with hand-written minimal fixtures (semantic checks need peers/parts beyond pure XSD). */
const SPECIAL: Record<string, () => FixtureParts> = {
    commentRangeStart: () => ({
        files: { "word/document.xml": minimalDocument(`<w:body>${COMMENT_MARKERS}</w:body>`), "word/comments.xml": commentsPart() },
        overrides: [
            { part: "/word/document.xml", contentType: MAIN_DOC_CT },
            { part: "/word/comments.xml", contentType: COMMENT_CT },
        ],
        defaults: [],
        partRels: { "word/document.xml": [{ id: "rId1", type: `${R}/comments`, target: "comments.xml" }] },
    }),
    commentRangeEnd: () => SPECIAL.commentRangeStart(),
    commentReference: () => SPECIAL.commentRangeStart(),
    footnoteReference: () => ({
        files: {
            "word/document.xml": minimalDocument(`<w:body><w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body>`),
            "word/footnotes.xml":
                XML_DECL +
                `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>\n`,
        },
        overrides: [
            { part: "/word/document.xml", contentType: MAIN_DOC_CT },
            { part: "/word/footnotes.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" },
        ],
        defaults: [],
        partRels: { "word/document.xml": [{ id: "rId1", type: `${R}/footnotes`, target: "footnotes.xml" }] },
    }),
    endnoteReference: () => ({
        files: {
            "word/document.xml": minimalDocument(`<w:body><w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:body>`),
            "word/endnotes.xml":
                XML_DECL +
                `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p/></w:endnote><w:endnote w:id="1"><w:p/></w:endnote></w:endnotes>\n`,
        },
        overrides: [
            { part: "/word/document.xml", contentType: MAIN_DOC_CT },
            { part: "/word/endnotes.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" },
        ],
        defaults: [],
        partRels: { "word/document.xml": [{ id: "rId1", type: `${R}/endnotes`, target: "endnotes.xml" }] },
    }),
    txbxContent: () => {
        const p = new Prefixes();
        p.of(W);
        p.of("urn:schemas-microsoft-com:vml");
        return {
            files: {
                "word/document.xml":
                    XML_DECL +
                    `<w:document${p.declarations()}><w:body><w:p><w:r><w:pict><v:shape id="s1"><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:body></w:document>\n`,
            },
            overrides: [{ part: "/word/document.xml", contentType: MAIN_DOC_CT }],
            defaults: [],
            partRels: {},
        };
    },
};

/** Add relationship sidecars for every r:id/r:embed/r:link reference found in the built XML parts. */
const fixupRelationshipRefs = (fx: FixtureParts): void => {
    const refRe = /<w:(\w+)[^>]*?\sr:(?:id|embed|link)="([^"]+)"/g;
    for (const [partPath, content] of Object.entries(fx.files)) {
        if (typeof content !== "string") continue;
        let m: RegExpExecArray | null;
        while ((m = refRe.exec(content)) !== null) {
            const [, elName, refId] = m;
            const rels = (fx.partRels[partPath] ??= []);
            if (rels.some((r) => r.id === refId)) continue;
            const spec = REL_FOR_ELEMENT[elName];
            if (spec?.part) {
                const target = path.posix.relative(path.posix.dirname(partPath), spec.part);
                if (!fx.files[spec.part]) {
                    if (spec.content) {
                        fx.files[spec.part] = spec.content;
                        fx.overrides.push({ part: "/" + spec.part, contentType: spec.contentType! });
                    } else if (spec.binary) {
                        fx.files[spec.part] = new Uint8Array([0x00, 0x01, 0x00, 0x00]);
                        fx.defaults.push({ ext: path.posix.extname(spec.part).slice(1), contentType: spec.contentType! });
                    } else {
                        const rootName = spec.part.includes("footer") ? "ftr" : "hdr";
                        fx.files[spec.part] = XML_DECL + `<w:${rootName} xmlns:w="${W}"/>\n`;
                        fx.overrides.push({ part: "/" + spec.part, contentType: spec.contentType! });
                    }
                }
                rels.push({ id: refId, type: spec.type, target });
            } else if (elName === "hyperlink" || elName === "longDesc") {
                // Only w:hyperlink and w:longDesc use the hyperlink relationship type.
                rels.push({ id: refId, type: `${R}/hyperlink`, target: "https://example.com/", external: true });
            } else {
                // Any other r:id/r:embed/r:link reference gets a typed internal
                // relationship to a placeholder part — never a hyperlink.
                const part = `word/${elName}.bin`;
                if (!fx.files[part]) {
                    fx.files[part] = new Uint8Array([0x00, 0x01, 0x00, 0x00]);
                    fx.defaults.push({ ext: "bin", contentType: "application/octet-stream" });
                }
                rels.push({ id: refId, type: `${R}/${elName}`, target: path.posix.relative(path.posix.dirname(partPath), part) });
            }
        }
    }
};

const buildFixture = (model: SchemaModel, paths: PathIndex, elementName: string): FixtureParts => {
    const target = key(W, elementName);
    if (SPECIAL[elementName]) return SPECIAL[elementName]();

    const found = paths.find(model, target);
    if (!found) throw new Error(`no path from any part root to ${target}`);
    const { root, path: pathArr } = found;
    if (process.env.REPAIR_DEBUG) console.log(`path for ${elementName}: [${root}] ${pathArr.map((p) => splitKey(p.id)[1]).join(" > ")}`);

    const fx: FixtureParts = {
        files: {},
        overrides: [{ part: "/word/document.xml", contentType: MAIN_DOC_CT }],
        defaults: [],
        partRels: {},
    };
    const xml = renderChain(model, pathArr);
    const rootInfo = ROOT_PARTS[root];
    if (rootInfo === null || root === "document") {
        fx.files["word/document.xml"] = XML_DECL + xml + "\n";
    } else {
        const body = root === "comments" && xml.includes("<w:comment ") ? `<w:body>${COMMENT_MARKERS}</w:body>` : "<w:body/>";
        fx.files["word/document.xml"] = minimalDocument(body);
        fx.files[rootInfo.part] = XML_DECL + xml + "\n";
        fx.overrides.push({ part: "/" + rootInfo.part, contentType: rootInfo.contentType });
        if (root === "recipients") {
            // Mail Merge Recipient Data is linked from the settings part's
            // sidecar (…/relationships/recipientData), not from document.xml.
            fx.files["word/settings.xml"] = XML_DECL + `<w:settings xmlns:w="${W}"/>` + "\n";
            fx.overrides.push({
                part: "/word/settings.xml",
                contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            });
            fx.partRels["word/settings.xml"] = [{ id: "rId1", type: rootInfo.relType, target: path.posix.relative("word", rootInfo.part) }];
            fx.partRels["word/document.xml"] = [{ id: "rId2", type: `${R}/settings`, target: "settings.xml" }];
        } else {
            fx.partRels["word/document.xml"] = [{ id: "rId1", type: rootInfo.relType, target: path.posix.relative("word", rootInfo.part) }];
        }
    }
    fixupRelationshipRefs(fx);
    return fx;
};

const packFixture = async (fx: FixtureParts): Promise<Buffer> => {
    const zip = new JSZip();
    const defaults = [
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
        `<Default Extension="xml" ContentType="application/xml"/>`,
        ...fx.defaults.map((d) => `<Default Extension="${d.ext}" ContentType="${d.contentType}"/>`),
    ].join("");
    const overrides = fx.overrides.map((o) => `<Override PartName="${o.part}" ContentType="${o.contentType}"/>`).join("");
    zip.file(
        "[Content_Types].xml",
        XML_DECL + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>\n`,
    );
    zip.file(
        "_rels/.rels",
        XML_DECL +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>\n`,
    );
    for (const [partPath, rels] of Object.entries(fx.partRels)) {
        if (!rels.length) continue;
        const dir = path.posix.dirname(partPath);
        const base = path.posix.basename(partPath);
        const relsXml = rels
            .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ""}/>`)
            .join("");
        zip.file(
            `${dir}/_rels/${base}.rels`,
            XML_DECL + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsXml}</Relationships>\n`,
        );
    }
    for (const [name, content] of Object.entries(fx.files)) zip.file(name, content);
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
};

const collectElementNames = (xsdPath: string): string[] => {
    const root = parseXmlBytes(readFileSync(xsdPath)).documentElement!;
    const names = new Set<string>();
    const walk = (element: Element): void => {
        for (const child of Array.from(element.childNodes)) {
            if (child.nodeType !== 1) continue;
            const declaration = child as Element;
            const name = declaration.namespaceURI === XSD && declaration.localName === "element" ? declaration.getAttribute("name") : null;
            if (name) names.add(name);
            walk(declaration);
        }
    };
    walk(root);
    return [...names].sort(compareCodeUnits);
};

interface ElementResult {
    element: string;
    file: string;
    sha256: string;
    profile: Profile;
    valid: boolean;
    issueCount: number;
    errorCodes: string[];
    warningCodes: string[];
    errorMessages?: string[];
    /** True when the row was written by the repair script (not the bare generator). */
    repaired?: boolean;
}

runCli(import.meta.url, async () => {
    const program = new Command();
    program
        .option("--out <dir>", "fixture directory", path.join(REPO, "word-element-fixtures"))
        .option("--results <file>", "JSONL results path", "")
        .option("--profile <profile>", "validation profile", "lenient")
        .option("--limit <n>", "stop after n elements", "")
        .option("--only <names>", "comma-separated element names to process", "")
        .option("--resume", "skip elements already recorded in the results file", false);
    program.parse(process.argv);
    const opts = program.opts<{ out: string; results: string; profile: Profile; limit: string; only: string; resume: boolean }>();
    // Validate --limit before any output file is created or cleared; an
    // explicit --limit 0 is a full no-op that touches nothing.
    let limit: number | null;
    try {
        limit = parseLimitOption(opts.limit, limitFlagPassed(process.argv));
    } catch (err) {
        console.error((err as Error).message);
        return 1;
    }
    if (limit === 0) {
        console.log("--limit 0: nothing to do; outputs left untouched");
        return 0;
    }
    let outDir: string;
    let resultsPath: string;
    try {
        outDir = resolveInside(REPO, opts.out, "--out");
        resultsPath = opts.results ? resolveInside(REPO, opts.results, "--results") : path.join(outDir, "manifest.jsonl");
    } catch (err) {
        console.error((err as Error).message);
        return 1;
    }
    mkdirSync(outDir, { recursive: true });

    const model = new SchemaModel(WML_XSD);
    const adj = buildAdjacency(model);
    const paths = new PathIndex(model, adj);
    if (process.env.REPAIR_DEBUG) {
        console.log("files:", model.files.length, "elements:", model.elements.size, "adj:", adj.size);
    }
    const all = collectElementNames(WML_XSD);
    // A partial (--only) or resuming run updates only the selected rows and
    // preserves rows for every other fixture; a full non-resume run replaces
    // the manifest wholesale.
    const keepExisting = opts.resume || Boolean(opts.only) || limit !== null;
    // A full run (no --resume/--only/--limit) starts from an empty map so stale
    // rows for elements no longer in the schema cannot be restored.
    const rowsByElement = keepExisting ? loadManifestRows<ElementResult>(resultsPath) : new Map<string, ElementResult>();
    if (!keepExisting) writeFileSync(resultsPath, "");
    const persistResults = (): void => persistManifestRows(resultsPath, rowsByElement);
    // Resume skips a fixture only when its row is marked as a successful repair
    // for the requested profile and the fixture on disk still matches the
    // recorded hash; generator rows, build-failed rows, and stale rows are
    // reprocessed and replaced on success.
    const done = new Set<string>();
    if (opts.resume) {
        for (const row of rowsByElement.values()) {
            if (row.repaired === true && row.profile === opts.profile && row.valid && fixtureMatchesRow(outDir, row)) done.add(row.element);
        }
    }
    let todo = all.filter((name) => !done.has(name));
    if (opts.only) {
        const wanted = new Set(
            opts.only
                .split(",")
                .map((s) => s.trim())
                .filter(Boolean),
        );
        todo = todo.filter((name) => wanted.has(name));
    }
    const elements = limit === null ? todo : todo.slice(0, limit);
    console.log(`repairing ${elements.length} of ${all.length} elements (${done.size} done)`);
    const validator = new Validate({ profile: opts.profile, suffix: ".docx" });
    let validCount = 0;
    let buildFailures = 0;
    for (const [index, element] of elements.entries()) {
        let row: ElementResult;
        let wasValid = false;
        try {
            const fx = buildFixture(model, paths, element);
            const bytes = await packFixture(fx);
            const result = await validator.run(bytes);
            const codes = (severity: "error" | "warning"): string[] =>
                [...new Set(result.issues.filter((issue) => issue.severity === severity).map((issue) => issue.code ?? "unknown"))].sort(compareCodeUnits);
            const file = `${element}.docx`;
            writeFileSync(path.join(outDir, file), bytes);
            wasValid = result.valid;
            row = {
                element,
                file,
                sha256: createHash("sha256").update(bytes).digest("hex"),
                profile: opts.profile,
                valid: result.valid,
                issueCount: result.issues.length,
                errorCodes: codes("error"),
                warningCodes: codes("warning"),
                repaired: true,
                ...(result.valid
                    ? {}
                    : { errorMessages: result.issues.filter((i) => i.severity === "error").map((i) => i.message.slice(0, 300)) }),
            };
        } catch (err) {
            buildFailures++;
            row = {
                element,
                file: `${element}.docx`,
                sha256: "",
                profile: opts.profile,
                valid: false,
                issueCount: 0,
                errorCodes: ["build-failed"],
                warningCodes: [],
                errorMessages: [err instanceof Error ? err.message : String(err)],
            };
        }
        rowsByElement.set(element, row);
        persistResults();
        // Count as valid only after both the fixture file and the manifest row
        // have been written successfully.
        if (wasValid) validCount++;
        if ((index + 1) % 25 === 0)
            console.log(`  ${index + 1}/${elements.length} (${validCount} valid, ${buildFailures} build failures so far)`);
    }
    console.log(`done: ${elements.length} processed, ${validCount} valid, ${buildFailures} build failures; results: ${resultsPath}`);
    return buildFailures > 0 ? 1 : 0;
});
