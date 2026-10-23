/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Task 7a integration evidence: every content-model placement check fires on
 * a real Word-rejected fixture through `validate(file, { profile:
 * "word-valid" })`, keeps firing when the XSD engine returns no diagnostic at
 * all (so the verdict no longer depends on libxml2 message wording), and its
 * hard-coded parent tables match the bundled wml.xsd content models.
 */

import { readFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import * as xsdMod from "../src/lib/xsd-validator";
import {
    RUN_PARENTS,
    RUN_PROPERTIES_PARENTS,
    UNDERLINE_PARENTS,
    wordRunMisplacedCheck,
    wordRunPropertiesMisplacedCheck,
    wordUnderlineMisplacedCheck,
} from "../src/scripts/office/checks/word-run-placement";
import { validate, type ValidateRunResult } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");
const REPO = path.resolve(HERE, "..");
const SCHEMAS = path.join(REPO, "src/scripts/office/schemas/ISO-IEC29500-4_2016");
const XSD = "http://www.w3.org/2001/XMLSchema";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math";

const manifest = JSON.parse(readFileSync(path.join(HERE, "fixtures-all.manifest.json"), "utf8")) as {
    entries: { relativePath: string; word?: string }[];
};

const wordVerdict = (relativePath: string): string | undefined =>
    manifest.entries.find((entry) => entry.relativePath === relativePath)?.word;

const errorCodes = (result: ValidateRunResult): string[] =>
    result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "");

const rejectingNoXsd = async (file: string, code: string): Promise<void> => {
    const factory = vi.spyOn(xsdMod, "createXsdValidator").mockResolvedValue({ validate: async () => [] });
    try {
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(code);
        expect(result.xsdWalk?.started).toBe(false);
    } finally {
        factory.mockRestore();
    }
};

// --- XSD pinning harness -----------------------------------------------------

const loadSchema = async (file: string): Promise<Document> => parseXml(await readFile(path.join(SCHEMAS, file), "utf8"));

interface ModelIndex {
    groups: Map<string, Element>;
    types: Map<string, Element>;
    elementTypes: Map<string, Set<string>>;
}

const buildIndex = (schema: Document): ModelIndex => {
    const groups = new Map<string, Element>();
    for (const group of getElementsByTagNameNSAll(schema, XSD, "group")) {
        if (group.getAttribute("name")) {
            groups.set(group.getAttribute("name")!, group);
        }
    }
    const types = new Map<string, Element>();
    for (const type of getElementsByTagNameNSAll(schema, XSD, "complexType")) {
        if (type.getAttribute("name")) {
            types.set(type.getAttribute("name")!, type);
        }
    }
    const elementTypes = new Map<string, Set<string>>();
    for (const element of getElementsByTagNameNSAll(schema, XSD, "element")) {
        const name = element.getAttribute("name");
        const type = element.getAttribute("type");
        if (!name || !type) {
            continue;
        }
        const names = elementTypes.get(name) ?? new Set<string>();
        names.add(type);
        elementTypes.set(name, names);
    }
    return { groups, types, elementTypes };
};

const elementNamesUnder = (node: Element, index: ModelIndex, seen: Set<string>): Set<string> => {
    const out = new Set<string>();
    const visit = (current: Element): void => {
        for (const child of Array.from(current.childNodes).filter((c): c is Element => c.nodeType === 1)) {
            if (child.namespaceURI !== XSD) {
                continue;
            }
            const local = child.localName ?? "";
            if (local === "element" && child.getAttribute("name")) {
                out.add(child.getAttribute("name")!);
            } else if (local === "group") {
                const ref = child.getAttribute("ref")?.split(":").pop();
                if (ref && !seen.has(ref) && index.groups.get(ref)) {
                    seen.add(ref);
                    for (const name of elementNamesUnder(index.groups.get(ref)!, index, seen)) {
                        out.add(name);
                    }
                }
            } else if (local === "extension" || local === "restriction") {
                const base = child.getAttribute("base");
                if (base && !seen.has(base) && index.types.get(base)) {
                    seen.add(base);
                    for (const name of elementNamesUnder(index.types.get(base)!, index, seen)) {
                        out.add(name);
                    }
                }
                visit(child);
            } else if (local === "sequence" || local === "choice" || local === "all" || local === "complexContent") {
                visit(child);
            }
        }
    };
    visit(node);
    return out;
};

/** Parent element names whose (transitively resolved) wml.xsd model admits `child`. */
const legalParentsInWml = (index: ModelIndex, child: string): Set<string> => {
    const memo = new Map<string, Set<string>>();
    const contentOf = (typeName: string): Set<string> => {
        const cached = memo.get(typeName);
        if (cached) {
            return cached;
        }
        const type = index.types.get(typeName);
        const names = type ? elementNamesUnder(type, index, new Set([typeName])) : new Set<string>();
        memo.set(typeName, names);
        return names;
    };
    const parents = new Set<string>();
    for (const [element, typeNames] of index.elementTypes) {
        for (const typeName of typeNames) {
            if (contentOf(typeName).has(child)) {
                parents.add(element);
                break;
            }
        }
    }
    return parents;
};

describe("content-model tables pinned to the bundled wml.xsd", () => {
    it("run, run-properties, and underline parents match the schema's content models", async () => {
        const index = buildIndex(await loadSchema("wml.xsd"));
        expect(legalParentsInWml(index, "r")).toEqual(RUN_PARENTS);
        expect(legalParentsInWml(index, "rPr")).toEqual(RUN_PROPERTIES_PARENTS[W]);
        expect(legalParentsInWml(index, "u")).toEqual(UNDERLINE_PARENTS);
    });

    it("the math r and ctrlPr parents come from shared-math.xsd referencing w:EG_RPr", async () => {
        const math = await readFile(path.join(SCHEMAS, "shared-math.xsd"), "utf8");
        expect(math).toMatch(/<xsd:complexType name="CT_CtrlPr">[\s\S]*?<xsd:group ref="w:EG_RPrMath" minOccurs="0"\/>/);
        expect(math).toMatch(/<xsd:complexType name="CT_R">[\s\S]*?<xsd:group ref="w:EG_RPr" minOccurs="0"\/>/);
        expect(RUN_PROPERTIES_PARENTS[MATH]).toEqual(new Set(["ctrlPr", "r"]));
    });
});

describe("run and run-property placement Word evidence", () => {
    it("rejects the misplaced-run fixture and its OPEN_ERROR probed twins", async () => {
        expect(wordVerdict("broken/word-rejects-misplaced-run.docx")).toBe("unknown");
        const probes = (await readFile(path.join(REPO, "word-invalid-fixtures/manifest.jsonl"), "utf8"))
            .split("\n")
            .filter((line) => line.includes("critic-variations"))
            .map((line) => JSON.parse(line) as { file: string; verdict: string });
        expect(probes).toHaveLength(2);
        for (const probe of probes) {
            expect(probe.verdict).toBe("OPEN_ERROR");
        }
        const result = await validate(path.join(FIXTURES, "broken/word-rejects-misplaced-run.docx"), { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordRunMisplacedCheck.id);
    });

    it("rejects the duplicated run properties of 5Errors.docx without any XSD diagnostic", async () => {
        expect(wordVerdict("external/open-xml-sdk/5Errors.docx")).toBe("unreadable-content-warning");
        const file = path.join(FIXTURES, "external/open-xml-sdk/5Errors.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordRunPropertiesMisplacedCheck.id);
        await rejectingNoXsd(file, wordRunPropertiesMisplacedCheck.id);
    });

    it("rejects the paragraph-level underline of invalid-underline.docx without any XSD diagnostic", async () => {
        expect(wordVerdict("external/superdoc/super-editor/invalid-underline.docx")).toBe("unreadable-content-warning");
        const file = path.join(FIXTURES, "external/superdoc/super-editor/invalid-underline.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordUnderlineMisplacedCheck.id);
        await rejectingNoXsd(file, wordUnderlineMisplacedCheck.id);
    });

    it("accepts the tolerated placements on a Word-clean control document", async () => {
        const bytes = await readFile(path.join(FIXTURES, "working/word-tolerated-misplaced-pgsz.docx"));
        const ctx = await ValidationContext.fromArchive(new Uint8Array(bytes));
        expect(await wordRunMisplacedCheck.run(ctx)).toEqual([]);
        expect(await wordRunPropertiesMisplacedCheck.run(ctx)).toEqual([]);
        expect(await wordUnderlineMisplacedCheck.run(ctx)).toEqual([]);
    });
});
