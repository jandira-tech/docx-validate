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
    PARAGRAPH_PROPERTIES_PARENTS,
    wordParagraphPropertiesMisplacedCheck,
} from "../src/scripts/office/checks/word-paragraph-properties";
import {
    RUN_PARENTS,
    RUN_PROPERTIES_PARENTS,
    UNDERLINE_PARENTS,
    wordRunMisplacedCheck,
    wordRunPropertiesMisplacedCheck,
    wordUnderlineMisplacedCheck,
} from "../src/scripts/office/checks/word-run-placement";
import { SECTION_PROPERTIES_PARENTS, wordSectionPropertiesMisplacedCheck } from "../src/scripts/office/checks/word-section-properties";
import {
    TABLE_BORDERS_PARENTS,
    TABLE_CELL_SPACING_PARENTS,
    TBLPRBASE_ORDER,
    wordTablePropertiesMisplacedCheck,
} from "../src/scripts/office/checks/word-table-properties";
import { wordLinkWithoutNamespaceCheck, wordUnknownElementCheck } from "../src/scripts/office/checks/word-unknown-elements";
import { WML_ELEMENT_NAMES } from "../src/scripts/office/checks/wml-vocabulary";
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

describe("content-model tables pinned to the bundled wml.xsd", { timeout: 30_000 }, () => {
    it("run, run-properties, underline, and paragraph-properties parents match the schema's content models", async () => {
        const index = buildIndex(await loadSchema("wml.xsd"));
        expect(legalParentsInWml(index, "r")).toEqual(RUN_PARENTS);
        expect(legalParentsInWml(index, "rPr")).toEqual(RUN_PROPERTIES_PARENTS[W]);
        expect(legalParentsInWml(index, "u")).toEqual(UNDERLINE_PARENTS);
        expect(legalParentsInWml(index, "pPr")).toEqual(PARAGRAPH_PROPERTIES_PARENTS);
        expect(legalParentsInWml(index, "sectPr")).toEqual(SECTION_PROPERTIES_PARENTS);
        // w:tcPr is deliberately WIDER than wml.xsd: Word opened the jubarte
        // lane specimens that carry w:tblBorders under w:tcPr.
        const borderParents = legalParentsInWml(index, "tblBorders");
        expect(borderParents).toEqual(new Set(["tblPr", "tblPrEx"]));
        expect(TABLE_BORDERS_PARENTS).toEqual(new Set([...borderParents, "tcPr"]));
        expect(legalParentsInWml(index, "tblCellSpacing")).toEqual(TABLE_CELL_SPACING_PARENTS);
    });

    it("the tblPr child order matches CT_TblPrBase", async () => {
        const schema = await loadSchema("wml.xsd");
        const type = getElementsByTagNameNSAll(schema, XSD, "complexType").find((entry) => entry.getAttribute("name") === "CT_TblPrBase");
        expect(type).toBeDefined();
        const children = getElementsByTagNameNSAll(type!, XSD, "element")
            .map((element) => element.getAttribute("name"))
            .filter((name): name is string => Boolean(name));
        expect(children).toEqual([...TBLPRBASE_ORDER]);
    });

    it("the WordprocessingML element vocabulary equals every xsd:element declaration in wml.xsd", async () => {
        const schema = await loadSchema("wml.xsd");
        const declared = new Set(
            getElementsByTagNameNSAll(schema, XSD, "element")
                .map((element) => element.getAttribute("name"))
                .filter((name): name is string => Boolean(name)),
        );
        expect(WML_ELEMENT_NAMES).toEqual(declared);
        expect(WML_ELEMENT_NAMES.has("p2")).toBe(false);
        expect(WML_ELEMENT_NAMES.has("p")).toBe(true);
        // No other bundled schema targets the WordprocessingML namespace, so
        // this vocabulary is the whole namespace.
        const { readdir } = await import("node:fs/promises");
        const schemaDir = path.join(REPO, "src/scripts/office/schemas");
        const others: string[] = [];
        const walk = async (dir: string): Promise<void> => {
            for (const entry of await readdir(dir, { withFileTypes: true })) {
                const full = path.join(dir, entry.name);
                if (entry.isDirectory()) {
                    await walk(full);
                } else if (entry.name.endsWith(".xsd") && !full.includes("ISO-IEC29500-4_2016/wml.xsd")) {
                    others.push(await readFile(full, "utf8"));
                }
            }
        };
        await walk(schemaDir);
        for (const source of others) {
            const target = /targetNamespace="([^"]*)"/.exec(source)?.[1];
            expect(target, "no other schema may declare elements in the WordprocessingML namespace").not.toBe(W);
        }
    });

    it("the math r and ctrlPr parents come from shared-math.xsd referencing w:EG_RPr", async () => {
        const math = await readFile(path.join(SCHEMAS, "shared-math.xsd"), "utf8");
        expect(math).toMatch(/<xsd:complexType name="CT_CtrlPr">[\s\S]*?<xsd:group ref="w:EG_RPrMath" minOccurs="0"\/>/);
        expect(math).toMatch(/<xsd:complexType name="CT_R">[\s\S]*?<xsd:group ref="w:EG_RPr" minOccurs="0"\/>/);
        expect(RUN_PROPERTIES_PARENTS[MATH]).toEqual(new Set(["ctrlPr", "r"]));
    });
});

describe("section-properties placement Word evidence", () => {
    it("rejects a synthetically misplaced sectPr through validate() without any XSD diagnostic", async () => {
        // No vendored fixture isolates a misplaced w:sectPr; the rejection
        // class is the corpus-probe one documented in word-invalid-fixtures/
        // README.md. Build the single defect from the minimal package.
        const { buildDocx, MINIMAL_PARTS } = await import("./support/minimal-docx");
        const { writeFileSync } = await import("node:fs");
        const { withTempDir } = await import("../src/lib/run-cli");
        const defective = MINIMAL_PARTS["word/document.xml"]!.replace(
            "<w:sectPr>",
            '<w:p><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr></w:p><w:sectPr>',
        );
        await withTempDir(async (dir) => {
            const file = path.join(dir, "misplaced-sectpr.docx");
            writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": defective }));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errorCodes(result)).toContain(wordSectionPropertiesMisplacedCheck.id);
            await rejectingNoXsd(file, wordSectionPropertiesMisplacedCheck.id);
        });
        expect(await wordSectionPropertiesMisplacedCheck.run(await ValidationContext.fromEntries(MINIMAL_PARTS))).toEqual([]);
    });
});

describe("table-properties placement Word evidence", { timeout: 120_000 }, () => {
    const LANE = path.join(REPO, "tests/lane/jubarte-0.13.0");

    it("rejects the reordered tblBorders of table_tests_v2.docx without any XSD diagnostic", async () => {
        const probe = (await readFile(path.join(REPO, "word-invalid-fixtures/manifest.jsonl"), "utf8"))
            .split("\n")
            .find((line) => line.includes("tagged_tests__table_tests_v2"));
        expect(probe).toMatch(/"verdict": "UNREADABLE_CONTENT"/);
        const file = path.join(FIXTURES, "external/docx-parser-converter/tagged_tests/table_tests_v2.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordTablePropertiesMisplacedCheck.id);
        await rejectingNoXsd(file, wordTablePropertiesMisplacedCheck.id);
    });

    it("rejects the reordered tblCellSpacing fixture and keeps the message-matched heuristic redundant there", async () => {
        const file = path.join(FIXTURES, "broken/word-rejects-table-cellspacing.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordTablePropertiesMisplacedCheck.id);
    });

    it("tolerates the Word-opened jubarte lane specimens whose tblBorders sit under tcPr, and the 1993 styles-part control", async () => {
        // The lane manifest records word_opened_and_exported=True for every
        // k07/k08/k09 specimen; word-valid must not fail them any more (the
        // tblBorders xsd-error was a message-matched false positive).
        const rows = (await readFile(path.join(LANE, "manifest.csv"), "utf8")).split("\n");
        const opened = rows.filter((row) => row.includes("True")).map((row) => row.match(/fixtures\/(k\d+_\d+\.redline\.docx)/)?.[1] ?? "");
        for (const name of [
            "k07_1.redline.docx",
            "k07_2.redline.docx",
            "k08_1.redline.docx",
            "k08_2.redline.docx",
            "k09_1.redline.docx",
            "k09_2.redline.docx",
        ]) {
            expect(opened, `lane manifest must record ${name} as opened by Word`).toContain(name);
            const bytes = await readFile(path.join(LANE, name));
            const ctx = await ValidationContext.fromArchive(new Uint8Array(bytes));
            expect(await wordTablePropertiesMisplacedCheck.run(ctx), name).toEqual([]);
            const result = await validate(path.join(LANE, name), { profile: "word-valid" });
            expect(
                result.issues.filter((issue) => issue.severity === "error"),
                `${name} must carry no word-valid error, got ${errorCodes(result).join(",")}`,
            ).toEqual([]);
            expect(result.valid).toBe(true);
        }
        const repaired = path.join(FIXTURES, "working/package.word-row-without-cell-repaired.docx");
        expect((await validate(repaired, { profile: "word-valid" })).valid).toBe(true);
    });
});

describe("unknown-element and namespace-less link Word evidence", { timeout: 120_000 }, () => {
    it("rejects the unknown w:p2 block of UnknownElement.docx without any XSD diagnostic", async () => {
        expect(wordVerdict("external/open-xml-sdk/UnknownElement.docx")).toBe("open-error");
        const file = path.join(FIXTURES, "external/open-xml-sdk/UnknownElement.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordUnknownElementCheck.id);
        await rejectingNoXsd(file, wordUnknownElementCheck.id);
    });

    it("rejects the namespace-less link of hyperlink-exported.docx without any XSD diagnostic", async () => {
        expect(wordVerdict("external/superdoc/super-editor/hyperlink-exported.docx")).toBe("open-error");
        const file = path.join(FIXTURES, "external/superdoc/super-editor/hyperlink-exported.docx");
        const result = await validate(file, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(errorCodes(result)).toContain(wordLinkWithoutNamespaceCheck.id);
        await rejectingNoXsd(file, wordLinkWithoutNamespaceCheck.id);
    });

    it("tolerates unknown names in property containers and the namespaced w:link style child", async () => {
        // annotations_import_2.docx is manifest word: clean-open while carrying
        // w:fieldType/w:fieldTypeShort/w:fieldColor in w:sdtPr; the misplaced-
        // link control carries the namespaced w:link style child.
        expect(wordVerdict("external/superdoc/super-editor/annotations_import_2.docx")).toBe("clean-open");
        const sdtFields = await ValidationContext.fromArchive(
            new Uint8Array(await readFile(path.join(FIXTURES, "external/superdoc/super-editor/annotations_import_2.docx"))),
        );
        expect(await wordUnknownElementCheck.run(sdtFields)).toEqual([]);
        expect(
            (await validate(path.join(FIXTURES, "external/superdoc/super-editor/annotations_import_2.docx"), { profile: "word-valid" }))
                .valid,
        ).toBe(true);
        const linkControl = await ValidationContext.fromArchive(
            new Uint8Array(await readFile(path.join(FIXTURES, "working/word-tolerated-misplaced-link.docx"))),
        );
        expect(await wordLinkWithoutNamespaceCheck.run(linkControl)).toEqual([]);
        expect(await wordUnknownElementCheck.run(linkControl)).toEqual([]);
    });
});

describe("run and run-property placement Word evidence", { timeout: 120_000 }, () => {
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
        expect(errorCodes(result)).toContain(wordParagraphPropertiesMisplacedCheck.id);
        await rejectingNoXsd(file, wordParagraphPropertiesMisplacedCheck.id);
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
        expect(await wordParagraphPropertiesMisplacedCheck.run(ctx)).toEqual([]);
        const duplicate = await readFile(path.join(FIXTURES, "working/word-tolerated-duplicate-ppr.docx"));
        const duplicateCtx = await ValidationContext.fromArchive(new Uint8Array(duplicate));
        expect(await wordParagraphPropertiesMisplacedCheck.run(duplicateCtx)).toEqual([]);
        expect((await validate(path.join(FIXTURES, "working/word-tolerated-duplicate-ppr.docx"), { profile: "word-valid" })).valid).toBe(
            true,
        );
    });
});
