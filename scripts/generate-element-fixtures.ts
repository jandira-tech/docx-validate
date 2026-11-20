/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Generate one minimal `.docx` fixture per element declared in
 * `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` (global and local
 * `xsd:element` declarations, de-duplicated by name), validate each fixture
 * with the package validator, and record the outcome as JSONL.
 *
 * Each fixture is a three-part package (`[Content_Types].xml`, `_rels/.rels`,
 * `word/document.xml`) whose body holds the bare element:
 * `<w:body><w:ELEMENT/></w:body>`.
 *
 * Usage:
 *   tsx scripts/generate-element-fixtures.ts [--out word-element-fixtures] [--results <jsonl>] [--profile lenient]
 */

import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import JSZip from "jszip";

import { runCli } from "../src/lib/run-cli";
import { fixtureMatchesRow, limitFlagPassed, loadManifestRows, parseLimitOption, persistManifestRows } from "./element-fixture-cli";
import { parseXmlBytes } from "../src/lib/xml-helpers";
import type { Profile } from "../src/lib/types";
import { Validate } from "../src/validate";

const XSD = "http://www.w3.org/2001/XMLSchema";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const WML_XSD = path.join(REPO, "src", "scripts", "office", "schemas", "ISO-IEC29500-4_2016", "wml.xsd");

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
    return [...names].sort((a, b) => a.localeCompare(b));
};

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
`;

const documentXml = (element: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body><w:${element}/></w:body></w:document>
`;

const buildDocx = async (element: string): Promise<Buffer> => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES);
    zip.file("_rels/.rels", ROOT_RELS);
    zip.file("word/document.xml", documentXml(element));
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
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
}

runCli(import.meta.url, async () => {
    const program = new Command();
    program
        .option("--out <dir>", "fixture output directory", path.join(REPO, "word-element-fixtures"))
        .option("--results <file>", "JSONL results path (default: <out>/manifest.jsonl)", "")
        .option("--profile <profile>", "validation profile", "lenient")
        .option("--limit <n>", "stop after n elements", "")
        .option("--resume", "skip elements already recorded in the results file", false);
    program.parse(process.argv);
    const opts = program.opts<{ out: string; results: string; profile: Profile; limit: string; resume: boolean }>();
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
    const outDir = path.resolve(opts.out);
    const resultsPath = path.resolve(opts.results || path.join(outDir, "manifest.jsonl"));
    mkdirSync(outDir, { recursive: true });

    const all = collectElementNames(WML_XSD);
    const keepExisting = opts.resume || limit !== null;
    // A full run (no --resume/--limit) starts from an empty map so stale rows
    // cannot be restored.
    const rowsByElement = keepExisting ? loadManifestRows<ElementResult>(resultsPath) : new Map<string, ElementResult>();
    if (!keepExisting) writeFileSync(resultsPath, "");
    const done = new Set<string>();
    if (opts.resume) {
        for (const row of rowsByElement.values()) {
            if (row.profile === opts.profile && row.valid && fixtureMatchesRow(outDir, row)) done.add(row.element);
        }
    }
    const todo = all.filter((name) => !done.has(name));
    const elements = limit === null ? todo : todo.slice(0, limit);
    console.log(`wml.xsd declares ${all.length} unique element names; ${done.size} already done; generating ${elements.length}`);
    const validator = new Validate({ profile: opts.profile, suffix: ".docx" });
    const results: ElementResult[] = [];
    let validCount = 0;
    for (const [index, element] of elements.entries()) {
        const bytes = await buildDocx(element);
        const result = await validator.run(bytes);
        const codes = (severity: "error" | "warning"): string[] =>
            [...new Set(result.issues.filter((issue) => issue.severity === severity).map((issue) => issue.code ?? "unknown"))].sort();
        if (result.valid) validCount++;
        const file = `${element}.docx`;
        writeFileSync(path.join(outDir, file), bytes);
        const row: ElementResult = {
            element,
            file,
            sha256: createHash("sha256").update(bytes).digest("hex"),
            profile: opts.profile,
            valid: result.valid,
            issueCount: result.issues.length,
            errorCodes: codes("error"),
            warningCodes: codes("warning"),
        };
        results.push(row);
        rowsByElement.set(element, row);
        persistManifestRows(resultsPath, rowsByElement);
        if ((index + 1) % 50 === 0) console.log(`  ${index + 1}/${elements.length} (${validCount} valid so far)`);
    }
    console.log(
        `wrote ${results.length} fixtures to ${outDir}; ${validCount} valid, ${results.length - validCount} invalid; results: ${resultsPath}`,
    );
    return 0;
});
