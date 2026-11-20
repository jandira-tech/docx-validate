/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Copy every `word-element-fixtures/<element>.docx` into
 * `word-element-invalid-fixtures/<element>.docx` with the absolute minimal
 * mutation that makes it FAIL the validation rule specific to that element:
 * the element itself is mutated (a required attribute or required child is
 * removed, else an undeclared attribute is added), so the XSD error names the
 * element — while the fixture still contains the element in its proper
 * context, i.e. stays representative of the element's XML rule.
 *
 * A mutation is accepted only if the lenient run is invalid AND at least one
 * error message names the target element (`{...}name`), so the failure is the
 * element's own rule, not an unrelated defect.
 *
 * Usage:
 *   tsx scripts/break-element-fixtures.ts [--src word-element-fixtures] [--out word-element-invalid-fixtures] [--results <jsonl>] [--limit n] [--only a,b] [--resume]
 */

import { createHash } from "node:crypto";
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import JSZip from "jszip";

import { runCli } from "../src/lib/run-cli";
import { fixtureMatchesRow, limitFlagPassed, loadManifestRows, parseLimitOption, persistManifestRows } from "./element-fixture-cli";
import { parseXml, serializeXml } from "../src/lib/xml-helpers";
import type { Profile } from "../src/lib/types";
import { Validate } from "../src/validate";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");

interface LoadedFixture {
    order: string[];
    files: Record<string, string>;
}

const loadFixture = async (file: string): Promise<LoadedFixture> => {
    const zip = await JSZip.loadAsync(readFileSync(file));
    const order: string[] = [];
    const files: Record<string, string> = {};
    for (const name of Object.keys(zip.files)) {
        order.push(name);
        files[name] = await zip.files[name].async("string");
    }
    return { order, files };
};

const packFixture = async (fx: LoadedFixture): Promise<Buffer> => {
    const zip = new JSZip();
    for (const name of fx.order) zip.file(name, fx.files[name]);
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
};

/** The target element occurrence: the last `w:name` in document order across the XML parts. */
const findTarget = (files: Record<string, string>, element: string): { file: string; el: Element } | null => {
    let best: { file: string; el: Element } | null = null;
    for (const [name, content] of Object.entries(files)) {
        if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
        const doc = parseXml(content);
        const matches = doc.getElementsByTagNameNS(W, element);
        for (let i = 0; i < matches.length; i++) best = { file: name, el: matches[i] };
    }
    return best;
};

type Mutation = (el: Element, doc: Document) => void;

/** Candidate mutations, most "own-rule" first. */
const mutations = (): { label: string; apply: Mutation }[] => [
    {
        label: "remove-first-attribute",
        apply: (el) => {
            const attr = el.attributes.item(0);
            if (attr) el.removeAttributeNode(attr);
        },
    },
    {
        label: "remove-first-child",
        apply: (el) => {
            const child = Array.from(el.childNodes).find((c) => c.nodeType === 1);
            if (child) el.removeChild(child);
        },
    },
    {
        label: "add-undeclared-attribute",
        apply: (el) => el.setAttributeNS(W, "bogus", "1"),
    },
    {
        label: "add-disallowed-child",
        apply: (el, doc) => {
            el.appendChild(doc.createElementNS(W, "sectPr"));
        },
    },
];

interface ElementResult {
    element: string;
    file: string;
    sha256: string;
    profile: Profile;
    valid: boolean;
    mutation: string;
    issueCount: number;
    errorCodes: string[];
    errorMessages: string[];
}

runCli(import.meta.url, async () => {
    const program = new Command();
    program
        .option("--src <dir>", "valid fixture directory", path.join(REPO, "word-element-fixtures"))
        .option("--out <dir>", "broken fixture output directory", path.join(REPO, "word-element-invalid-fixtures"))
        .option("--results <file>", "JSONL results path", "")
        .option("--profile <profile>", "validation profile", "lenient")
        .option("--limit <n>", "stop after n elements", "")
        .option("--only <names>", "comma-separated element names to process", "")
        .option("--resume", "skip elements already recorded in the results file", false);
    program.parse(process.argv);
    const opts = program.opts<{
        src: string;
        out: string;
        results: string;
        profile: Profile;
        limit: string;
        only: string;
        resume: boolean;
    }>();
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
    const srcDir = path.resolve(opts.src);
    const outDir = path.resolve(opts.out);
    const resultsPath = path.resolve(opts.results || path.join(outDir, "manifest.jsonl"));
    mkdirSync(outDir, { recursive: true });

    const all = readdirSorted(srcDir);
    // A partial (--only) or resuming run updates only the selected rows and
    // preserves rows for every other fixture; a full non-resume run replaces
    // the manifest wholesale.
    const keepExisting = opts.resume || Boolean(opts.only) || limit !== null;
    // A full run (no --resume/--only/--limit) starts from an empty map so stale
    // rows cannot be restored.
    const rowsByElement = keepExisting ? loadManifestRows<ElementResult>(resultsPath) : new Map<string, ElementResult>();
    if (!keepExisting) writeFileSync(resultsPath, "");
    // Resume skips a fixture only when its row records a successful break (a
    // real mutation, invalid under the requested profile) and the fixture on
    // disk still matches the recorded hash.
    const done = new Set<string>();
    if (opts.resume) {
        for (const row of rowsByElement.values()) {
            if (row.profile === opts.profile && !row.valid && row.mutation && row.mutation !== "none" && fixtureMatchesRow(outDir, row)) {
                done.add(row.element);
            }
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
    console.log(`breaking ${elements.length} of ${all.length} fixtures (${done.size} done)`);

    const validator = new Validate({ profile: opts.profile, suffix: ".docx" });
    let broken = 0;
    let noMutation = 0;
    for (const [index, element] of elements.entries()) {
        const fx = await loadFixture(path.join(srcDir, `${element}.docx`));
        const target = findTarget(fx.files, element);
        let row: ElementResult | null = null;
        if (target) {
            const qname = `{${W}}${element}`;
            for (const { label, apply } of mutations()) {
                const mutated: LoadedFixture = { order: fx.order, files: { ...fx.files } };
                const doc = parseXml(mutated.files[target.file]);
                const el = doc.getElementsByTagNameNS(W, element).item(doc.getElementsByTagNameNS(W, element).length - 1)!;
                apply(el, doc);
                mutated.files[target.file] = XML_DECL + serializeXml(doc.documentElement);
                const bytes = await packFixture(mutated);
                const result = await validator.run(bytes);
                const errors = result.issues.filter((i) => i.severity === "error");
                if (!result.valid && errors.some((e) => e.message.includes(qname))) {
                    const file = `${element}.docx`;
                    writeFileSync(path.join(outDir, file), bytes);
                    row = {
                        element,
                        file,
                        sha256: createHash("sha256").update(bytes).digest("hex"),
                        profile: opts.profile,
                        valid: result.valid,
                        mutation: label,
                        issueCount: result.issues.length,
                        errorCodes: [...new Set(errors.map((e) => e.code ?? "unknown"))].sort(),
                        errorMessages: errors.map((e) => e.message.slice(0, 300)),
                    };
                    broken++;
                    break;
                }
            }
        }
        if (!row) {
            noMutation++;
            row = {
                element,
                file: `${element}.docx`,
                sha256: "",
                profile: opts.profile,
                valid: true,
                mutation: "none",
                issueCount: 0,
                errorCodes: [],
                errorMessages: ["no accepted mutation"],
            };
        }
        rowsByElement.set(element, row);
        persistManifestRows(resultsPath, rowsByElement);
        if ((index + 1) % 25 === 0)
            console.log(`  ${index + 1}/${elements.length} (${broken} broken, ${noMutation} without accepted mutation so far)`);
    }
    console.log(
        `done: ${elements.length} processed, ${broken} broken with own-rule failure, ${noMutation} unresolved; results: ${resultsPath}`,
    );
    return noMutation > 0 ? 1 : 0;
});

const readdirSorted = (dir: string): string[] =>
    readdirSync(dir)
        .filter((f) => f.endsWith(".docx"))
        .map((f) => f.replace(/\.docx$/, ""))
        .sort((a, b) => a.localeCompare(b));
