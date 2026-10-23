/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Mode policy (docs/plans/2026-10-23-fixing-validator.md, tasks 10 and 11).
 *
 * Contract from plan § 1: `strict` succeeds only when the complete Word check
 * set passed AND schema validation completed. Every result therefore reports
 * execution metadata — profile, executed/omitted checks, schema completion —
 * so callers can tell a real success from a partial one.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

import JSZip from "jszip";

import { withTempDir } from "../src/lib/run-cli";
import * as xsdEngine from "../src/lib/xsd-validator";
import type { XsdValidator } from "../src/lib/xsd-validator";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken");
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");

const errorCodes = (issues: readonly { severity: string; code?: string }[]): string[] => [
    ...new Set(issues.filter((issue) => issue.severity === "error").map((issue) => issue.code ?? "")),
];

describe("strict policy: success implies complete Word checks and completed schema validation", { timeout: 120_000 }, () => {
    it("reports a full strict run as complete, with the schema walk finished", async () => {
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { profile: "strict" });
        });

        expect(result.valid).toBe(true);
        expect(result.metadata?.profile).toBe("strict");
        expect(result.metadata?.omittedChecks).toEqual([]);
        expect(result.metadata?.executedChecks?.length).toBeGreaterThan(0);
        expect(result.metadata?.schemaCompleted).toBe(true);
    });

    it("never reports a strict success when a Word check fails", async () => {
        const result = await validate(path.join(BROKEN, "package.word-row-without-cell.docx"), { profile: "strict" });

        expect(errorCodes(result.issues)).toContain("word-row-without-cell");
        expect(result.valid).toBe(false);
        expect(result.metadata?.profile).toBe("strict");
        expect(result.metadata?.executedChecks).toContain("word-row-without-cell");
        expect(result.metadata?.omittedChecks).toEqual([]);
    });

    it("never reports a strict success when the schema engine cannot load a schema", async () => {
        const failing: XsdValidator = {
            validate: async (_xml: string, schemaPath: string) => [
                { severity: "error", code: "xsd-schema-load-failed", message: `Schema load failed (${schemaPath}): injected` },
            ],
        };
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { profile: "strict", xsdValidator: failing });
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.schemaCompleted).toBe(false);
    });
});

describe("default profile is resolved from the package format", { timeout: 120_000 }, () => {
    it("validates a DOCX as word-valid, so schema-only hygiene is not a Word-opening error", async () => {
        // k06_1 carries 60 schema-only `w:shd w:fill=""` findings. Microsoft
        // Word opens it, so the default profile must call it valid while
        // keeping the findings visible as warnings.
        const result = await validate(path.join(LANE, "k06_1.redline.docx"));

        expect(result.valid).toBe(true);
        expect(result.metadata?.profile).toBe("word-valid");
        const shading = result.issues.filter((issue) => issue.code === "word-shading-color");
        expect(shading.length).toBeGreaterThan(0);
        expect(shading.every((issue) => issue.severity === "warning")).toBe(true);
    });

    it("keeps the historical lenient default for a PPTX", async () => {
        const [defaultRun, explicitLenient] = await withTempDir(async (dir) => {
            const file = path.join(dir, "deck.pptx");
            writeFileSync(file, await buildPptx());
            return Promise.all([validate(file), validate(file, { profile: "lenient" })]);
        });

        expect(defaultRun.metadata?.profile).toBe("lenient");
        expect(defaultRun.valid).toBe(explicitLenient.valid);
        expect(defaultRun.issues).toEqual(explicitLenient.issues);
    });
});

const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** The smallest OPC package that dispatches to `PPTXSchemaValidator`. */
async function buildPptx(): Promise<Buffer> {
    const zip = new JSZip();
    zip.file(
        "[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
            "</Types>",
    );
    zip.file(
        "_rels/.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${PR}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    );
    zip.file(
        "ppt/presentation.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation xmlns:p="${P}" xmlns:r="${PR}">` +
            `<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    );
    return zip.generateAsync({ type: "nodebuffer" });
}

describe("quick mode runs the common-defect preset and discloses its omissions", { timeout: 120_000 }, () => {
    afterEach(() => {
        vi.restoreAllMocks();
    });

    const quickDocx = async (dir: string, name = "control.docx"): Promise<string> => {
        const file = path.join(dir, name);
        writeFileSync(file, await buildDocx());
        return file;
    };

    it("covers the quick preset groups and names every check it left out", async () => {
        const result = await withTempDir(async (dir) => validate(await quickDocx(dir), { quick: true }));

        expect(result.valid).toBe(true);
        expect(result.metadata?.profile).toBe("word-valid");
        expect(result.metadata?.schemaCompleted).toBe(false);
        expect(result.metadata?.executedChecks).toEqual([
            "zip-bytes-after-eocd",
            "zip-eocd-comment-overrun",
            "xml-syntax",
            "word-row-without-cell",
            "word-grid-column-attributes",
            "word-cell-margin-order",
            "word-table-properties-required",
            "word-textbox-content-namespace",
        ]);
        expect(result.metadata?.omittedChecks).toEqual([
            "word-math-run-insertion",
            "word-math-run-deletion",
            "word-shading-color",
            "word-page-margin-attributes",
        ]);
    });

    it("still rejects a defect the preset covers", async () => {
        const result = await validate(path.join(BROKEN, "package.word-row-without-cell.docx"), { quick: true });

        expect(result.valid).toBe(false);
        expect(errorCodes(result.issues)).toContain("word-row-without-cell");
    });

    it("compiles no XSD schema", async () => {
        const spy = vi.spyOn(xsdEngine, "createXsdValidator");

        await withTempDir(async (dir) => {
            const file = await quickDocx(dir);
            await validate(file, { quick: true });
            await validate(file, { checks: ["tables"] });
        });

        expect(spy).not.toHaveBeenCalled();
    });
});
