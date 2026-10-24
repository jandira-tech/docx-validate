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

import { createHash } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

import JSZip from "jszip";

import { withTempDir } from "../src/lib/run-cli";
import type { WordOutcome } from "../src/lib/word-verifier";
import * as xsdEngine from "../src/lib/xsd-validator";
import type { XsdValidator } from "../src/lib/xsd-validator";
import { validate, type ValidateOptions, type WordVerifier } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

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
            // The package group's stored-name diagnosis runs in quick too (#351).
            "opc-equivalent-part-names",
            "opc-part-name-escaped-unreserved",
            "xml-syntax",
            "word-grid-column-attributes",
            "word-cell-margin-order",
            "word-table-properties-required",
            "word-textbox-content-namespace",
            "word-drawing-extent-missing",
            "word-drawing-extent-misplaced",
            "word-picture-nvpicpr-missing",
            "word-picture-sppr-missing",
            "word-table-properties-misplaced",
            "word-markup-compatibility",
            "word-document-root",
            "word-main-document-part",
            "opc-relationships-structure",
            "word-row-without-cell",
            "word-choice-requires-undeclared",
            "word-duplicate-theme-relationship",
            "word-duplicate-font-table-relationship",
            "word-duplicate-web-settings-relationship",
            "word-duplicate-settings-relationship",
            "word-hyperlink-relationship",
            "word-cell-ends-with-table",
            "word-empty-drawing",
            "word-anchor-missing-attribute",
            "word-position-offset-and-align",
            "word-drawing-scalar-whitespace",
            "word-drawing-arc-radius",
            "word-drawing-extent-negative",
            "word-content-type-invalid",
        ]);
        expect(result.metadata?.omittedChecks).toEqual([
            "comment-markers",
            "comment-threading",
            "word-math-run-insertion",
            "word-math-run-deletion",
            "word-shading-color",
            "word-page-margin-attributes",
            "word-custom-property-filetime-empty",
            "word-run-misplaced",
            "word-run-properties-misplaced",
            "word-underline-misplaced",
            "word-paragraph-properties-misplaced",
            "word-section-properties-misplaced",
            "word-unknown-element",
            "word-link-without-namespace",
            "word-math-parse",
            "word-math-spre-body",
            "word-footnote-reference-missing",
            "word-numbering-picture-bullet",
            "word-attached-toolbar-utf8",
            "word-sdt-id-empty",
            "word-text-contains-paragraph",
            "word-textbox-deleted-text",
            "word-move-from-deleted-text",
        ]);
    });

    it("rejects an archive it cannot open instead of passing the selected checks (#332)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "notzip.docx");
            writeFileSync(file, "this is not a zip archive at all");
            // Only the ZIP-envelope checks can read an archive that does not open;
            // every other check has no part to read, so it did not execute.
            const envelope = ["zip-bytes-after-eocd", "zip-eocd-comment-overrun"];
            const rows: [ValidateOptions, readonly string[]][] = [
                [{ quick: true }, envelope],
                [{ checks: ["xml-syntax"] }, []],
                [{}, envelope],
            ];
            const control = await quickDocx(dir);
            for (const [opts, executed] of rows) {
                const result = await validate(file, opts);
                const label = JSON.stringify(opts);
                expect(result.valid, label).toBe(false);
                expect(errorCodes(result.issues), label).toContain("package-open-failed");
                expect(result.metadata?.executedChecks, label).toEqual(executed);
                // Omitted means "not selected", as for a package that opens: a
                // selected check the archive kept from running is not an
                // omission (Jules on #337).
                expect(result.metadata?.omittedChecks, label).toEqual((await validate(control, opts)).metadata?.omittedChecks);
            }
        });
    });

    it("does not list the archive-only checks as executed for an unpacked directory", async () => {
        const result = await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            for (const [name, xml] of Object.entries(MINIMAL_PARTS)) {
                mkdirSync(path.dirname(path.join(unpacked, name)), { recursive: true });
                writeFileSync(path.join(unpacked, name), xml);
            }
            return validate(unpacked, { quick: true, original: await quickDocx(dir) });
        });

        expect(result.valid).toBe(true);
        // An unpacked directory has no ZIP envelope, so those checks never ran.
        expect(result.metadata?.executedChecks).not.toContain("zip-bytes-after-eocd");
        expect(result.metadata?.executedChecks).not.toContain("zip-eocd-comment-overrun");
        expect(result.metadata?.executedChecks).toContain("xml-syntax");
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

/** A `verifyInWord` stand-in: records the bytes it was handed, returns a canned observation. */
const fakeWordVerifier = (outcome: WordOutcome) => {
    const seen: Uint8Array[] = [];
    const verify: WordVerifier = async (bytes: Uint8Array) => {
        seen.push(bytes);
        return {
            requested: true,
            outcome,
            satisfied: outcome === "clean-open",
            fileSha256: "0".repeat(64),
            platform: "darwin",
            observedAt: "2026-10-23T00:00:00.000Z",
            observationMs: 3000,
            details: "injected",
        };
    };
    return { verify, seen };
};

describe("--verify-word asks Microsoft Word about the exact bytes", { timeout: 120_000 }, () => {
    it("satisfies a portable pass when Word opens the same bytes cleanly", async () => {
        const { verify, seen } = fakeWordVerifier("clean-open");
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            const bytes = await buildDocx();
            writeFileSync(file, bytes);
            return validate(file, { verifyWord: true, wordVerifier: verify });
        });

        expect(result.valid).toBe(true);
        expect(result.metadata?.wordVerification).toMatchObject({ requested: true, outcome: "clean-open", satisfied: true });
        expect(seen).toHaveLength(1);
    });

    it("is NOT satisfied when Word verification is unavailable", async () => {
        const { verify } = fakeWordVerifier("unavailable");
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { verifyWord: true, wordVerifier: verify });
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.wordVerification).toMatchObject({ outcome: "unavailable", satisfied: false });
    });

    it("is NOT satisfied when Word reports a repair warning", async () => {
        const { verify } = fakeWordVerifier("repair-warning");
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx());
            return validate(file, { verifyWord: true, wordVerifier: verify });
        });

        expect(result.valid).toBe(false);
        expect(result.metadata?.wordVerification?.outcome).toBe("repair-warning");
    });

    it("reports a verifier that throws as unavailable instead of rejecting (#333)", async () => {
        const bytes = await buildDocx();
        const run = withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, bytes);
            return validate(file, {
                verifyWord: true,
                wordVerifier: async () => {
                    throw new Error("transport launch failed");
                },
            });
        });

        await expect(run).resolves.toMatchObject({
            valid: false,
            metadata: {
                wordVerification: {
                    requested: true,
                    outcome: "unavailable",
                    satisfied: false,
                    fileSha256: createHash("sha256").update(bytes).digest("hex"),
                    details: expect.stringContaining("transport launch failed"),
                },
            },
        });
    });

    it("stamps an unpacked directory's unavailable verification with the current time, not the epoch (#329)", async () => {
        const before = Date.now();
        const result = await withTempDir(async (dir) => {
            // A directory named like a package: unpacked, so it has no bytes to verify.
            const unpacked = path.join(dir, "unpacked.docx");
            for (const [name, xml] of Object.entries(MINIMAL_PARTS)) {
                mkdirSync(path.dirname(path.join(unpacked, name)), { recursive: true });
                writeFileSync(path.join(unpacked, name), xml);
            }
            return validate(unpacked, { verifyWord: true });
        });

        expect(result.metadata?.wordVerification?.outcome).toBe("unavailable");
        expect(Date.parse(result.metadata?.wordVerification?.observedAt ?? "")).toBeGreaterThanOrEqual(before);
    });

    it("hands the verifier the exact file bytes", async () => {
        const { verify, seen } = fakeWordVerifier("clean-open");
        const bytes = await buildDocx();
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, bytes);
            await validate(file, { verifyWord: true, wordVerifier: verify });
        });

        expect(seen).toHaveLength(1);
        expect(Buffer.from(seen[0] as Uint8Array).equals(bytes)).toBe(true);
    });
});

describe("M1.8 profile policy (#408)", { timeout: 120_000 }, () => {
    it.each(["word-valid", "lenient", "strict"] as const)("reports an escaped-only main part under %s", async (profile) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "escaped-main.docx");
            const parts = Object.fromEntries(
                Object.entries(MINIMAL_PARTS).map(([name, text]) => [name === "word/document.xml" ? "word/%64ocument.xml" : name, text]),
            );
            writeFileSync(file, await buildDocx(parts));
            const result = await validate(file, { profile });
            expect(result.issues).toContainEqual(
                expect.objectContaining({
                    code: "opc-part-name-escaped-unreserved",
                    path: "word/%64ocument.xml",
                    severity: profile === "lenient" ? "warning" : "error",
                }),
            );
            expect(result.valid).toBe(profile === "lenient");
            expect(result.metadata?.executedChecks).toContain("opc-part-name-escaped-unreserved");
        });
    });
});
