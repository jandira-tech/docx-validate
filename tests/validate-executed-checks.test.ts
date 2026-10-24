/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `metadata.executedChecks` must disclose what a run ACTUALLY executed
 * (#319).
 *
 * The check registry lists every check for a format, but a profile only
 * reaches some of them: the Word rules run under `word-valid` and `strict`
 * only. Reporting the whole registry as "executed" told a caller nothing
 * about what was checked — a check that passed and a check that never ran
 * looked identical.
 *
 * The specimen is the `external/superdoc/super-editor/hyperlink-exported.docx`
 * shape, rebuilt in memory: a namespace-less `<link/>` inside `w:rPr`. Word
 * refuses to open that file, and the Word rule that says so only runs in the
 * two Word profiles.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const WORD_RULE_IDS: readonly string[] = WORD_RULE_CHECKS.map((check) => check.id);

/** `hyperlink-exported.docx`: `<link/>` with no XML namespace inside `w:rPr`. */
const LINK_WITHOUT_NAMESPACE = MINIMAL_PARTS["word/document.xml"]!.replace(
    "<w:r><w:t>Hello</w:t></w:r>",
    "<w:r><w:rPr><link/></w:rPr><w:t>Hello</w:t></w:r>",
);

const writeSpecimen = async (dir: string, name: string, documentXml: string): Promise<string> => {
    const file = path.join(dir, name);
    writeFileSync(file, await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": documentXml }));
    return file;
};

describe("metadata.executedChecks reports the checks a run actually executed (#319)", { timeout: 120_000 }, () => {
    it("lists no Word rule under lenient, and every Word rule under word-valid and strict", async () => {
        await withTempDir(async (dir) => {
            const file = await writeSpecimen(dir, "link.docx", LINK_WITHOUT_NAMESPACE);

            const lenient = await validate(file, { profile: "lenient" });
            const wordValid = await validate(file, { profile: "word-valid" });
            const strict = await validate(file, { profile: "strict" });

            // Lenient never runs the Word rules, so it must not list them.
            const lenientChecks = lenient.metadata?.executedChecks ?? [];
            for (const id of WORD_RULE_IDS) {
                expect(lenientChecks).not.toContain(id);
            }
            // Lenient runs stored-name diagnosis, XML syntax, content-type uniqueness,
            // the content-types PartName grammar and the shared comment checks; the
            // ZIP-envelope checks remain Word-only.
            expect(lenientChecks).toEqual([
                "opc-equivalent-part-names",
                "opc-part-name-escaped-unreserved",
                "opc-part-name-escaped-slash",
                "opc-part-name-word-refused",
                "xml-syntax",
                "ct-duplicate-declaration",
                "ct-override-partname-invalid",
                "ct-override-partname-word-refused",
                // The content-types Markup Compatibility refusal (#459).
                "ct-markup-compatibility",
                "word-part-content-type-generic",
                // An inexact Word content type (#550).
                "word-part-content-type-inexact",
                "word-part-content-type-mismatch",
                "comment-markers",
                "comment-threading",
            ]);
            // The finding that rule reports exists only in the profiles that run it.
            expect(lenient.issues.some((issue) => issue.code === "word-link-without-namespace")).toBe(false);
            expect(wordValid.issues.some((issue) => issue.code === "word-link-without-namespace")).toBe(true);

            // Both Word profiles run every Word rule, and say so.
            for (const id of WORD_RULE_IDS) {
                expect(wordValid.metadata?.executedChecks ?? []).toContain(id);
                expect(strict.metadata?.executedChecks ?? []).toContain(id);
            }
        });
    });

    it("keeps the XML well-formedness check in the disclosure of every profile", async () => {
        await withTempDir(async (dir) => {
            const file = await writeSpecimen(dir, "control.docx", MINIMAL_PARTS["word/document.xml"]!);

            for (const profile of ["lenient", "word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.metadata?.executedChecks).toContain("xml-syntax");
            }
        });
    });

    it("leaves a quick run's disclosure unchanged", async () => {
        const result = await withTempDir(async (dir) => {
            const file = await writeSpecimen(dir, "control.docx", MINIMAL_PARTS["word/document.xml"]!);
            return validate(file, { quick: true });
        });

        // A quick run reports its own selection; the full-run disclosure
        // change must not add the Word rules it never selected.
        expect(result.metadata?.executedChecks).toContain("word-row-without-cell");
        expect(result.metadata?.executedChecks).toContain("word-main-document-part");
        expect(result.metadata?.executedChecks).not.toContain("word-shading-color");
        expect(result.metadata?.executedChecks).not.toContain("word-link-without-namespace");
    });

    it("lists the envelope checks that ran before the package failed to extract (CodeRabbit on #335)", async () => {
        const originalLoadAsync = JSZip.loadAsync;
        // A zip-slip entry makes extraction refuse the package after it opened.
        JSZip.loadAsync = async function loadAsyncPatched(data, options) {
            const loaded = await originalLoadAsync.call(this, data, options);
            loaded.files["../../escaped.txt"] = {
                name: "../../escaped.txt",
                dir: false,
                async: async () => Buffer.from("malicious"),
            } as unknown as JSZip.JSZipObject;
            return loaded;
        };
        try {
            await withTempDir(async (dir) => {
                const file = await writeSpecimen(dir, "slip.docx", MINIMAL_PARTS["word/document.xml"]!);
                for (const profile of ["word-valid", "strict"] as const) {
                    const result = await validate(file, { profile });
                    expect(
                        result.issues.map((issue) => issue.code),
                        profile,
                    ).toEqual(["package-extract-failed"]);
                    // The ZIP envelope is read from the bytes before extraction, so its checks ran.
                    const executed = result.metadata?.executedChecks ?? [];
                    expect(executed.slice(0, 7), profile).toEqual([
                        "zip-bytes-after-eocd",
                        "zip-bytes-before-archive",
                        "zip-eocd-comment-overrun",
                        "opc-equivalent-part-names",
                        "opc-part-name-escaped-unreserved",
                        "opc-part-name-escaped-slash",
                        "opc-part-name-word-refused",
                    ]);
                    // The archive still reads, so the checks a partial run runs
                    // on it ran too, and their findings are in the failure (#454).
                    expect(executed, profile).toContain("ct-override-partname-invalid");
                }
                // Lenient never runs the envelope checks, and it runs no Word rule.
                const lenient = await validate(file, { profile: "lenient" });
                expect(lenient.metadata?.executedChecks).toEqual([
                    "opc-equivalent-part-names",
                    "opc-part-name-escaped-unreserved",
                    "opc-part-name-escaped-slash",
                    "opc-part-name-word-refused",
                    "xml-syntax",
                    "ct-duplicate-declaration",
                    "ct-override-partname-invalid",
                    "ct-override-partname-word-refused",
                    "ct-markup-compatibility",
                    // A DOCX content-types check, run under every profile (#432).
                    "word-part-content-type-generic",
                    // An inexact Word content type, likewise (#550).
                    "word-part-content-type-inexact",
                    "word-part-content-type-mismatch",
                    "comment-markers",
                    "comment-threading",
                ]);
            });
        } finally {
            JSZip.loadAsync = originalLoadAsync;
        }
    });
});
