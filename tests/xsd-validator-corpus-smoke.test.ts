/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * PR A Task A.2 (revised): wasm engine corpus smoke.
 *
 * The original plan called for a parity test asserting wasm and libxmljs2
 * return identical issue counts. PR A Task A.1 discovered the wasm engine is
 * STRICTLY STRONGER than libxmljs2 — it resolves OOXML schema imports cleanly
 * where libxmljs2 silently degraded (CLAUDE.md note 4). So a parity assertion
 * would fail by design.
 *
 * This smoke test is the canary instead: pick a small set of real fixtures
 * across `working/` and `broken/`, validate each one's `word/document.xml`
 * against the bundled wml.xsd via the wasm validator, and assert that:
 *   - no fixture causes the validator to throw
 *   - the validator returns a structured `ValidationIssue[]`
 *
 * Issue counts are NOT asserted because they will legitimately differ from
 * libxmljs2 (and that's the upgrade). PR B's full-test-suite re-run is the
 * authoritative regression check.
 */

import { describe, it, expect, beforeAll } from "vitest";
import path from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { createXsdValidator, type XsdValidator } from "../src/lib/xsd-validator";
import { unpack } from "../src/scripts/office/unpack";

const SCHEMAS_DIR = path.resolve(__dirname, "..", "src", "scripts", "office", "schemas");
const WML_SCHEMA = path.join(SCHEMAS_DIR, "ISO-IEC29500-4_2016", "wml.xsd");

const SMOKE_FIXTURES = [
    "tests/fixtures/working/sample-document.afterword-repaired-word-repaired.docx",
    "tests/fixtures/working/sample-document.really-repaired-word-repaired.docx",
    "tests/fixtures/broken/comments.unmatched-comment-marker.docx",
    "tests/fixtures/broken/sample-document.broken-tables.docx",
];

describe("xsd-validator wasm engine corpus smoke", () => {
    let validator: XsdValidator;
    let scratchDir: string;

    beforeAll(async () => {
        validator = await createXsdValidator();
        scratchDir = mkdtempSync(path.join(tmpdir(), "xsd-corpus-smoke-"));
    });

    it.each(SMOKE_FIXTURES)(
        "wasm validator does not throw on %s",
        async (fixturePath) => {
            const unpackedDir = path.join(scratchDir, path.basename(fixturePath));
            await unpack(fixturePath, unpackedDir);

            const documentXmlPath = path.join(unpackedDir, "word", "document.xml");
            const documentXml = readFileSync(documentXmlPath, "utf-8");

            // The contract: validate returns an array, regardless of what's in
            // it. The wasm engine should never throw on real DOCX content.
            const issues = await validator.validate(documentXml, WML_SCHEMA);
            expect(Array.isArray(issues)).toBe(true);
            for (const issue of issues) {
                expect(issue.severity).toMatch(/^(error|warning|info)$/);
                expect(typeof issue.message).toBe("string");
                expect(issue.message.length).toBeGreaterThan(0);
            }
        },
        30_000,
    );

    it("cleanup", () => {
        rmSync(scratchDir, { recursive: true, force: true });
    });
});
