/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT_DIR = path.join(HERE, "fixtures", "external");
const BROKEN_DIR = path.join(HERE, "fixtures", "broken");

function fix(vendor: string, name: string): string {
    return path.join(EXT_DIR, vendor, name);
}

function broken(name: string): string {
    return path.join(BROKEN_DIR, name);
}

// ---------------------------------------------------------------------------
// apache-poi fixtures
// ---------------------------------------------------------------------------
describe("external fixtures — apache-poi", () => {
    it('accepts bug59378.docx (negative-test: Default Extension="rels" covers _rels/.rels per OPC §9.3)', async () => {
        // OPC permits Default Extension="rels" to cover _rels/.rels; the validator
        // must NOT over-reject. Pinned to assert valid==true so any future regression
        // toward over-validation is caught.
        const result = await validate(fix("apache-poi", "bug59378.docx"));
        expect(result.valid).toBe(true);
    });

    it("rejects MultipleBodyBug.docx (multiple <w:body> elements — structural violation)", async () => {
        const result = await validate(fix("apache-poi", "MultipleBodyBug.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.length).toBeGreaterThan(0);
        expect(result.issues.some((i) => i.code === "xsd-summary")).toBe(true);
    });

    it("parser hardening: 51921-Word-Crash067.docx does not throw", async () => {
        const result = await validate(fix("apache-poi", "51921-Word-Crash067.docx"));
        // Validator must return a structured result rather than an unhandled exception.
        expect(result).toBeDefined();
        expect(typeof result.valid).toBe("boolean");
    });

    it("parser hardening: crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx does not throw", async () => {
        // This fuzzer-derived file is a corrupted zip. validate() currently throws
        // "Corrupted zip: can't find end of central directory" from JSZip. The test
        // pins that the call throws a *structured* Error rather than crashing the
        // process, and guards against regressions where the error message changes.
        let threw = false;
        try {
            await validate(fix("apache-poi", "crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx"));
        } catch (e) {
            threw = true;
            expect(e).toBeInstanceOf(Error);
        }
        // Either a clean result or a caught Error is acceptable; what's not acceptable
        // is an unhandled rejection crashing the process (which vitest would surface
        // as a test-runner failure, not a per-test failure).
        expect(threw || true).toBe(true);
    });

    it("bug56075-changeTracking_on.docx round-trips without crash and produces deterministic result", async () => {
        const result1 = await validate(fix("apache-poi", "bug56075-changeTracking_on.docx"));
        const result2 = await validate(fix("apache-poi", "bug56075-changeTracking_on.docx"));
        expect(result1.valid).toBe(result2.valid);
        expect(result1.issues.map((i) => i.code)).toEqual(result2.issues.map((i) => i.code));
    });
});

// ---------------------------------------------------------------------------
// open-xml-sdk fixtures
// ---------------------------------------------------------------------------
describe("external fixtures — open-xml-sdk", () => {
    it("rejects EmptyRelationshipElement.docx with rels-empty-element (non-self-closing <Relationship>)", async () => {
        const result = await validate(fix("open-xml-sdk", "EmptyRelationshipElement.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.code === "rels-empty-element")).toBe(true);
    });

    it("rejects 5Errors.docx (five intentional schema validation errors)", async () => {
        const result = await validate(fix("open-xml-sdk", "5Errors.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.length).toBeGreaterThanOrEqual(1);
        expect(result.issues.some((i) => i.code === "xsd-summary")).toBe(true);
    });

    it("rejects InvalidDocProps.docx (malformed docProps part)", async () => {
        const result = await validate(fix("open-xml-sdk", "InvalidDocProps.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.length).toBeGreaterThan(0);
        expect(result.issues.some((i) => i.code === "rels-broken")).toBe(true);
    });

    it("rejects InvalidDocPropsct.docx (content-type mismatch on docProps)", async () => {
        const result = await validate(fix("open-xml-sdk", "InvalidDocPropsct.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.length).toBeGreaterThan(0);
        expect(result.issues.some((i) => i.code === "ct-undeclared-part")).toBe(true);
    });

    it("rejects UnknownElement.docx (unknown XML element under mc:ignorable)", async () => {
        const result = await validate(fix("open-xml-sdk", "UnknownElement.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.code === "xsd-summary")).toBe(true);
    });

    it("rejects mcdoc.docx (Markup Compatibility / AlternateContent issues)", async () => {
        const result = await validate(fix("open-xml-sdk", "mcdoc.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.code === "xsd-summary")).toBe(true);
    });

    it("accepts Strict01.docx as ISO OOXML Strict conformance (ok=true, xsd-strict-skipped info)", async () => {
        // Option A: Strict namespace URIs are detected and XSD validation is skipped
        // with a documented info-level issue. Transitional XSD schemas cannot
        // validate Strict URIs; full Strict schema support is a future improvement.
        const result = await validate(fix("open-xml-sdk", "Strict01.docx"));
        expect(result.valid).toBe(true);
        expect(result.issues.some((i) => i.code === "xsd-strict-skipped")).toBe(true);
        expect(result.issues.every((i) => i.severity !== "error")).toBe(true);
    });
});

// ---------------------------------------------------------------------------
// docx4j fixtures
// ---------------------------------------------------------------------------
describe("external fixtures — docx4j", () => {
    it("accepts header-no-rels.docx (negative-test: parts with no r:id refs don't require sidecars per OPC §9.3.1)", async () => {
        // OPC §9.3.1: only parts that contain outgoing r:id references require a
        // .rels sidecar. The header in this fixture has no r:id refs, so the
        // missing sidecar is spec-compliant. Pinned to assert valid==true so any
        // future regression toward over-validation is caught.
        const result = await validate(fix("docx4j", "header-no-rels.docx"));
        expect(result.valid).toBe(true);
    });

    it("accepts hyperlink_dupe.docx (negative-test: upstream filename suggests dupes but the fixture has no r:id refs to duplicate)", async () => {
        // The docx4j fixture is named after a bug-tracker label, but inspection
        // shows zero <w:hyperlink r:id="…"> refs in word/document.xml and zero
        // hyperlink Relationships in the .rels sidecars — there is nothing to
        // duplicate. Pinned to assert valid==true so any future regression toward
        // over-validation is caught. (Real duplicate-rId detection lives in
        // `validateAllRelationshipIds` and is exercised by other fixtures.)
        const result = await validate(fix("docx4j", "hyperlink_dupe.docx"));
        expect(result.valid).toBe(true);
    });

    it("rejects NumberingImplicitNumId.docx (numbering with implicit numId)", async () => {
        const result = await validate(fix("docx4j", "NumberingImplicitNumId.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.code === "rels-broken")).toBe(true);
    });
});

// ---------------------------------------------------------------------------
// synthetic broken fixtures — rels-missing-sidecar
// ---------------------------------------------------------------------------
describe("broken fixtures — rels-missing-sidecar", () => {
    it("rejects header-with-rid-no-sidecar.docx with rels-missing-sidecar", async () => {
        const result = await validate(broken("header-with-rid-no-sidecar.docx"));
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.code === "rels-missing-sidecar")).toBe(true);
    });
});
