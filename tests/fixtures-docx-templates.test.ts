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

/**
 * Corpus test against the docx-templates fixture set
 * (`tests/fixtures/external/docx-templates/`, copied from the upstream
 * `docx-templates` project's `__tests__/fixtures/`).
 *
 * Those files are real-world Word/Office 365 templates docx-templates uses
 * for its own end-to-end suite. From an OOXML perspective they should all be
 * structurally valid; we use them as a wide negative-test surface to catch
 * validator regressions where we'd otherwise over-reject perfectly fine
 * documents.
 *
 * Empirically catalogued exceptions (pinned so behavioural drift surfaces):
 *
 *  - office365.docx → every part begins with a UTF-8 BOM (Microsoft Office
 *                     routinely emits these and the XML spec permits them).
 *                     Under the default lenient profile we silently strip
 *                     the BOM and the document validates as ok. Under the
 *                     strict profile the BOM is still reported, but as a
 *                     `warning` — a BOM never fails a document on its own.
 *                     Both cases are pinned below.
 *  - zipGeneration.docx → declares its only media part by an exact
 *                         <Override PartName="/word/media/image1.png"/> and no
 *                         png <Default>. OPC resolves Override first (ECMA-376
 *                         Part 2 §10.1.2.4), so the part IS covered and Word
 *                         opens the file cleanly; `ct-undeclared-ext` used to
 *                         flag it (#208). Lenient now passes it; strict still
 *                         rejects it for its paraId and style-default gaps.
 *  - macroEnabledTemplate.docm → skipped: validate() currently only handles
 *                                .docx/.pptx/.xlsx; .docm is structurally
 *                                identical to .docx but the suffix gate
 *                                rejects it before dispatch.
 */
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = path.resolve(HERE, "fixtures", "external", "docx-templates");

interface ExpectedFailure {
    /** issue codes that MUST appear in `result.issues` */
    codes: string[];
}

const KNOWN_FAILURES_LENIENT: Record<string, ExpectedFailure> = {
    // BOMs no longer trip us (parseXml strips them under both profiles), but
    // office365.docx has a separate real-world Office 365 quirk: word/settings.xml
    // has <w:zoom> in a position the ISO Transitional XSD does not allow in its
    // CT_Settings sequence. Word produces and reads this fine; the bundled XSDs
    // are stricter than Microsoft's actual implementation. A future "lenient
    // tolerates Office sequence-order deviations" pass would unblock this.
    "office365.docx": { codes: ["xsd-summary"] },
};

const KNOWN_FAILURES_STRICT: Record<string, ExpectedFailure> = {
    // Strict surfaces xml-bom-leading as a WARNING (a BOM is XML-legal), so the
    // strict rejection of office365.docx comes from the XSD sequence-order
    // deviation alone — see KNOWN_FAILURES_LENIENT.
    "office365.docx": { codes: ["xml-bom-leading", "xsd-summary"] },
    "zipGeneration.docx": { codes: ["paraid-missing-element", "style-default-missing"] },
};

const SKIP = new Set<string>([
    // .docm — validate() suffix gate rejects it; structurally identical to .docx.
    "macroEnabledTemplate.docm",
]);

const allFixtures = (await fs.readdir(FIXTURE_DIR))
    // Office writes lockfiles as `~$*.docx` while a document is open. They are
    // not real DOCX files; defensively filter even though the copy step
    // already excludes them.
    .filter((f) => !f.startsWith("~$"))
    .filter((f) => f.endsWith(".docx") || f.endsWith(".docm"))
    .sort();

const validFixturesLenient = allFixtures.filter((f) => !KNOWN_FAILURES_LENIENT[f] && !SKIP.has(f));

describe("docx-templates corpus — valid templates pass under default (lenient) profile", () => {
    it.each(validFixturesLenient)("%s validates as ok", async (name) => {
        const result = await validate(path.join(FIXTURE_DIR, name));
        if (!result.valid) {
            console.error(`unexpected failure for ${name}:`, result.issues);
        }
        expect(result.valid).toBe(true);
        expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    });
});

describe("docx-templates corpus — media declared by Override (#208)", () => {
    it.each(["lenient", "strict"] as const)("zipGeneration.docx carries no ct-undeclared-ext under %s", async (profile) => {
        const result = await validate(path.join(FIXTURE_DIR, "zipGeneration.docx"), { profile });
        expect(result.issues.map((i) => i.code)).not.toContain("ct-undeclared-ext");
    });
});

describe("docx-templates corpus — known failures under lenient profile", () => {
    it.each(Object.keys(KNOWN_FAILURES_LENIENT).sort())("%s is rejected with the expected codes", async (name) => {
        const expected = KNOWN_FAILURES_LENIENT[name];
        if (!expected) throw new Error(`missing expectation for ${name}`);
        const result = await validate(path.join(FIXTURE_DIR, name));
        expect(result.valid).toBe(false);
        for (const code of expected.codes) {
            expect(result.issues.some((i) => i.code === code)).toBe(true);
        }
    });
});

describe("docx-templates corpus — strict profile flags BOM and other quirks", () => {
    it.each(Object.keys(KNOWN_FAILURES_STRICT).sort())("%s under strict profile is rejected with the expected codes", async (name) => {
        const expected = KNOWN_FAILURES_STRICT[name];
        if (!expected) throw new Error(`missing expectation for ${name}`);
        const result = await validate(path.join(FIXTURE_DIR, name), {
            profile: "strict",
        });
        expect(result.valid).toBe(false);
        for (const code of expected.codes) {
            expect(result.issues.some((i) => i.code === code)).toBe(true);
        }
    });

    it("office365.docx: lenient strips BOM and surfaces it as a WARNING (not an error)", async () => {
        // Pin the BOM behaviour under lenient: parseXml still strips the BOM
        // so no xml-syntax error appears, but `validateNoBom` now records
        // each BOM as severity:warning rather than going silent. That keeps
        // the gap visible in `result.issues` for callers who audit them,
        // without flipping `valid` to false. The XSD sequence-order issue
        // is a separate concern (see KNOWN_FAILURES_LENIENT).
        const result = await validate(path.join(FIXTURE_DIR, "office365.docx"));
        expect(result.issues.some((i) => i.code === "xml-syntax")).toBe(false);
        const bomIssues = result.issues.filter((i) => i.code === "xml-bom-leading");
        expect(bomIssues.length).toBeGreaterThan(0);
        expect(bomIssues.every((i) => i.severity === "warning")).toBe(true);
    });

    it("office365.docx: strict also reports BOMs as warnings, never as errors", async () => {
        // Same contract under strict: XML 1.0 §4.3.3 permits a leading BOM and
        // Microsoft Office emits one in every part of this fixture, so it must
        // not reject the document. It is still visible for callers that audit
        // canonical bytes.
        const result = await validate(path.join(FIXTURE_DIR, "office365.docx"), { profile: "strict" });
        const bomIssues = result.issues.filter((i) => i.code === "xml-bom-leading");
        expect(bomIssues.length).toBeGreaterThan(0);
        expect(bomIssues.every((i) => i.severity === "warning")).toBe(true);
        // The strict rejection is the Office XSD sequence-order deviation,
        // never the BOM.
        expect(result.issues.filter((i) => i.severity === "error").map((i) => i.code)).not.toContain("xml-bom-leading");
    });
});

describe("docx-templates corpus — skipped formats", () => {
    it("macroEnabledTemplate.docm: validate() treats .docm as a supported format (validated like .docx)", async () => {
        // .docm has been in SUPPORTED_SUFFIXES since the first commit, so it is
        // dispatched to the DOCX validators rather than skipped. This fixture is
        // well-formed, so it validates cleanly with no unsupported-file-type issue.
        const result = await validate(path.join(FIXTURE_DIR, "macroEnabledTemplate.docm"));
        expect(result.suffix).toBe(".docm");
        expect(result.issues.some((i) => i.code === "unsupported-file-type")).toBe(false);
        expect(result.valid).toBe(true);
    });
});
