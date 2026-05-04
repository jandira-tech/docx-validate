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
 *                     strict profile we emit `xml-bom-leading` and reject.
 *                     Both cases are pinned below.
 *  - zipGeneration.docx → genuinely missing <Default Extension="png" .../>
 *                         in [Content_Types].xml. Real OOXML violation —
 *                         pin both `valid===false` and the `ct-undeclared-ext`
 *                         code as a positive-test for that detector.
 *  - macroEnabledTemplate.docm → skipped: validate() currently only handles
 *                                .docx/.pptx/.xlsx; .docm is structurally
 *                                identical to .docx but the suffix gate
 *                                rejects it before dispatch.
 */
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate.ts";

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
    "zipGeneration.docx": { codes: ["ct-undeclared-ext"] },
};

const KNOWN_FAILURES_STRICT: Record<string, ExpectedFailure> = {
    // Strict additionally surfaces xml-bom-leading (the canonical strict-only check).
    "office365.docx": { codes: ["xml-bom-leading", "xsd-summary"] },
    "zipGeneration.docx": { codes: ["ct-undeclared-ext"] },
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
        const result = await validate(path.join(FIXTURE_DIR, name), { profile: "strict" });
        expect(result.valid).toBe(false);
        for (const code of expected.codes) {
            expect(result.issues.some((i) => i.code === code)).toBe(true);
        }
    });

    it("office365.docx: lenient strips BOM, so no xml-syntax error appears", async () => {
        // Pin the BOM-strip behaviour: under lenient, no xml-syntax issue is
        // raised even though the file has UTF-8 BOMs on every part. The XSD
        // sequence-order issue is a separate concern (see KNOWN_FAILURES_LENIENT).
        const result = await validate(path.join(FIXTURE_DIR, "office365.docx"));
        expect(result.issues.some((i) => i.code === "xml-syntax")).toBe(false);
        expect(result.issues.some((i) => i.code === "xml-bom-leading")).toBe(false);
    });
});

describe("docx-templates corpus — skipped formats", () => {
    it("macroEnabledTemplate.docm: validate() returns unsupported-file-type for unknown suffix (current behaviour)", async () => {
        const result = await validate(path.join(FIXTURE_DIR, "macroEnabledTemplate.docm"));
        expect(result.valid).toBe(false);
        expect(result.suffix).toBe(".docm");
        expect(result.issues.some((i) => i.code === "unsupported-file-type")).toBe(true);
    });
});
