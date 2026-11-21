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
 * Unit tests for the active-content finding
 * (`src/scripts/office/checks/word-macro-project.ts`): a VBA macro project
 * in the package is named as a warning — never executed — under every
 * profile.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext, type PartContent } from "../src/lib/validation-context";
import { WORD_MACRO_PROJECT, wordMacroProjectCheck } from "../src/scripts/office/checks/word-macro-project";

const findings = async (parts: Readonly<Record<string, PartContent>>): Promise<readonly ValidationIssue[]> =>
    wordMacroProjectCheck.run(ValidationContext.fromEntries(parts));

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>`;

describe("word-macro-project check", () => {
    it("names a vbaProject.bin part as a warning", async () => {
        const issues = await findings({ "word/vbaProject.bin": new Uint8Array([0, 1, 2]), "word/document.xml": "<w:document/>" });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "warning", code: WORD_MACRO_PROJECT, path: "word/vbaProject.bin" });
        expect(issues[0]!.message).toContain("never executes");
    });

    it("names a macro-enabled content type in the content-types stream", async () => {
        const issues = await findings({ "[Content_Types].xml": CONTENT_TYPES, "word/document.xml": "<w:document/>" });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "warning", code: WORD_MACRO_PROJECT, path: "[Content_Types].xml" });
    });

    it("reports nothing for a plain package", async () => {
        const issues = await findings({
            "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
            "word/document.xml": "<w:document/>",
        });
        expect(issues).toEqual([]);
    });
});
