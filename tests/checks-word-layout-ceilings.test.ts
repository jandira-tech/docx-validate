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
 * Unit tests for the layout-bomb ceilings
 * (`src/scripts/office/checks/word-layout-ceilings.ts`): an oversized
 * `w:gridSpan` and an over-ceiling field-marker count in a story part are
 * error findings naming the ceiling.
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext, type PartContent } from "../src/lib/validation-context";
import type { ResourceCeilingsInput } from "../src/lib/resource-ceilings";
import {
    WORD_FIELD_COUNT_CEILING,
    WORD_GRID_SPAN_CEILING,
    wordLayoutCeilingsCheck,
} from "../src/scripts/office/checks/word-layout-ceilings";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const documentXml = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

const findings = async (
    parts: Readonly<Record<string, PartContent>>,
    ceilings?: ResourceCeilingsInput,
): Promise<readonly ValidationIssue[]> => wordLayoutCeilingsCheck.run(ValidationContext.fromEntries(parts, { ceilings }));

describe("word-layout-ceilings check", () => {
    it("reports a w:gridSpan over the ceiling (default 63), naming it", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(`<w:tbl><w:tc><w:tcPr><w:gridSpan w:val="100"/></w:tcPr></w:tc></w:tbl>`),
        });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: WORD_GRID_SPAN_CEILING, path: "word/document.xml" });
        expect(issues[0]!.message).toContain("100");
        expect(issues[0]!.message).toContain("maxGridSpan");
    });

    it("accepts a grid span under a raised ceiling", async () => {
        const issues = await findings(
            { "word/document.xml": documentXml(`<w:tbl><w:tc><w:tcPr><w:gridSpan w:val="100"/></w:tcPr></w:tc></w:tbl>`) },
            { maxGridSpan: 200 },
        );
        expect(issues).toEqual([]);
    });

    it("reports a field-marker count over the ceiling", async () => {
        const markers = Array.from({ length: 12 }, () => `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`).join("");
        const issues = await findings({ "word/document.xml": documentXml(`<w:p>${markers}</w:p>`) }, { maxFieldMarkers: 10 });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: WORD_FIELD_COUNT_CEILING, path: "word/document.xml" });
        expect(issues[0]!.message).toContain("12");
    });

    it("accepts a normal table and a normal field count under the defaults", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:tbl><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr></w:tc></w:tbl><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>`,
            ),
        });
        expect(issues).toEqual([]);
    });

    it.each(["1e999", "64.5", "00000000000000000000000000000000000064", "1" + "0".repeat(400)])(
        "reads w:gridSpan %s against the ceiling instead of letting it slip under (r4649505071)",
        async (val) => {
            const issues = await findings({
                "word/document.xml": documentXml(`<w:tbl><w:tc><w:tcPr><w:gridSpan w:val="${val}"/></w:tcPr></w:tc></w:tbl>`),
            });
            expect(issues.map((issue) => issue.code)).toEqual([WORD_GRID_SPAN_CEILING]);
            expect(issues[0]!.message).toContain(`worst is ${val}`);
        },
    );

    it("ignores a w:gridSpan spelling that names no number", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:tbl><w:tc><w:tcPr><w:gridSpan w:val="wide"/><w:gridSpan w:val=""/></w:tcPr></w:tc></w:tbl>`,
            ),
        });
        expect(issues).toEqual([]);
    });
});
