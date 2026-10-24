/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for the grid-column attribute check
 * (`src/scripts/office/checks/grid-column-attributes.ts`, Task 8 family 8d,
 * lane issues #209, #219, #224, #226, #233, #240).
 *
 * `CT_TblGridCol` declares exactly one attribute — `w:w`. The jubarte 0.13.0
 * lane writes `w:type="dxa"` on every grid column (including the seven inside
 * `w:tblGridChange/w:tblGrid`), which the schema rejects while Word opens the
 * file regardless. Every defect below is paired with the nearby legal control
 * (`w:w` alone).
 */

import { describe, expect, it } from "vitest";

import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext, type PartContent } from "../src/lib/validation-context";
import { wordGridColumnAttributesCheck } from "../src/scripts/office/checks/grid-column-attributes";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const documentXml = (body: string, ns: string = W): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + `<w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`;

const findings = async (parts: Readonly<Record<string, PartContent>>): Promise<readonly ValidationIssue[]> =>
    wordGridColumnAttributesCheck.run(ValidationContext.fromEntries(parts));

describe("grid-column-attributes check", () => {
    it("reports every attribute CT_TblGridCol does not declare, one finding per attribute", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:tbl><w:tblGrid>` +
                    `<w:gridCol w:w="1995" w:type="dxa"/>` +
                    `<w:gridCol w:w="399" w:type="pct" w:unit="twips"/>` +
                    `</w:tblGrid></w:tbl>`,
            ),
        });
        expect(issues).toHaveLength(3);
        for (const issue of issues) {
            expect(issue.severity).toBe("warning");
            expect(issue.code).toBe("word-grid-column-attributes");
            expect(issue.path).toBe("word/document.xml");
        }
        // The message names the undeclared attribute and keeps the width visible.
        expect(issues[0]!.message).toContain("w:type");
        expect(issues[0]!.message).toContain(`"dxa"`);
        expect(issues[0]!.message).toContain(`w:w="1995"`);
        expect(issues[1]!.message).toContain(`w:w="399"`);
        expect(issues[2]!.message).toContain("w:unit");
    });

    it("names the absent width when a grid column carries only the undeclared attribute", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(`<w:tbl><w:tblGrid><w:gridCol w:type="dxa"/></w:tblGrid></w:tbl>`),
        });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.message).toContain("w:type");
        expect(issues[0]!.message).toContain("w:w");
    });

    it("checks the grid of a w:tblGridChange as well as the live table grid", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
                    `<w:tblPrEx><w:tblPrChange w:id="1"><w:tblPr/></w:tblPrChange></w:tblPrEx>` +
                    `<w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="1995" w:type="dxa"/></w:tblGrid></w:tblGridChange>` +
                    `<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.message).toContain(`w:w="1995"`);
    });

    it("keeps every nearby legal grid column unreported", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(
                `<w:tbl><w:tblGrid>` +
                    `<w:gridCol w:w="100"/>` + // w:w alone: the control
                    `<w:gridCol/>` + // every attribute optional, all omitted
                    `<w:gridCol w:w="1995" w:w14:durableId="ABCDEF01" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>` + // foreign namespace: MCE/extension territory
                    `</w:tblGrid></w:tbl>`,
            ),
        });
        expect(issues).toEqual([]);
    });

    it("reports an unprefixed attribute, which no CT_TblGridCol declaration covers either", async () => {
        const issues = await findings({
            "word/document.xml": documentXml(`<w:tbl><w:tblGrid><w:gridCol w:w="100" type="dxa"/></w:tblGrid></w:tbl>`),
        });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.message).toContain(`type="dxa"`);
    });

    it("scans every story part and matches the strict namespace whatever prefix carries it", async () => {
        const strictDoc =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<word:document xmlns:word="${STRICT}"><word:body>` +
            `<word:tbl><word:tblGrid><word:gridCol word:w="100" word:type="dxa"/></word:tblGrid></word:tbl>` +
            `</word:body></word:document>`;
        const footer = documentXml(`<w:tbl><w:tblGrid><w:gridCol w:w="100" w:type="dxa"/></w:tblGrid></w:tbl>`);
        const issues = await wordGridColumnAttributesCheck.run(
            ValidationContext.fromEntries({ "word/document.xml": strictDoc, "word/footer1.xml": footer }),
        );
        expect(issues.map((issue) => [issue.path, issue.message])).toEqual([
            ["word/document.xml", expect.stringContaining("word:type")],
            ["word/footer1.xml", expect.stringContaining("w:type")],
        ]);
    });

    it("skips a part that fails to parse without throwing, and still checks the others", async () => {
        const issues = await wordGridColumnAttributesCheck.run(
            ValidationContext.fromEntries({
                "word/document.xml": `<w:document xmlns:w="${W}"><w:body>`, // unterminated
                "word/header1.xml": documentXml(`<w:tbl><w:tblGrid><w:gridCol w:w="100" w:type="dxa"/></w:tblGrid></w:tbl>`),
            }),
        );
        expect(issues).toHaveLength(1);
        expect(issues[0]!.path).toBe("word/header1.xml");
    });
});
