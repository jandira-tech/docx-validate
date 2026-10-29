/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * #1091: the full-run verdicts #1065 claimed for a Strict-namespace id. Since
 * #1065, outside lenient `validateAllRelationshipIds` reads the Strict
 * relationships namespace. Its rows called the method, which skips the
 * profile severity policy, so these rows run `validate()` on a uniformly
 * Strict package, `r` bound to the Strict namespace as Word saves one. Each
 * package is the minimal one translated to Strict, clean under every
 * profile, plus one dangling id on an element no Word rule judges
 * (`w:altChunk`, `p:sldId`), so the run's only finding is that id's.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/** `parts` with every Transitional namespace and relationship type in them spelled in the Strict conformance class. */
const toStrict = (parts: Readonly<Record<string, string>>): Record<string, string> =>
    Object.fromEntries(
        Object.entries(parts).map(([name, xml]) => [
            name,
            xml
                .replaceAll(
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "http://purl.oclc.org/ooxml/wordprocessingml/main",
                )
                .replaceAll("http://schemas.openxmlformats.org/presentationml/2006/main", "http://purl.oclc.org/ooxml/presentationml/main")
                .replaceAll("http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main")
                .replaceAll("http://schemas.openxmlformats.org/officeDocument/2006/relationships", STRICT_R),
        ]),
    );

const strictDocx = toStrict(MINIMAL_PARTS);
const strictDocument = strictDocx["word/document.xml"]!.replace("<w:document ", `<w:document xmlns:r="${STRICT_R}" `);
const docx = (body: string) => buildDocx({ ...strictDocx, "word/document.xml": strictDocument.replace("<w:sectPr>", `${body}<w:sectPr>`) });

const strictDeck = toStrict(MINIMAL_PPTX_PARTS);
const pptx = (rid: string) =>
    buildPptx({ ...strictDeck, "ppt/presentation.xml": strictDeck["ppt/presentation.xml"]!.replace('r:id="rId2"', `r:id="${rid}"`) });

const SPECIMENS = {
    "docx calibration": { suffix: ".docx", build: () => docx("") },
    "docx w:altChunk naming no relationship": { suffix: ".docx", build: () => docx(`<w:altChunk r:id="rIdX"/>`) },
    "pptx calibration": { suffix: ".pptx", build: () => pptx("rId2") },
    "pptx p:sldId naming no relationship": { suffix: ".pptx", build: () => pptx("rId9") },
} as const;

/** The run's verdict and its non-info findings as `severity code path`. */
const run = async (specimen: keyof typeof SPECIMENS, profile: "lenient" | "word-valid" | "strict") =>
    withTempDir(async (dir) => {
        const file = path.join(dir, `specimen${SPECIMENS[specimen].suffix}`);
        writeFileSync(file, await SPECIMENS[specimen].build());
        const result = await validate(file, { profile });
        return {
            valid: result.valid,
            findings: result.issues
                .filter((issue) => issue.severity !== "info")
                .map((issue) => `${issue.severity} ${issue.code} ${issue.path}`),
        };
    });

describe("a Strict package's dangling id in a full run (#1091)", () => {
    it.each([
        ["docx calibration", "lenient"],
        ["docx calibration", "word-valid"],
        ["docx calibration", "strict"],
        ["pptx calibration", "lenient"],
        ["pptx calibration", "word-valid"],
        ["pptx calibration", "strict"],
    ] as const)("the %s is clean under %s", async (specimen, profile) => {
        expect(await run(specimen, profile)).toEqual({ valid: true, findings: [] });
    });

    it.each([
        ["docx w:altChunk naming no relationship", "lenient", true, []],
        ["docx w:altChunk naming no relationship", "word-valid", true, ["warning rels-id-missing word/document.xml"]],
        ["docx w:altChunk naming no relationship", "strict", false, ["error rels-id-missing word/document.xml"]],
        ["pptx p:sldId naming no relationship", "lenient", true, []],
        ["pptx p:sldId naming no relationship", "word-valid", false, ["error rels-id-missing ppt/presentation.xml"]],
        ["pptx p:sldId naming no relationship", "strict", false, ["error rels-id-missing ppt/presentation.xml"]],
    ] as const)("a %s under %s: valid %s, findings %j", async (specimen, profile, valid, findings) => {
        expect(await run(specimen, profile)).toEqual({ valid, findings });
    });
});
