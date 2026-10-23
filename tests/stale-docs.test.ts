/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The strings and comments around the validator must describe the current
 * Word-first design (#321).
 *
 * Before the Word-first stack, every run validated the package against the
 * XSD schemas and the Word rules did not exist. Several user-facing strings
 * and comments still described that design, so the CLI told users something
 * untrue about what a default run does.
 */

import { promises as fs, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { buildValidateCommand, validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SRC = path.join(HERE, "..", "src");

const TRANSITIONAL = "http://schemas.openxmlformats.org/";
const STRICT = "http://purl.oclc.org/ooxml/";

/** Rewrite the OOXML namespace URIs of the minimal package to Strict URIs. */
const strictify = (xml: string): string =>
    xml
        .replaceAll(`${TRANSITIONAL}wordprocessingml/2006/main`, `${STRICT}wordprocessingml/main`)
        .replaceAll(`${TRANSITIONAL}officeDocument/2006/relationships`, `${STRICT}officeDocument/relationships`);

const STRICT_PARTS: Readonly<Record<string, string>> = Object.freeze({
    ...MINIMAL_PARTS,
    "_rels/.rels": strictify(MINIMAL_PARTS["_rels/.rels"]!),
    "word/_rels/document.xml.rels": strictify(MINIMAL_PARTS["word/_rels/document.xml.rels"]!),
    "word/document.xml": strictify(MINIMAL_PARTS["word/document.xml"]!),
    "word/styles.xml": strictify(MINIMAL_PARTS["word/styles.xml"]!),
});

/** The JSDoc block that documents one field, e.g. `profile?: Profile;`. */
const docCommentOf = (source: string, field: string): string => {
    const match = new RegExp(`(/\\*\\*(?:[^*]|\\*(?!/))*\\*/)\\s*${field}`).exec(source);
    return match?.[1] ?? "";
};

describe("the validate command description (#321)", () => {
    it("does not promise XSD validation for a default run", () => {
        const description = buildValidateCommand().description();

        // The default DOCX run is word-valid: it compiles no schema.
        expect(description).not.toMatch(/XSD/i);
        expect(description).toMatch(/Word/i);
    });
});

describe("the xsd-strict-skipped message (#321)", () => {
    it("says the ISO Strict schemas are bundled and used by the strict profile", async () => {
        const result = await withTempDir(async (dir) => {
            const file = path.join(dir, "strict.docx");
            await fs.writeFile(file, await buildDocx(STRICT_PARTS));
            return validate(file, { profile: "lenient" });
        });

        const skipped = result.issues.find((issue) => issue.code === "xsd-strict-skipped");
        expect(skipped, "the lenient profile skips Strict parts with xsd-strict-skipped").toBeDefined();
        expect(skipped?.message).not.toMatch(/only Transitional schemas are bundled/i);
        expect(skipped?.message).toMatch(/strict profile/i);
    });
});

describe("the ValidateOptions.profile doc comment (#321)", () => {
    it("names the per-format default instead of claiming lenient", () => {
        const source = readFileSync(path.join(SRC, "scripts", "office", "validate.ts"), "utf-8");
        const doc = docCommentOf(source, "profile\\?: Profile;");

        expect(doc).toMatch(/word-valid/);
        expect(doc).toMatch(/\.docx|DOCX/i);
        expect(doc).not.toMatch(/Defaults to `"lenient"`/);
    });
});

describe("the base.ts schema comments (#321)", () => {
    const source = readFileSync(path.join(SRC, "scripts", "office", "validators", "base.ts"), "utf-8");

    it("no longer claims the removed remote Dublin Core import is filtered", () => {
        expect(source).not.toMatch(/remote\s+Dublin\s+Core\s+import/i);
    });

    it("no longer claims word-valid consults the ignored-error list", () => {
        // Word-valid runs no XSD, so it never consults IGNORED_VALIDATION_ERRORS.
        expect(source).not.toMatch(/Lenient and word-valid keep it for compatibility/i);
    });
});

describe("the DOCX validator header (#321)", () => {
    it("does not describe the validator as an XSD-only one", () => {
        const source = readFileSync(path.join(SRC, "scripts", "office", "validators", "docx.ts"), "utf-8");
        const header = /\/\*\*[\s\S]*?\*\//.exec(source)?.[0] ?? "";

        expect(header).not.toMatch(/against XSD schemas/i);
        expect(header).toMatch(/word-valid/);
    });
});

describe("the strict-schema spec header (#321)", () => {
    it("says word-valid compiles no schema", () => {
        const source = readFileSync(path.join(HERE, "strict-schema-validation.test.ts"), "utf-8");
        const header = /\/\*\*[\s\S]*?\*\//.exec(source)?.[0] ?? "";

        expect(header).toMatch(/word-valid/);
        expect(header).not.toMatch(/Lenient\/word-valid keep the documented/);
    });
});
