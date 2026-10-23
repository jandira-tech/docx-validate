/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Package-level content-type resolution (docs/plans/2026-10-23-fixing-validator.md,
 * Task 4; issues #208, #237, #249).
 *
 * Builds tiny DOCX packages in memory from MINIMAL_PARTS, adds media parts
 * plus [Content_Types].xml declarations, and drives the public `validate()`
 * entry point under every profile. Resolver-level behaviour is unit-tested in
 * `tests/content-types.test.ts`.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PROFILES: readonly Profile[] = ["lenient", "strict", "word-valid"];

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELS_DEFAULT = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`;
const XML_DEFAULT = `<Default Extension="xml" ContentType="application/xml"/>`;
const DOC_OVERRIDE =
    `<Override PartName="/word/document.xml"` +
    ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`;
const STYLES_OVERRIDE =
    `<Override PartName="/word/styles.xml"` + ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;

/** [Content_Types].xml for the minimal package plus extra declarations. */
const contentTypes = (extra: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="${CT_NS}">${RELS_DEFAULT}${XML_DEFAULT}${DOC_OVERRIDE}${STYLES_OVERRIDE}${extra}</Types>`;

const JPEG_BYTES = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]);

type Parts = Record<string, string | Uint8Array>;

const withMedia = (extraDeclarations: string, media: Parts): Parts => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": contentTypes(extraDeclarations),
    ...media,
});

const ctIssues = (issues: readonly { code?: string }[]): { code?: string }[] =>
    issues.filter((issue) => (issue.code ?? "").startsWith("ct-"));

async function validateBytes(bytes: Uint8Array, profile: Profile) {
    return withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
}

describe("content-type resolution across profiles", { timeout: 120_000 }, () => {
    for (const profile of PROFILES) {
        it(`${profile}: media covered only by an exact Override yields neither ct-undeclared-ext nor ct-uncovered-part`, async () => {
            // No <Default Extension="jpeg"> — the ONLY declaration is the
            // exact Override, which the old extension-only media check never
            // consulted (issues #208/#237: jubarte redlines declare every
            // media part this way).
            const bytes = await buildDocx(
                withMedia(`<Override PartName="/word/media/image1.jpeg" ContentType="image/jpeg"/>`, {
                    "word/media/image1.jpeg": JPEG_BYTES,
                }),
            );
            const result = await validateBytes(bytes, profile);
            expect(ctIssues(result.issues)).toEqual([]);
        });
    }
});
