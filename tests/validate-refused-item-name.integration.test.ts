/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile, ValidationResult } from "../src/lib/types";
import { validate, type ValidateRunResult } from "../src/scripts/office/validate";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import { MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";
import { buildRawZip, type RawZipEntry } from "./support/raw-zip";

/*
 * An item name that places no part (#372, #378) is refused for staging, as
 * before, but it no longer hides the rest of the package (#629): the refusal
 * is reported per name, as `opc-item-name-refused`, and the checks that read
 * the archive directly still judge every other entry, as a failed extraction
 * already does (#454).
 */
const ID = "opc-item-name-refused";
const HOSTILE = "word/a/../b.xml";
const MALFORMED_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.slice(0, -20);
const PROFILES: readonly Profile[] = ["word-valid", "strict", "lenient"];

const entries = (parts: Readonly<Record<string, string>>): RawZipEntry[] =>
    Object.entries(parts).map(([name, content]) => ({ name, content }));
const malformedDocx = (...extra: RawZipEntry[]): Buffer =>
    buildRawZip([...entries({ ...MINIMAL_PARTS, "word/document.xml": MALFORMED_DOCUMENT }), ...extra]);
const errors = (result: ValidationResult): string[] =>
    result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);

const run = async (bytes: Buffer, options: Parameters<typeof validate>[1] = {}, suffix = ".docx"): Promise<ValidateRunResult> =>
    withTempDir(async (dir) => {
        const file = path.join(dir, `refused${suffix}`);
        writeFileSync(file, bytes);
        return validate(file, options);
    });

describe("a refused item name no longer hides the rest of the package (#629)", { timeout: 120_000 }, () => {
    it.each(PROFILES)("%s: reports the refused name as its own error and still judges the other entries", async (profile) => {
        const result = await run(malformedDocx({ name: HOSTILE, content: "<x/>" }), { profile });
        expect(result.valid).toBe(false);
        expect(errors(result)).toEqual(
            expect.arrayContaining([`${ID}@${HOSTILE}`, "package-extract-failed@", "xml-syntax@word/document.xml"]),
        );
        expect(result.issues.find((issue) => issue.code === "package-extract-failed")?.message).toContain(HOSTILE);
        expect(result.metadata?.executedChecks).toEqual(expect.arrayContaining([ID, "xml-syntax"]));
        expect(result.xsdWalk.started).toBe(false);
    });

    it("reports every refused name, while the staging failure names the first", async () => {
        const result = await run(
            malformedDocx(
                { name: HOSTILE, content: "<x/>" },
                { name: "/word/abs.xml", content: "<x/>" },
                { name: "word//empty.xml", content: "<x/>" },
            ),
        );
        expect(errors(result).filter((entry) => entry.startsWith(`${ID}@`))).toEqual(
            [HOSTILE, "/word/abs.xml", "word//empty.xml"].sort().map((name) => `${ID}@${name}`),
        );
        const failures = result.issues.filter((issue) => issue.code === "package-extract-failed");
        expect(failures).toHaveLength(1);
        expect(failures[0]!.message).toContain(HOSTILE);
        expect(errors(result)).toContain("xml-syntax@word/document.xml");
    });

    it("judges the entry a later hostile name hid, not the hostile entry's bytes", async () => {
        // JSZip resolves `word/x/../document.xml` onto `word/document.xml`, and the later entry wins its view.
        const hidesMalformed = await run(malformedDocx({ name: "word/x/../document.xml", content: MINIMAL_PARTS["word/document.xml"]! }));
        expect(errors(hidesMalformed)).toEqual(expect.arrayContaining([`${ID}@word/x/../document.xml`, "xml-syntax@word/document.xml"]));

        const hidesWellFormed = await run(
            buildRawZip([...entries(MINIMAL_PARTS), { name: "word/x/../document.xml", content: MALFORMED_DOCUMENT }]),
        );
        expect(errors(hidesWellFormed)).toContain(`${ID}@word/x/../document.xml`);
        expect(errors(hidesWellFormed)).not.toContain("xml-syntax@word/document.xml");
    });

    it("a partial run reports no error the full run does not (#381)", async () => {
        const bytes = malformedDocx({ name: HOSTILE, content: "<x/>" });
        for (const profile of ["word-valid", "lenient"] as const) {
            const full = errors(await run(bytes, { profile }));
            for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: [ID] }]) {
                const partial = await run(bytes, { profile, ...selection });
                expect(full, `${profile} ${JSON.stringify(selection)}`).toEqual(expect.arrayContaining(errors(partial)));
                if (selection.checks === undefined || selection.checks.includes(ID)) {
                    expect(errors(partial), `${profile} ${JSON.stringify(selection)}`).toContain(`${ID}@${HOSTILE}`);
                }
            }
        }
    });

    it("runs no fallback when the rest of the package view is refused too (#413)", async () => {
        // Two spellings of one part collide under ASCII-case folding (#441), so no package view remains.
        const result = await run(
            malformedDocx(
                { name: HOSTILE, content: "<x/>" },
                { name: "word/Extra.xml", content: "<x/>" },
                { name: "word/extra.xml", content: "<x/>" },
            ),
        );
        expect(errors(result)).toEqual(expect.arrayContaining([`${ID}@${HOSTILE}`, "package-extract-failed@"]));
        expect(errors(result)).not.toContain("xml-syntax@word/document.xml");
    });

    it("runs no fallback for an envelope the subset rewrite cannot honor (ZIP64)", async () => {
        const result = await run(
            buildRawZip([...entries({ ...MINIMAL_PARTS, "word/document.xml": MALFORMED_DOCUMENT }), { name: HOSTILE, content: "<x/>" }], {
                zip64: true,
            }),
        );
        expect(errors(result)).toEqual(expect.arrayContaining([`${ID}@${HOSTILE}`, "package-extract-failed@"]));
        expect(errors(result)).not.toContain("xml-syntax@word/document.xml");
    });

    it("reports a refused name in a PPTX package and judges the rest", async () => {
        const result = await run(
            buildRawZip([
                ...entries({ ...MINIMAL_PPTX_PARTS, "ppt/presentation.xml": MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!.slice(0, -20) }),
                { name: "/ppt/abs.xml", content: "<x/>" },
            ]),
            {},
            ".pptx",
        );
        expect(errors(result)).toEqual(
            expect.arrayContaining([`${ID}@/ppt/abs.xml`, "package-extract-failed@", "xml-syntax@ppt/presentation.xml"]),
        );
    });
});
