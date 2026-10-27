/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile, ValidationResult } from "../src/lib/types";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate, type ValidateRunResult } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
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
const ABS_HOSTILE = "/abs/hostile.xml";
const MALFORMED_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.slice(0, -20);
const SCHEMA_DEFECT_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", "<w:bogusDefect/><w:sectPr>");
const PARAID_DOCUMENT = MINIMAL_PARTS["word/document.xml"]!.replace(' w14:paraId="00000001" w14:textId="00000001"', "");
const PROFILES: readonly Profile[] = ["word-valid", "strict", "lenient"];

const entries = (parts: Readonly<Record<string, string>>): RawZipEntry[] =>
    Object.entries(parts).map(([name, content]) => ({ name, content }));
const malformedDocx = (...extra: RawZipEntry[]): Buffer =>
    buildRawZip([...entries({ ...MINIMAL_PARTS, "word/document.xml": MALFORMED_DOCUMENT }), ...extra]);
/** The minimal package with `document` as its main document, plus one refused stored name. */
const hostileDocx = (document: string, options: { zip64?: boolean } = {}): Buffer =>
    buildRawZip(
        [...entries({ ...MINIMAL_PARTS, "word/document.xml": document }), { name: ABS_HOSTILE, content: "<x/>" }],
        options.zip64 === true ? { zip64: true } : {},
    );
/**
 * Put junk between the ZIP64 end record and its locator, a tail no rewrite can
 * move (#479). The record's declared directory size counts the junk, so both
 * readers keep finding the EOCD where the record says it is.
 */
const zip64RecordOffItsLocator = (bytes: Buffer, junk = 5): Buffer => {
    const eocd = bytes.length - 22;
    const locator = eocd - 20;
    const record = locator - 56;
    const out = Buffer.concat([bytes.subarray(0, locator), Buffer.alloc(junk, 0x5a), bytes.subarray(locator)]);
    out.writeBigUInt64LE(out.readBigUInt64LE(record + 40) + BigInt(junk), record + 40);
    return out;
};
const errors = (result: ValidationResult): string[] =>
    result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);
/** The distinct error findings, in order: one malformed part can draw several parser messages. */
const errorPairs = (result: ValidationResult): string[] => [...new Set(errors(result))];

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
        // The retained entries are staged and validated like a clean package's
        // (#629): strict walks the well-formed parts around the malformed one,
        // word-valid compiles no schema, and lenient stops at the malformed
        // part before the walk starts.
        expect(result.xsdWalk.started).toBe(profile === "strict");
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

    it.each(["lenient", "strict"] as const)("%s: the retained entries keep their schema walk", async (profile) => {
        const twin = await run(hostileDocx(SCHEMA_DEFECT_DOCUMENT), { profile });
        expect(twin.valid).toBe(false);
        expect(errors(twin)).toEqual(
            expect.arrayContaining([
                `${ID}@${ABS_HOSTILE}`,
                "package-extract-failed@",
                "xsd-error@word/document.xml",
                ...(profile === "lenient" ? ["xsd-summary@word/document.xml"] : []),
            ]),
        );
        expect(twin.xsdWalk.started).toBe(true);
        expect(twin.xsdWalk.parts).toContain("word/document.xml");
        // The walk judges the retained parts as the clean package of the same content: same findings, message for message.
        const control = await run(await buildDocx({ ...MINIMAL_PARTS, "word/document.xml": SCHEMA_DEFECT_DOCUMENT }), { profile });
        const xsdErrors = (result: ValidationResult): string[] =>
            result.issues.filter((issue) => issue.code === "xsd-error").map((issue) => issue.message);
        expect(xsdErrors(twin)).toEqual(xsdErrors(control));
    });

    it("word-valid: reports a defect only the format validator's own rules find", async () => {
        // No registered check reports a missing w14:paraId, so only the staged
        // run of the retained entries can surface the DOCX validator's rule.
        const twin = await run(hostileDocx(PARAID_DOCUMENT));
        expect(twin.valid).toBe(false);
        expect(twin.issues).toEqual(
            expect.arrayContaining([
                expect.objectContaining({ code: "paraid-missing-element", path: "word/document.xml" }),
                expect.objectContaining({ code: ID, path: ABS_HOSTILE, severity: "error" }),
                expect.objectContaining({ code: "package-extract-failed" }),
            ]),
        );
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

    it("a partial run of every hostile specimen reports no error the full run does not (#381)", async () => {
        const specimens: readonly {
            readonly name: string;
            readonly bytes: () => Buffer;
            /** An error the full run must report, so the subset claim is load-bearing. */
            readonly anchors: Readonly<Record<"word-valid" | "lenient", readonly string[]>>;
        }[] = [
            {
                name: "a schema defect only the schema walk finds",
                bytes: () => hostileDocx(SCHEMA_DEFECT_DOCUMENT),
                anchors: { "word-valid": ["word-unknown-element@word/document.xml"], lenient: ["xsd-error@word/document.xml"] },
            },
            {
                name: "a defect only the format validator finds",
                bytes: () => hostileDocx(PARAID_DOCUMENT),
                anchors: { "word-valid": [], lenient: [] },
            },
            {
                name: "a malformed main in a ZIP64 envelope",
                bytes: () => hostileDocx(MALFORMED_DOCUMENT, { zip64: true }),
                anchors: { "word-valid": ["xml-syntax@word/document.xml"], lenient: ["xml-syntax@word/document.xml"] },
            },
        ];
        for (const { name, bytes, anchors } of specimens) {
            for (const profile of ["word-valid", "lenient"] as const) {
                const tag = `${name} ${profile}`;
                const full = errorPairs(await run(bytes(), { profile }));
                expect(full, tag).toEqual(expect.arrayContaining([`${ID}@${ABS_HOSTILE}`, "package-extract-failed@", ...anchors[profile]]));
                for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }, { checks: [ID] }]) {
                    const partial = errorPairs(await run(bytes(), { profile, ...selection }));
                    expect(full, `${tag} ${JSON.stringify(selection)}`).toEqual(expect.arrayContaining(partial));
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

    it.each(PROFILES)("judges the retained entries of a ZIP64 envelope (%s)", async (profile) => {
        const result = await run(hostileDocx(MALFORMED_DOCUMENT, { zip64: true }), { profile });
        expect(result.valid).toBe(false);
        expect(errors(result)).toEqual(
            expect.arrayContaining([`${ID}@${ABS_HOSTILE}`, "package-extract-failed@", "xml-syntax@word/document.xml"]),
        );
    });

    it("runs no fallback for a ZIP64 tail the subset rewrite cannot move", async () => {
        const result = await run(zip64RecordOffItsLocator(hostileDocx(MALFORMED_DOCUMENT, { zip64: true })));
        expect(errors(result)).toEqual(expect.arrayContaining([`${ID}@${ABS_HOSTILE}`, "package-extract-failed@"]));
        expect(errors(result)).not.toContain("xml-syntax@word/document.xml");
    });

    it("a partial run judges the rest of the package too, as the full run does", async () => {
        for (const profile of ["word-valid", "lenient"] as const) {
            const tag = profile;
            const selected = await run(malformedDocx({ name: HOSTILE, content: "<x/>" }), { profile, checks: ["xml-syntax"] });
            // The selection does not name the refusal check, so only the extract failure says why.
            expect(errorPairs(selected), tag).toEqual(["xml-syntax@word/document.xml", "package-extract-failed@"]);
            expect(selected.metadata?.executedChecks, tag).toEqual(["xml-syntax"]);
            expect(selected.metadata?.omittedChecks, tag).not.toContain("xml-syntax");

            const quick = await run(malformedDocx({ name: HOSTILE, content: "<x/>" }), { profile, quick: true });
            expect(errorPairs(quick), tag).toEqual(
                expect.arrayContaining([`${ID}@${HOSTILE}`, "xml-syntax@word/document.xml", "package-extract-failed@"]),
            );
            // Every applicable id is in exactly one list (#420).
            const executed = quick.metadata?.executedChecks ?? [];
            const omitted = quick.metadata?.omittedChecks ?? [];
            expect(
                executed.filter((id) => omitted.includes(id)),
                tag,
            ).toEqual([]);
            expect(new Set([...executed, ...omitted]), tag).toEqual(
                new Set(
                    createDefaultCheckRegistry()
                        .list("docx")
                        .map((definition) => definition.id),
                ),
            );
        }
    });

    it("a partial run naming only the refusal check runs no part check", async () => {
        const result = await run(malformedDocx({ name: HOSTILE, content: "<x/>" }), { checks: [ID] });
        expect(errorPairs(result)).toEqual([`${ID}@${HOSTILE}`, "package-extract-failed@"]);
        expect(result.metadata?.executedChecks).toEqual([ID]);
    });

    it("a partial run judges the entry a later hostile name hid", async () => {
        const hidesMalformed = await run(malformedDocx({ name: "word/x/../document.xml", content: MINIMAL_PARTS["word/document.xml"]! }), {
            checks: ["xml-syntax"],
        });
        expect(errorPairs(hidesMalformed)).toEqual(["xml-syntax@word/document.xml", "package-extract-failed@"]);

        const hidesWellFormed = await run(
            buildRawZip([...entries(MINIMAL_PARTS), { name: "word/x/../document.xml", content: MALFORMED_DOCUMENT }]),
            {
                checks: ["xml-syntax"],
            },
        );
        expect(errorPairs(hidesWellFormed)).toEqual(["package-extract-failed@"]);
        expect(hidesWellFormed.metadata?.executedChecks).toEqual(["xml-syntax"]);
    });

    it("a partial run runs no fallback when the rest of the package view is refused too (#413)", async () => {
        const result = await run(
            malformedDocx(
                { name: HOSTILE, content: "<x/>" },
                { name: "word/Extra.xml", content: "<x/>" },
                { name: "word/extra.xml", content: "<x/>" },
            ),
            { checks: ["xml-syntax"] },
        );
        expect(errorPairs(result)).toEqual(["package-extract-failed@"]);
        // The selected check could not run, so it is omitted, never in neither list (#420).
        expect(result.metadata?.executedChecks).toEqual([]);
        expect(result.metadata?.omittedChecks).toContain("xml-syntax");
    });

    it("a partial run of a PPTX package judges the rest", async () => {
        const result = await run(
            buildRawZip([
                ...entries({ ...MINIMAL_PPTX_PARTS, "ppt/presentation.xml": MINIMAL_PPTX_PARTS["ppt/presentation.xml"]!.slice(0, -20) }),
                { name: "/ppt/abs.xml", content: "<x/>" },
            ]),
            { checks: ["xml-syntax"] },
            ".pptx",
        );
        expect(errorPairs(result)).toEqual(["xml-syntax@ppt/presentation.xml", "package-extract-failed@"]);
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
