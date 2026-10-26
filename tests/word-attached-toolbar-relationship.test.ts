/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Session 48 (s48a) specimens for #716: Word 16.115 offers recovery when a
 * `wne:toolbarData r:id` names a relationship whose Type is not exactly
 * `http://schemas.microsoft.com/office/2006/relationships/attachedToolbars`,
 * whatever the target's bytes — clean toolbar bytes are refused once the
 * Type changes. Each retyped specimen is one of the pinned toolbar
 * fixtures — the refused `broken/package.word-attached-toolbar-utf8.docx`
 * and its clean twin `working/…-repaired.docx` — with the Type of
 * `word/_rels/customizations.xml.rels`'s rId1 replaced; the fixtures
 * themselves are the calibrations.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-attached-toolbar-utf8.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-attached-toolbar-utf8-repaired.docx");
const RELS = "word/_rels/customizations.xml.rels";
const REAL = "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";
const CUSTOM = "http://schemas.microsoft.com/office/2006/relationships/custom/attachedToolbars";
const FOREIGN = "http://example.com/attachedToolbars";
const CODE = "word-attached-toolbar-relationship";

/** The fixture with rId1's Type in the customizations sidecar replaced by `type`; every other entry unchanged. */
const retyped = async (file: string, type: string): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(await fs.readFile(file));
    const rels = await zip.file(RELS)!.async("string");
    if (rels.split(`Type="${REAL}"`).length !== 2) throw new Error(`expected exactly one ${REAL} in ${RELS}`);
    zip.file(RELS, rels.replace(`Type="${REAL}"`, `Type="${type}"`));
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};
/** The fixture with rId1's relationship removed from the customizations sidecar. */
const dangling = async (file: string): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(await fs.readFile(file));
    const rels = await zip.file(RELS)!.async("string");
    const rel = `<Relationship Id="rId1" Type="${REAL}" Target="attachedToolbars.bin"/>`;
    if (rels.split(rel).length !== 2) throw new Error(`expected exactly one rId1 relationship in ${RELS}`);
    zip.file(RELS, rels.replace(rel, ""));
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

describe("session 54 (s54c) spellings (#743)", { timeout: 60000 }, () => {
    const UTF8 = "word-attached-toolbar-utf8";
    const TOOLBAR = [CODE, UTF8];

    /** The fixture re-zipped unchanged, or with each `renames` pair applied in order (the session's shape). */
    const renamed = async (file: string, ...renames: readonly [string, string][]): Promise<Uint8Array> => {
        const zip = await JSZip.loadAsync(await fs.readFile(file));
        for (const [from, to] of renames) {
            const entry = zip.file(from);
            if (!entry) throw new Error(`no ${from}`);
            zip.file(to, await entry.async("uint8array"));
            zip.remove(from);
        }
        return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    };
    /** The fixture re-zipped with `part`'s only occurrence of `from` replaced by `to`. */
    const replaced = async (file: string, part: string, from: string, to: string): Promise<Uint8Array> => {
        const zip = await JSZip.loadAsync(await fs.readFile(file));
        const text = await zip.file(part)!.async("string");
        if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from} in ${part}`);
        zip.file(part, text.replace(from, to));
        return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    };
    const run = async (dir: string, name: string, bytes: Uint8Array) => {
        const file = path.join(dir, `${name}.docx`);
        await fs.writeFile(file, bytes);
        return {
            full: await validate(file, { profile: "word-valid" }),
            selected: await validate(file, { profile: "word-valid", checks: TOOLBAR }),
        };
    };

    it("keeps the clean calibration clean re-zipped unchanged", async () => {
        await withTempDir(async (dir) => {
            const { full, selected } = await run(dir, "toolbar-clean-rezipped", await renamed(CLEAN));
            expect(full.valid).toBe(true);
            expect(selected.valid).toBe(true);
        });
    });

    it("keeps the refused calibration refused re-zipped unchanged", async () => {
        await withTempDir(async (dir) => {
            const { full, selected } = await run(dir, "toolbar-refused-rezipped", await renamed(BROKEN));
            expect(full.valid).toBe(false);
            expect(full.issues.some((issue) => issue.severity === "error" && issue.code === UTF8)).toBe(true);
            expect(selected.issues.some((issue) => issue.severity === "error" && issue.code === UTF8)).toBe(true);
        });
    });

    it("reads the sidecar stored as customizations.xml.RELS (s54c: Word opens clean)", async () => {
        await withTempDir(async (dir) => {
            const { full, selected } = await run(
                dir,
                "toolbar-clean-upper-rels",
                await renamed(CLEAN, [RELS, "word/_rels/customizations.xml.RELS"]),
            );
            expect(full.valid).toBe(true);
            expect(selected.valid).toBe(true);
        });
    });

    it("judges customizations.XML and its customizations.XML.rels (s54c: recovery offer)", async () => {
        await withTempDir(async (dir) => {
            const { full, selected } = await run(
                dir,
                "toolbar-refused-upper-xml",
                await renamed(BROKEN, ["word/customizations.xml", "word/customizations.XML"], [RELS, "word/_rels/customizations.XML.rels"]),
            );
            expect(full.valid).toBe(false);
            expect(full.issues.some((issue) => issue.severity === "error" && issue.code === UTF8)).toBe(true);
            expect(selected.issues.some((issue) => issue.severity === "error" && issue.code === UTF8)).toBe(true);
        });
    });

    it("keeps an empty wne:toolbarData r:id clean (s54c)", async () => {
        await withTempDir(async (dir) => {
            const { full, selected } = await run(
                dir,
                "toolbar-clean-empty-id",
                await replaced(CLEAN, "word/customizations.xml", 'r:id="rId1"', 'r:id=""'),
            );
            expect(full.valid).toBe(true);
            expect(selected.valid).toBe(true);
        });
    });

    it.each(["_rels", "_RELS"])("resolves toolbar bytes from a custom/%s sidecar (#758)", async (folder) => {
        const zip = await JSZip.loadAsync(await fs.readFile(BROKEN));
        for (const [from, to] of [
            ["word/customizations.xml", "custom/customizations.xml"],
            [RELS, `custom/${folder}/customizations.xml.rels`],
            ["word/attachedToolbars.bin", "custom/attachedToolbars.bin"],
        ] as const) {
            zip.file(to, await zip.file(from)!.async("uint8array"));
            zip.remove(from);
        }
        const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
        zip.file("word/_rels/document.xml.rels", rels.replace('Target="customizations.xml"', 'Target="../custom/customizations.xml"'));
        const types = await zip.file("[Content_Types].xml")!.async("string");
        zip.file("[Content_Types].xml", types.replace('PartName="/word/customizations.xml"', 'PartName="/custom/customizations.xml"'));
        await withTempDir(async (dir) => {
            const { full, selected } = await run(
                dir,
                `toolbar-custom-${folder}`,
                await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
            );
            for (const result of [full, selected]) {
                expect.soft(result.valid).toBe(false);
                expect
                    .soft(result.issues)
                    .toContainEqual(expect.objectContaining({ severity: "error", code: UTF8, path: "custom/attachedToolbars.bin" }));
            }
        });
    });
});

describe("word-attached-toolbar-relationship (#716; session 48)", { timeout: 60000 }, () => {
    const run = async (dir: string, name: string, bytes: Uint8Array) => {
        const file = path.join(dir, `${name}.docx`);
        await fs.writeFile(file, bytes);
        return { wordValid: await validate(file, { profile: "word-valid" }), strict: await validate(file, { profile: "strict" }) };
    };

    it("keeps the refused-bytes calibration refused", async () => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "toolbar-refused-calibration", await fs.readFile(BROKEN));
            expect(wordValid.valid).toBe(false);
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === "word-attached-toolbar-utf8")).toBe(true);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(false);
        });
    });

    it("keeps the clean-bytes calibration clean", async () => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "toolbar-clean-calibration", await fs.readFile(CLEAN));
            expect(wordValid.valid).toBe(true);
            expect(wordValid.issues.some((issue) => issue.code === CODE)).toBe(false);
            expect(strict.issues.some((issue) => issue.code === CODE)).toBe(false);
        });
    });

    it.each([CUSTOM, FOREIGN])("reports refused toolbar bytes whose rId1 Type is %s", async (type) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "toolbar-refused-retyped", await retyped(BROKEN, type));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
        });
    });

    it.each([CUSTOM, FOREIGN])("reports clean toolbar bytes whose rId1 Type is %s", async (type) => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "toolbar-clean-retyped", await retyped(CLEAN, type));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
        });
    });

    it("reports a toolbarData r:id naming no relationship", async () => {
        await withTempDir(async (dir) => {
            const { wordValid, strict } = await run(dir, "toolbar-dangling-rid", await dangling(CLEAN));
            expect(wordValid.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
            expect(wordValid.valid).toBe(false);
            expect(strict.issues.some((issue) => issue.severity === "error" && issue.code === CODE)).toBe(true);
        });
    });
});
