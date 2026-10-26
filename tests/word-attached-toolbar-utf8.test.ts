/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
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

describe("word-valid attached toolbar UTF-8 replacement", { timeout: 30000 }, () => {
    it("rejects a toolbarData target that contains EF BF BD", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        const errors = result.issues.filter((issue) => issue.severity === "error");
        expect(errors.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(true);
    });

    it("does not reject the toolbar once those sequences are gone", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-attached-toolbar-utf8")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-attached-toolbar-utf8")).toBe(true);
    });
});

// The sidecar lookup reads an escaped spelling (#758); the folder its
// targets resolve against must be picked by the same equivalence, not an
// ASCII-case fold, or the escaped spellings resolve inside the sidecar's
// own folder and name no part (#772).
describe("escaped _rels sidecar folders (#772)", { timeout: 60000 }, () => {
    it.each(["_rels", "_RELS", "_%72els", "%5Frels", "_r%65ls"])("resolves toolbar bytes from a custom/%s sidecar", async (folder) => {
        const zip = await JSZip.loadAsync(await fs.readFile(BROKEN));
        for (const [from, to] of [
            ["word/customizations.xml", "custom/customizations.xml"],
            ["word/_rels/customizations.xml.rels", `custom/${folder}/customizations.xml.rels`],
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
            const file = path.join(dir, `toolbar-custom-${folder.replace(/%/g, "pct")}.docx`);
            await fs.writeFile(file, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
            const result = await validate(file, { checks: ["word-attached-toolbar-utf8"] });
            expect.soft(result.valid).toBe(false);
            expect
                .soft(result.issues)
                .toContainEqual(
                    expect.objectContaining({ severity: "error", code: "word-attached-toolbar-utf8", path: "custom/attachedToolbars.bin" }),
                );
        });
    });
});
