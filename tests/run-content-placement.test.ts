/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * `validateRunContentPlacement`: EG_RunInnerContent (<w:tab>, <w:br>,
 * <w:drawing>, <w:pict>, <w:object>, <w:t>, …) is valid only inside <w:r>. As a
 * direct child of <w:p> it is dropped by Word on open (data loss). The rule
 * reports `run-content-misplaced` at error severity and is Word-blocking.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function validatorFor(dir: string, bodyXml: string): Promise<DOCXSchemaValidator> {
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.writeFile(
        path.join(dir, "word", "document.xml"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${bodyXml}</w:body></w:document>`,
        "utf-8",
    );
    return new DOCXSchemaValidator({ unpackedDir: dir });
}

describe("validateRunContentPlacement", () => {
    it("flags a bare <w:tab/> placed directly under <w:p>", async () => {
        await withTempDir(async (dir) => {
            const v = await validatorFor(dir, "<w:p><w:r><w:t>a</w:t></w:r><w:tab/></w:p>");
            const result = await v.validateRunContentPlacement();
            expect(result.valid).toBe(false);
            const issue = result.issues.find((i) => i.code === "run-content-misplaced");
            expect(issue?.severity).toBe("error");
            expect(issue?.message).toContain("<w:tab>");
        });
    });

    it("flags a bare <w:drawing> placed directly under <w:p>", async () => {
        await withTempDir(async (dir) => {
            const v = await validatorFor(dir, "<w:p><w:drawing/></w:p>");
            const result = await v.validateRunContentPlacement();
            expect(result.issues.some((i) => i.code === "run-content-misplaced")).toBe(true);
        });
    });

    it("accepts run-inner content correctly wrapped in <w:r>", async () => {
        await withTempDir(async (dir) => {
            const v = await validatorFor(
                dir,
                "<w:p><w:r><w:t>a</w:t></w:r><w:r><w:tab/></w:r><w:r><w:drawing/></w:r></w:p>",
            );
            const result = await v.validateRunContentPlacement();
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("does not flag legitimate paragraph-level siblings of runs (hyperlink, bookmarks, range markers)", async () => {
        await withTempDir(async (dir) => {
            const v = await validatorFor(
                dir,
                '<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:hyperlink r:id="rId1"><w:r><w:t>x</w:t></w:r></w:hyperlink><w:commentRangeStart w:id="0"/><w:bookmarkEnd w:id="0"/></w:p>',
            );
            const result = await v.validateRunContentPlacement();
            expect(result.issues.some((i) => i.code === "run-content-misplaced")).toBe(false);
        });
    });
});
