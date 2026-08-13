/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

// Two Word-blocking defects discovered by root-causing the word-invalid-fixtures
// false-negatives against REAL Microsoft Word (minimal-repair + re-probe):
//   1. A foreign OpenXmlPowerTools DocumentBuilder <Insert> directive left in a
//      header/footer/body part -> Word "experienced an error" (OPEN_ERROR).
//      Removing the element flips the file to clean. 0/350 Word-clean files have it.
//   2. A docProps/app.xml boolean property (ScaleCrop/LinksUpToDate/SharedDoc/
//      HyperlinksChanged) whose value carries whitespace (e.g. "false\n  ") ->
//      Word "unreadable content". Stripping the boolean whitespace flips the file
//      to clean (integer whitespace like <TotalTime>0\n  </TotalTime> is tolerated
//      by Word and present in clean files, so it is intentionally NOT flagged).

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const EP_NS = `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"`;
const VT_NS = `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"`;
const DB_NS = `xmlns="http://powertools.codeplex.com/documentbuilder/2011/insert"`;

async function write(p: string, content: string): Promise<void> {
    await fs.mkdir(path.dirname(p), { recursive: true });
    await fs.writeFile(p, content, "utf-8");
}

describe("validateDocumentBuilderInserts", () => {
    it("flags an unresolved DocumentBuilder <Insert> in a header part", async () => {
        await withTempDir(async (dir) => {
            await write(
                path.join(dir, "word", "header1.xml"),
                `<?xml version="1.0"?><w:hdr ${W_NS}><Insert Id="Templafy" ${DB_NS}/></w:hdr>`,
            );
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocumentBuilderInserts();
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("documentbuilder-insert-unresolved");
            expect(result.issues[0].severity).toBe("error");
            expect(result.issues[0].path).toBe("word/header1.xml");
        });
    });

    it("flags it in a footer part too", async () => {
        await withTempDir(async (dir) => {
            await write(
                path.join(dir, "word", "footer2.xml"),
                `<?xml version="1.0"?><w:ftr ${W_NS}><Insert Id="Templafy" ${DB_NS}/></w:ftr>`,
            );
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocumentBuilderInserts();
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("documentbuilder-insert-unresolved");
        });
    });

    it("passes a normal header with no DocumentBuilder directive", async () => {
        await withTempDir(async (dir) => {
            await write(path.join(dir, "word", "header1.xml"), `<?xml version="1.0"?><w:hdr ${W_NS}><w:p/></w:hdr>`);
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocumentBuilderInserts();
            expect(result.valid).toBe(true);
        });
    });

    it("does not flag the DocumentBuilder URI when it appears only as text, not as a namespace", async () => {
        await withTempDir(async (dir) => {
            await write(
                path.join(dir, "word", "header1.xml"),
                `<?xml version="1.0"?><w:hdr ${W_NS}><w:p><w:r><w:t>see http://powertools.codeplex.com/documentbuilder</w:t></w:r></w:p></w:hdr>`,
            );
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocumentBuilderInserts();
            expect(result.valid).toBe(true);
        });
    });
});

describe("validateDocPropsBooleans", () => {
    const app = (body: string): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties ${EP_NS} ${VT_NS}>${body}</Properties>`;

    it("flags a boolean property with trailing whitespace", async () => {
        await withTempDir(async (dir) => {
            await write(path.join(dir, "docProps", "app.xml"), app(`<ScaleCrop>false\n  </ScaleCrop>`));
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocPropsBooleans();
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("docprops-boolean-invalid");
            expect(result.issues[0].severity).toBe("error");
        });
    });

    it("flags every affected boolean property", async () => {
        await withTempDir(async (dir) => {
            await write(
                path.join(dir, "docProps", "app.xml"),
                app(`<ScaleCrop>false\n  </ScaleCrop><LinksUpToDate>false\n  </LinksUpToDate><SharedDoc>false\n  </SharedDoc><HyperlinksChanged>false\n  </HyperlinksChanged>`),
            );
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocPropsBooleans();
            expect(result.valid).toBe(false);
            expect(result.issues.length).toBe(4);
        });
    });

    it("passes canonical boolean values", async () => {
        await withTempDir(async (dir) => {
            await write(
                path.join(dir, "docProps", "app.xml"),
                app(`<ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>true</HyperlinksChanged>`),
            );
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocPropsBooleans();
            expect(result.valid).toBe(true);
        });
    });

    it("does NOT flag integer-property whitespace (Word tolerates it; present in clean files)", async () => {
        await withTempDir(async (dir) => {
            await write(path.join(dir, "docProps", "app.xml"), app(`<TotalTime>0\n  </TotalTime><Pages>1\n  </Pages>`));
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            const result = await v.validateDocPropsBooleans();
            expect(result.valid).toBe(true);
        });
    });
});
