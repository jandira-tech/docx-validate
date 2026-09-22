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
 * `validateAllRelationshipIds` — rel-TYPE check (`rels-id-mismatch`).
 *
 * Word-first regression. A `<w:hyperlink r:id="X">` whose X RESOLVES in the
 * sidecar but to a NON-hyperlink part (header / footer / fontTable / theme) is
 * accepted by the OOXML XSD (rId target semantics are not schema-constrained),
 * yet real Microsoft Word REFUSES to open the file ("Word found unreadable
 * content"). This was hit in practice by a jubarte redline whose inserted
 * hyperlinks carried plain rIds that collided with the destination package's
 * structural relationships. Populating `DOCXSchemaValidator.elementRelationshipTypes`
 * enables the type check so docx-validate FAILS where Word fails.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function buildDocx(
    dir: string,
    bodyXml: string,
    rels: { id: string; type: string; target: string; external?: boolean }[],
): Promise<DOCXSchemaValidator> {
    await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
    await fs.writeFile(
        path.join(dir, "word", "document.xml"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${bodyXml}</w:body></w:document>`,
        "utf-8",
    );
    const relXml = rels
        .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ""}/>`)
        .join("");
    await fs.writeFile(
        path.join(dir, "word", "_rels", "document.xml.rels"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}">${relXml}</Relationships>`,
        "utf-8",
    );
    return new DOCXSchemaValidator({ unpackedDir: dir });
}

describe("validateAllRelationshipIds — hyperlink/header/footer rId type check", () => {
    it("FLAGS a <w:hyperlink r:id> that points at a header relationship", async () => {
        await withTempDir(async (dir) => {
            const v = await buildDocx(dir, `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`, [
                { id: "rId7", type: `${RT}/header`, target: "header1.xml" },
            ]);
            const result = await v.validateAllRelationshipIds();
            const issue = result.issues.find((i) => i.code === "rels-id-mismatch");
            expect(issue?.severity).toBe("error");
            expect(issue?.message).toContain("hyperlink");
            expect(issue?.message).toContain("header");
        });
    });

    it("ACCEPTS a <w:hyperlink r:id> that points at a hyperlink relationship", async () => {
        await withTempDir(async (dir) => {
            const v = await buildDocx(dir, `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`, [
                {
                    id: "rId7",
                    type: `${RT}/hyperlink`,
                    target: "https://example.com/",
                    external: true,
                },
            ]);
            const result = await v.validateAllRelationshipIds();
            expect(result.issues.some((i) => i.code === "rels-id-mismatch")).toBe(false);
        });
    });

    it("does NOT false-positive on legitimate headerReference/footerReference", async () => {
        await withTempDir(async (dir) => {
            const v = await buildDocx(dir, `<w:p/><w:sectPr><w:headerReference r:id="rId2"/><w:footerReference r:id="rId3"/></w:sectPr>`, [
                { id: "rId2", type: `${RT}/header`, target: "header1.xml" },
                { id: "rId3", type: `${RT}/footer`, target: "footer1.xml" },
            ]);
            const result = await v.validateAllRelationshipIds();
            expect(result.issues.some((i) => i.code === "rels-id-mismatch")).toBe(false);
        });
    });
});
