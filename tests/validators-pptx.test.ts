/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { PPTXSchemaValidator, PRESENTATIONML_NAMESPACE } from "../src/scripts/office/validators/pptx";

const PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function writeFile(file: string, contents: string): Promise<void> {
    await fs.mkdir(path.dirname(file), { recursive: true });
    await fs.writeFile(file, contents, "utf8");
}

function presentationXml(parts: string): string {
    return (
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<p:presentation xmlns:p="${PRESENTATIONML_NAMESPACE}" xmlns:r="${R_NS}">${parts}</p:presentation>`
    );
}

function relsXml(rels: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` + `<Relationships xmlns="${PR_NS}">${rels}</Relationships>`;
}

function makeValidator(unpackedDir: string): PPTXSchemaValidator {
    return new PPTXSchemaValidator({ unpackedDir, verbose: false });
}

describe("PPTXSchemaValidator.validateUuidIds", () => {
    it("passes when no UUID-shaped IDs are present", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "presentation.xml"),
                presentationXml('<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>'),
            );
            const result = await makeValidator(dir).validateUuidIds();
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("passes when UUID-shaped IDs contain valid hex", async () => {
        await withTempDir(async (dir) => {
            const validUuid = "{12345678-1234-5678-9ABC-123456789ABC}";
            await writeFile(path.join(dir, "ppt", "slides", "slide1.xml"), presentationXml(`<p:custom durableId="${validUuid}"/>`));
            const result = await makeValidator(dir).validateUuidIds();
            expect(result.valid).toBe(true);
        });
    });

    it("flags UUID-shaped IDs with non-hex characters", async () => {
        await withTempDir(async (dir) => {
            // 32 alnum chars (no hyphens, contains 'g'/'z' which are non-hex)
            const badUuid = "1234567890abcdefGHIJKLMNOpqrstuv";
            await writeFile(path.join(dir, "ppt", "slides", "slide1.xml"), presentationXml(`<p:custom paraId="${badUuid}"/>`));
            const result = await makeValidator(dir).validateUuidIds();
            expect(result.valid).toBe(false);
            expect(result.issues[0].message).toContain(badUuid);
            expect(result.issues[0].message).toContain("invalid hex characters");
        });
    });
});

describe("PPTXSchemaValidator.validateSlideLayoutIds", () => {
    it("passes when no slide masters exist", async () => {
        await withTempDir(async (dir) => {
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.valid).toBe(true);
            expect(result.issues).toHaveLength(0);
        });
    });

    it("passes when each sldLayoutId resolves to a slideLayout relationship", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slideMasters", "slideMaster1.xml"),
                presentationXml(
                    "<p:sldLayoutIdLst>" +
                        '<p:sldLayoutId id="2147483649" r:id="rId1"/>' +
                        '<p:sldLayoutId id="2147483650" r:id="rId2"/>' +
                        "</p:sldLayoutIdLst>",
                ),
            );
            await writeFile(
                path.join(dir, "ppt", "slideMasters", "_rels", "slideMaster1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
                        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.valid).toBe(true);
        });
    });

    it("flags sldLayoutId references that are missing from the rels file", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slideMasters", "slideMaster1.xml"),
                presentationXml('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdMissing"/></p:sldLayoutIdLst>'),
            );
            await writeFile(
                path.join(dir, "ppt", "slideMasters", "_rels", "slideMaster1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.valid).toBe(false);
            expect(result.issues[0].message).toContain("rIdMissing");
            expect(result.issues[0].code).toBe("sldlayout-missing-rid");
        });
    });

    it("flags missing rels file beside a slide master", async () => {
        await withTempDir(async (dir) => {
            await writeFile(path.join(dir, "ppt", "slideMasters", "slideMaster1.xml"), presentationXml("<p:sldLayoutIdLst/>"));
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.valid).toBe(false);
            expect(result.issues[0].code).toBe("sldlayout-missing-rels");
        });
    });
});

describe("PPTXSchemaValidator.validateNoDuplicateSlideLayouts", () => {
    it("passes when slides reference exactly one layout each", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateNoDuplicateSlideLayouts();
            expect(result.valid).toBe(true);
        });
    });

    it("flags slides referencing more than one slideLayout", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
                        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateNoDuplicateSlideLayouts();
            expect(result.valid).toBe(false);
            expect(result.issues[0].message).toContain("2 slideLayout references");
        });
    });
});

describe("PPTXSchemaValidator.validateNotesSlideReferences", () => {
    it("passes when each notes slide is referenced by a single slide", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>',
                ),
            );
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide2.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateNotesSlideReferences();
            expect(result.valid).toBe(true);
        });
    });

    it("passes when the slides directory is missing", async () => {
        await withTempDir(async (dir) => {
            const result = await makeValidator(dir).validateNotesSlideReferences();
            expect(result.valid).toBe(true);
        });
    });

    it("flags notes slides referenced by multiple parent slides", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>',
                ),
            );
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide2.xml.rels"),
                relsXml(
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>',
                ),
            );
            const result = await makeValidator(dir).validateNotesSlideReferences();
            expect(result.valid).toBe(false);
            const sharedIssue = result.issues.find((i) => i.code === "notes-shared");
            expect(sharedIssue).toBeDefined();
            expect(sharedIssue!.message).toContain("notesSlides/notesSlide1.xml");
            expect(sharedIssue!.message).toContain("slide1");
            expect(sharedIssue!.message).toContain("slide2");
        });
    });
});

describe("PPTXSchemaValidator.validate", () => {
    it("aggregates validate_xml failure and short-circuits the suite", async () => {
        await withTempDir(async (dir) => {
            // Malformed XML so the first check (validateXml) fails immediately.
            await writeFile(path.join(dir, "ppt", "presentation.xml"), "<p:not closed");
            // Need [Content_Types].xml — but the malformed XML check fires first
            // and short-circuits, so the absence here is not exercised.
            const result = await makeValidator(dir).validate();
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xml-syntax")).toBe(true);
        });
    });

    it("subclasses BaseSchemaValidator and exposes the PPTX element relationship table", () => {
        const validator = makeValidator(".");
        // protected accessor — assert via behavior of _getExpectedRelationshipType (inherited)
        const exp = validator as unknown as {
            _getExpectedRelationshipType(n: string): string | null;
        };
        expect(exp._getExpectedRelationshipType("sldId")).toBe("slide");
        expect(exp._getExpectedRelationshipType("themeId")).toBe("theme");
        expect(exp._getExpectedRelationshipType("sldLayoutId")).toBe("slidelayout");
    });
});
