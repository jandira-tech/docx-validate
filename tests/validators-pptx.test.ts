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
import { validate } from "../src/scripts/office/validate";
import { PPTXSchemaValidator, PRESENTATIONML_NAMESPACE } from "../src/scripts/office/validators/pptx";
import { buildPptx, MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

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

describe("direct PPTX relationships outside legacy lenient (#434)", () => {
    it("does not resolve a slide layout ID through a nested relationship", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt/slideMasters/slideMaster1.xml"),
                `<p:sldMaster xmlns:p="${PRESENTATIONML_NAMESPACE}" xmlns:r="${R_NS}"><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="nested"/></p:sldLayoutIdLst></p:sldMaster>`,
            );
            await writeFile(
                path.join(dir, "ppt/slideMasters/_rels/slideMaster1.xml.rels"),
                relsXml(
                    `<x:Wrap xmlns:x="urn:x"><Relationship Id="nested" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></x:Wrap>`,
                ),
            );
            const strict = await new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" }).validateSlideLayoutIds();
            expect(strict.issues).toContainEqual(expect.objectContaining({ code: "sldlayout-missing-rid", severity: "error" }));
            expect((await makeValidator(dir).validateSlideLayoutIds()).issues).toEqual([]);
        });
    });

    it.each([
        ["validateNoDuplicateSlideLayouts", "slideLayout", "slide-layout-duplicate"],
        ["validateNotesSlideReferences", "notesSlide", "notes-shared"],
    ] as const)("%s ignores a nested second reference and preserves lenient behavior", async (method, type, code) => {
        await withTempDir(async (dir) => {
            const relationship = `<Relationship Id="nested" Type="${R_NS}/${type}" Target="../parts/shared.xml"/>`;
            await writeFile(
                path.join(dir, "ppt/slides/_rels/slide1.xml.rels"),
                relsXml(relationship + `<x:Wrap xmlns:x="urn:x">${relationship}</x:Wrap>`),
            );
            const strict = await new PPTXSchemaValidator({ unpackedDir: dir, profile: "strict" })[method]();
            expect(strict.issues.filter((issue) => issue.code === code)).toEqual([]);
            const lenient = await makeValidator(dir)[method]();
            expect(lenient.issues).toContainEqual(expect.objectContaining({ code, severity: "error" }));
        });
    });
});

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

    // #1007: the key is the part the target names, under OPC part-name
    // equivalence, so two slides sharing one notes slide through equivalent
    // spellings are reported, under the part's name.
    it.each([
        "../../ppt/notesSlides/notesSlide1.xml",
        "/ppt/notesSlides/notesSlide1.xml",
        "../notesSlides//notesSlide1.xml",
        "../notesSlides/notesSlide1.xml?x",
        "../notesSlides/notesSlide1.xml#x",
        "..\\notesSlides\\notesSlide1.xml",
        "../notesSlides/NotesSlide1.xml",
        "../notesSlides/notesSlide%31.xml",
    ])("flags a notes slide two slides reference through the spelling %s and the plain one", async (target) => {
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
                    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${target}"/>`,
                ),
            );
            const result = await makeValidator(dir).validateNotesSlideReferences();
            expect(result.valid).toBe(false);
            expect(result.issues.filter((i) => i.code === "notes-shared").map((i) => i.message)).toEqual([
                "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2",
            ]);
        });
    });

    it("keeps two notes slides two slides reference apart", async () => {
        await withTempDir(async (dir) => {
            for (const n of [1, 2]) {
                await writeFile(
                    path.join(dir, "ppt", "slides", "_rels", `slide${n}.xml.rels`),
                    relsXml(
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${n}.xml"/>`,
                    ),
                );
            }
            const result = await makeValidator(dir).validateNotesSlideReferences();
            expect(result.issues).toEqual([]);
            expect(result.valid).toBe(true);
        });
    });

    // #991: a target whose last segment is `.` or `..`, or that ends in `/`,
    // names a folder, which names no part (#962, #975), so it shares no notes
    // slide with the plain target the stored part answers.
    it.each(["../notesSlides/notesSlide1.xml/.", "../notesSlides/notesSlide1.xml/x/..", "../notesSlides/notesSlide1.xml/"])(
        "does not count the folder target %s as a reference to the notes slide its path spells",
        async (target) => {
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
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${target}"/>`,
                    ),
                );
                const result = await makeValidator(dir).validateNotesSlideReferences();
                expect(result.issues).toEqual([]);
                expect(result.valid).toBe(true);
            });
        },
    );
});

// #1013: OPC lets a presentation store its slides and masters under any part
// name, so the slide and master checks judge every one the presentation
// relates, beside the historical `ppt/slides/_rels` and `ppt/slideMasters`.
describe("slides and masters the presentation relates outside the historical folders (#1013)", () => {
    const relationship = (id: string, type: string, target: string) =>
        `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;

    /** `_rels/.rels` naming `presentation` as the main part, whose sidecar relates each target by `type`. */
    async function presentationRelating(
        dir: string,
        type: string,
        targets: string[],
        presentation = "ppt/presentation.xml",
    ): Promise<void> {
        await writeFile(path.join(dir, "_rels", ".rels"), relsXml(relationship("rId1", "officeDocument", presentation)));
        await writeFile(path.join(dir, presentation), presentationXml(""));
        await writeFile(
            path.join(dir, path.dirname(presentation), "_rels", `${path.basename(presentation)}.rels`),
            relsXml(targets.map((target, i) => relationship(`rId${i + 1}`, type, target)).join("")),
        );
    }

    const notesShared = async (dir: string) =>
        (await makeValidator(dir).validateNotesSlideReferences()).issues.map((issue) => [issue.code, issue.message]);

    it("flags a notes slide two slides stored in ppt/myslides share", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slide", ["myslides/slide1.xml", "myslides/slide2.xml"]);
            for (const n of [1, 2]) {
                await writeFile(path.join(dir, "ppt", "myslides", `slide${n}.xml`), presentationXml(""));
                await writeFile(
                    path.join(dir, "ppt", "myslides", "_rels", `slide${n}.xml.rels`),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect(await notesShared(dir)).toEqual([
                ["notes-shared", "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2"],
                ["notes-shared-detail", "    - ppt/myslides/_rels/slide1.xml.rels"],
                ["notes-shared-detail", "    - ppt/myslides/_rels/slide2.xml.rels"],
            ]);
        });
    });

    it("resolves each related slide's notes target against that slide's folder, beside a ppt/slides sidecar", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slide", ["deck/a/slide2.xml"]);
            await writeFile(path.join(dir, "ppt", "deck", "a", "slide2.xml"), presentationXml(""));
            await writeFile(
                path.join(dir, "ppt", "deck", "a", "_rels", "slide2.xml.rels"),
                relsXml(relationship("rId1", "notesSlide", "../../notesSlides/notesSlide1.xml")),
            );
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
            );
            expect(await notesShared(dir)).toEqual([
                ["notes-shared", "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide2, slide1"],
                ["notes-shared-detail", "    - ppt/deck/a/_rels/slide2.xml.rels"],
                ["notes-shared-detail", "    - ppt/slides/_rels/slide1.xml.rels"],
            ]);
        });
    });

    it("follows a renamed presentation and an equivalent spelling of the slide's name", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slide", ["../ppt/MySlides/Slide1.xml", "/ppt/myslides/slide2.xml"], "deck/main.xml");
            for (const n of [1, 2]) {
                await writeFile(path.join(dir, "ppt", "myslides", `slide${n}.xml`), presentationXml(""));
                await writeFile(
                    path.join(dir, "ppt", "myslides", "_rels", `slide${n}.xml.rels`),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect((await notesShared(dir))[0]).toEqual([
                "notes-shared",
                "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2",
            ]);
        });
    });

    it("judges a slide both lists name once", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slide", ["slides/slide1.xml"]);
            await writeFile(path.join(dir, "ppt", "slides", "slide1.xml"), presentationXml(""));
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(
                    relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml") +
                        relationship("rId2", "slideLayout", "../slideLayouts/slideLayout1.xml") +
                        relationship("rId3", "slideLayout", "../slideLayouts/slideLayout2.xml"),
                ),
            );
            expect(await notesShared(dir)).toEqual([]);
            expect((await makeValidator(dir).validateNoDuplicateSlideLayouts()).issues.map((issue) => issue.path)).toEqual([
                "ppt/slides/_rels/slide1.xml.rels",
            ]);
        });
    });

    it("reads only slide relationships as slides", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slideLayout", ["myslides/slide1.xml", "myslides/slide2.xml"]);
            for (const n of [1, 2]) {
                await writeFile(path.join(dir, "ppt", "myslides", `slide${n}.xml`), presentationXml(""));
                await writeFile(
                    path.join(dir, "ppt", "myslides", "_rels", `slide${n}.xml.rels`),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect(await notesShared(dir)).toEqual([]);
        });
    });

    it("flags a related slide in ppt/myslides with two slideLayout references", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slide", ["myslides/slide1.xml"]);
            await writeFile(path.join(dir, "ppt", "myslides", "slide1.xml"), presentationXml(""));
            await writeFile(
                path.join(dir, "ppt", "myslides", "_rels", "slide1.xml.rels"),
                relsXml(
                    relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml") +
                        relationship("rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"),
                ),
            );
            const result = await makeValidator(dir).validateNoDuplicateSlideLayouts();
            expect(result.issues.map((issue) => [issue.code, issue.message, issue.path])).toEqual([
                ["slide-layout-duplicate", "has 2 slideLayout references", "ppt/myslides/_rels/slide1.xml.rels"],
            ]);
            expect(result.valid).toBe(false);
        });
    });

    it("flags a related master in ppt/masters whose sldLayoutId names no slideLayout relationship", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slideMaster", ["masters/slideMaster1.xml"]);
            await writeFile(
                path.join(dir, "ppt", "masters", "slideMaster1.xml"),
                presentationXml('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId9"/></p:sldLayoutIdLst>'),
            );
            await writeFile(path.join(dir, "ppt", "masters", "_rels", "slideMaster1.xml.rels"), relsXml(""));
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.issues.map((issue) => [issue.code, issue.path])).toEqual([
                ["sldlayout-missing-rid", "ppt/masters/slideMaster1.xml"],
            ]);
            expect(result.valid).toBe(false);
        });
    });

    it("flags a related master in ppt/masters with no relationships part", async () => {
        await withTempDir(async (dir) => {
            await presentationRelating(dir, "slideMaster", ["masters/slideMaster1.xml"]);
            await writeFile(path.join(dir, "ppt", "masters", "slideMaster1.xml"), presentationXml("<p:sldLayoutIdLst/>"));
            const result = await makeValidator(dir).validateSlideLayoutIds();
            expect(result.issues.map((issue) => [issue.code, issue.path])).toEqual([
                ["sldlayout-missing-rels", "ppt/masters/slideMaster1.xml"],
            ]);
        });
    });
});

// #1017: the historical lists judge every stored spelling they match, ASCII
// case-insensitively, as OPC compares part names. Before #1013 the file-system
// walk judged both spellings of an equivalent pair, and on macOS a sidecar
// under a case-variant folder too.
describe("historical sidecars and masters under every stored spelling (#1017)", () => {
    const relationship = (id: string, type: string, target: string) =>
        `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;
    const notesShared = async (dir: string) =>
        (await makeValidator(dir).validateNotesSlideReferences()).issues.map((issue) => [issue.code, issue.message]);

    it("judges both stored spellings of one sidecar", async () => {
        await withTempDir(async (dir) => {
            for (const name of ["slide1.xml.rels", "slide%31.xml.rels"]) {
                await writeFile(
                    path.join(dir, "ppt", "slides", "_rels", name),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect(await notesShared(dir)).toEqual([
                ["notes-shared", "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide%31, slide1"],
                ["notes-shared-detail", "    - ppt/slides/_rels/slide%31.xml.rels"],
                ["notes-shared-detail", "    - ppt/slides/_rels/slide1.xml.rels"],
            ]);
        });
    });

    it("flags the second stored spelling's two slideLayout references", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"),
                relsXml(relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")),
            );
            await writeFile(
                path.join(dir, "ppt", "slides", "_rels", "slide%31.xml.rels"),
                relsXml(
                    relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml") +
                        relationship("rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"),
                ),
            );
            expect((await makeValidator(dir).validateNoDuplicateSlideLayouts()).issues.map((issue) => issue.path)).toEqual([
                "ppt/slides/_rels/slide%31.xml.rels",
            ]);
        });
    });

    it("judges an orphan sidecar pair under ppt/Slides/_rels", async () => {
        await withTempDir(async (dir) => {
            for (const n of [1, 2]) {
                await writeFile(
                    path.join(dir, "ppt", "Slides", "_rels", `slide${n}.xml.rels`),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect(await notesShared(dir)).toEqual([
                ["notes-shared", "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2"],
                ["notes-shared-detail", "    - ppt/Slides/_rels/slide1.xml.rels"],
                ["notes-shared-detail", "    - ppt/Slides/_rels/slide2.xml.rels"],
            ]);
        });
    });

    it("judges a sidecar whose extension is spelled in upper case", async () => {
        await withTempDir(async (dir) => {
            for (const name of ["slide1.XML.rels", "slide2.xml.rels"]) {
                await writeFile(
                    path.join(dir, "ppt", "slides", "_rels", name),
                    relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml")),
                );
            }
            expect((await notesShared(dir))[0]).toEqual([
                "notes-shared",
                "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2",
            ]);
        });
    });

    it("judges a master under ppt/SlideMasters", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "SlideMasters", "slideMaster1.xml"),
                presentationXml('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId9"/></p:sldLayoutIdLst>'),
            );
            await writeFile(path.join(dir, "ppt", "SlideMasters", "_rels", "slideMaster1.xml.rels"), relsXml(""));
            expect((await makeValidator(dir).validateSlideLayoutIds()).issues.map((issue) => [issue.code, issue.path])).toEqual([
                ["sldlayout-missing-rid", "ppt/SlideMasters/slideMaster1.xml"],
            ]);
        });
    });

    it("keeps a lenient run invalid when two spellings of one sidecar share a notes slide", async () => {
        await withTempDir(async (dir) => {
            const sidecar = relsXml(
                relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml") +
                    relationship("rId2", "notesSlide", "../notesSlides/notesSlide1.xml"),
            );
            const notesSlide =
                `<?xml version="1.0" encoding="UTF-8"?>\n` +
                `<p:notes xmlns:p="${PRESENTATIONML_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}">` +
                `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>`;
            const file = path.join(dir, "deck.pptx");
            await fs.writeFile(
                file,
                await buildPptx({
                    ...MINIMAL_PPTX_PARTS,
                    "ppt/slides/_rels/slide1.xml.rels": sidecar,
                    "ppt/slides/_rels/slide%31.xml.rels": sidecar,
                    "ppt/notesSlides/notesSlide1.xml": notesSlide,
                    "ppt/notesSlides/_rels/notesSlide1.xml.rels": relsXml(relationship("rId1", "slide", "../slides/slide1.xml")),
                    "[Content_Types].xml": (MINIMAL_PPTX_PARTS["[Content_Types].xml"] ?? "").replace(
                        "</Types>",
                        '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>',
                    ),
                }),
            );
            const result = await validate(file, { profile: "lenient" });
            expect(result.issues.filter((issue) => issue.code === "notes-shared").map((issue) => issue.message)).toEqual([
                "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide%31, slide1",
            ]);
            expect(result.valid).toBe(false);
        });
    });
});

// #1037: the historical folders are matched as OPC compares part names, so a
// folder segment spelled with an escaped unreserved character (`%5Frels`,
// `%73lides`, `slide%4Dasters`) is inside them, as `canonicalPartName` reads it.
describe("historical folders spelled with escaped unreserved characters (#1037)", () => {
    const relationship = (id: string, type: string, target: string) =>
        `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;
    const notesSidecar = () => relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml"));

    it("judges an orphan sidecar pair under ppt/slides/%5Frels", async () => {
        await withTempDir(async (dir) => {
            for (const n of [1, 2]) {
                await writeFile(path.join(dir, "ppt", "slides", "%5Frels", `slide${n}.xml.rels`), notesSidecar());
            }
            expect((await makeValidator(dir).validateNotesSlideReferences()).issues.map((issue) => [issue.code, issue.message])).toEqual([
                ["notes-shared", "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2"],
                ["notes-shared-detail", "    - ppt/slides/%5Frels/slide1.xml.rels"],
                ["notes-shared-detail", "    - ppt/slides/%5Frels/slide2.xml.rels"],
            ]);
        });
    });

    it("judges a sidecar under ppt/%73lides/_rels beside its plain twin", async () => {
        await withTempDir(async (dir) => {
            const twoLayouts = relsXml(
                relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml") +
                    relationship("rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"),
            );
            await writeFile(path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"), twoLayouts);
            await writeFile(path.join(dir, "ppt", "%73lides", "_rels", "slide1.xml.rels"), twoLayouts);
            expect((await makeValidator(dir).validateNoDuplicateSlideLayouts()).issues.map((issue) => issue.path)).toEqual([
                "ppt/%73lides/_rels/slide1.xml.rels",
                "ppt/slides/_rels/slide1.xml.rels",
            ]);
        });
    });

    it("judges a master under ppt/slide%4Dasters", async () => {
        await withTempDir(async (dir) => {
            await writeFile(
                path.join(dir, "ppt", "slide%4Dasters", "slideMaster1.xml"),
                presentationXml('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId9"/></p:sldLayoutIdLst>'),
            );
            await writeFile(path.join(dir, "ppt", "slide%4Dasters", "_rels", "slideMaster1.xml.rels"), relsXml(""));
            expect((await makeValidator(dir).validateSlideLayoutIds()).issues.map((issue) => [issue.code, issue.path])).toEqual([
                ["sldlayout-missing-rid", "ppt/slide%4Dasters/slideMaster1.xml"],
            ]);
        });
    });

    it.each(["lenient", "word-valid"] as const)(
        "keeps a %s run invalid when two sidecars under ppt/slides/%5Frels share a notes slide",
        async (profile) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "deck.pptx");
                await fs.writeFile(
                    file,
                    await buildPptx({
                        ...MINIMAL_PPTX_PARTS,
                        "ppt/slides/%5Frels/slide9.xml.rels": notesSidecar(),
                        "ppt/slides/%5Frels/slide10.xml.rels": notesSidecar(),
                        "ppt/notesSlides/notesSlide1.xml":
                            `<?xml version="1.0" encoding="UTF-8"?>\n` +
                            `<p:notes xmlns:p="${PRESENTATIONML_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}">` +
                            `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>`,
                        "[Content_Types].xml": (MINIMAL_PPTX_PARTS["[Content_Types].xml"] ?? "").replace(
                            "</Types>",
                            '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>',
                        ),
                    }),
                );
                const result = await validate(file, { profile });
                expect(result.issues.filter((issue) => issue.code === "notes-shared").map((issue) => issue.message)).toEqual([
                    "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide10, slide9",
                ]);
                expect(result.valid).toBe(false);
            });
        },
    );
});

// #1038: the slide a sidecar belongs to is named by the stem of its stored
// name, `.rels` and `.xml` stripped ASCII case-insensitively, `\` read as `/`.
describe("the slide a notes-shared message names (#1038)", () => {
    const relationship = (id: string, type: string, target: string) =>
        `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;
    const notesSidecar = () => relsXml(relationship("rId1", "notesSlide", "../notesSlides/notesSlide1.xml"));

    it.each([
        ["slide1.xml.RELS", "slide2.xml.RELS"],
        ["slide1.XML.rels", "slide2.Xml.Rels"],
    ])("names the slides of %s and %s by their stems", async (first, second) => {
        await withTempDir(async (dir) => {
            for (const name of [first, second]) {
                await writeFile(path.join(dir, "ppt", "slides", "_rels", name), notesSidecar());
            }
            expect((await makeValidator(dir).validateNotesSlideReferences()).issues[0]?.message).toBe(
                "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2",
            );
        });
    });

    // Two equivalent spellings of one sidecar give one stem twice, so the
    // message names their relationships parts instead (#1041).
    it("names the relationships parts when two equivalent spellings give one stem (#1041)", async () => {
        await withTempDir(async (dir) => {
            await writeFile(path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"), notesSidecar());
            await writeFile(path.join(dir, "ppt", "slides", "%5Frels", "slide1.xml.rels"), notesSidecar());
            expect((await makeValidator(dir).validateNotesSlideReferences()).issues.map((issue) => [issue.code, issue.message])).toEqual([
                [
                    "notes-shared",
                    "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: ppt/slides/%5Frels/slide1.xml.rels, ppt/slides/_rels/slide1.xml.rels",
                ],
                ["notes-shared-detail", "    - ppt/slides/%5Frels/slide1.xml.rels"],
                ["notes-shared-detail", "    - ppt/slides/_rels/slide1.xml.rels"],
            ]);
        });
    });

    // On Windows `\` is a path separator, so the file would land in the plain
    // folder and the row would test nothing (#1042).
    it.skipIf(process.platform === "win32")(
        "names the slide of a sidecar whose stored name spells its folders with backslashes",
        async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "ppt", "slides", "_rels", "slide1.xml.rels"), notesSidecar());
                await writeFile(path.join(dir, "ppt\\slides\\_rels\\slide2.xml.rels"), notesSidecar());
                expect((await makeValidator(dir).validateNotesSlideReferences()).issues[0]?.message).toBe(
                    "Notes slide 'ppt/notesSlides/notesSlide1.xml' is referenced by multiple slides: slide1, slide2",
                );
            });
        },
    );
});

// #1018: a master's relationships part is found by OPC part-name
// equivalence, as the slide checks find a slide's, so a sidecar stored under
// an escaped spelling is read, not reported missing.
describe("a slide master's relationships part stored under an equivalent spelling (#1018)", () => {
    const relationship = (id: string, type: string, target: string) =>
        `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;

    /** A master in `ppt/<folder>/` naming `rId`, its sidecar stored as `slideMaster%31.xml.rels`; the presentation relates any master outside `ppt/slideMasters/`. */
    async function storeMaster(dir: string, folder: string, rId: string): Promise<void> {
        if (folder !== "slideMasters") {
            await writeFile(path.join(dir, "_rels", ".rels"), relsXml(relationship("rId1", "officeDocument", "ppt/presentation.xml")));
            await writeFile(path.join(dir, "ppt", "presentation.xml"), presentationXml(""));
            await writeFile(
                path.join(dir, "ppt", "_rels", "presentation.xml.rels"),
                relsXml(relationship("rId1", "slideMaster", `${folder}/slideMaster1.xml`)),
            );
        }
        await writeFile(
            path.join(dir, "ppt", folder, "slideMaster1.xml"),
            presentationXml(`<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="${rId}"/></p:sldLayoutIdLst>`),
        );
        await writeFile(
            path.join(dir, "ppt", folder, "_rels", "slideMaster%31.xml.rels"),
            relsXml(relationship("rId9", "slideLayout", "../slideLayouts/slideLayout1.xml")),
        );
    }

    const findings = async (dir: string) =>
        (await makeValidator(dir).validateSlideLayoutIds()).issues.map((issue) => [issue.code, issue.path]);

    it.each(["masters", "slideMasters"])("reads the escaped sidecar of a master in ppt/%s", async (folder) => {
        await withTempDir(async (dir) => {
            await storeMaster(dir, folder, "rId9");
            expect(await findings(dir)).toEqual([]);
        });
    });

    it.each(["masters", "slideMasters"])("flags a sldLayoutId the escaped sidecar of a master in ppt/%s lacks", async (folder) => {
        await withTempDir(async (dir) => {
            await storeMaster(dir, folder, "rId8");
            expect(await findings(dir)).toEqual([["sldlayout-missing-rid", `ppt/${folder}/slideMaster1.xml`]]);
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

// #1016: Word sessions 65 and 67 opened DOCX packages only, so a PowerPoint
// part's AlternateContent character data is reported without them.
describe("AlternateContent character data in a slide cites no Word session (#1016)", () => {
    it.each(["lenient", "strict"] as const)("reports it uncited under %s", async (profile) => {
        await withTempDir(async (dir) => {
            const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            const parts: Record<string, string> = { ...MINIMAL_PPTX_PARTS };
            const slide = "ppt/slides/slide1.xml";
            expect(parts[slide]).toContain("<p:cSld");
            parts[slide] = parts[slide]!.replace(
                /(<p:cSld[^>]*>)/,
                `$1<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="p14" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>Stray<mc:Fallback/></mc:AlternateContent>`,
            );
            const file = path.join(dir, "stray.pptx");
            await fs.writeFile(file, await buildPptx(parts));
            const result = await validate(file, { profile });
            expect(result.valid).toBe(false);
            expect(result.issues.filter((issue) => issue.code === "mce-malformed").map((issue) => [issue.path, issue.message])).toEqual([
                [slide, "AlternateContent holds character data directly inside it."],
            ]);
        });
    });
});
