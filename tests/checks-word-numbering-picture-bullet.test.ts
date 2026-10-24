/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document } from "./support/word-rule-test-xml";
import { wordNumberingPictureBulletCheck } from "../src/scripts/office/checks/word-numbering-picture-bullet";

describe("word-numbering-picture-bullet", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/numbering.xml": document(`<w:numPicBullet w:numPicBulletId="1"/><w:lvlPicBulletId w:val="2"/>`),
        });
        expect(await wordNumberingPictureBulletCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-numbering-picture-bullet",
                path: "word/numbering.xml",
                message:
                    "A numbering level references picture bullet 2, and this part has no w:numPicBullet with that id. Microsoft Word offers to recover the file. Remove the w:lvlPicBulletId, or add the matching picture bullet.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/numbering.xml": document(
                `<w:numPicBullet w:numPicBulletId="1"/><w:numPicBullet/><w:lvlPicBulletId w:val="1"/><w:lvlPicBulletId/>`,
            ),
            "word/document.xml": document(`<w:lvlPicBulletId w:val="2"/>`),
        });
        expect(await wordNumberingPictureBulletCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordNumberingPictureBulletCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordNumberingPictureBulletCheck.run(
                ValidationContext.fromEntries({ "word/numbering.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
