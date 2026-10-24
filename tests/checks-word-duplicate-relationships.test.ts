/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, rels, relation } from "./support/word-rule-test-xml";
import {
    wordDuplicateThemeRelationshipCheck,
    wordDuplicateFontTableRelationshipCheck,
    wordDuplicateWebSettingsRelationshipCheck,
    wordDuplicateSettingsRelationshipCheck,
} from "../src/scripts/office/checks/word-duplicate-relationships";

describe("word-duplicate-theme-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "theme") + relation("rId2", "theme")),
        });
        expect(await wordDuplicateThemeRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-theme-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 theme relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one theme relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "theme") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateThemeRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateThemeRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateThemeRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-font-table-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "fontTable") + relation("rId2", "fontTable")),
        });
        expect(await wordDuplicateFontTableRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-font-table-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 font table relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one font table relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "fontTable") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateFontTableRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateFontTableRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateFontTableRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-web-settings-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "webSettings") + relation("rId2", "webSettings")),
        });
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-web-settings-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 web settings relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one web settings relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "webSettings") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateWebSettingsRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-settings-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "settings") + relation("rId2", "settings")),
        });
        expect(await wordDuplicateSettingsRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-settings-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 settings relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one settings relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "settings") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateSettingsRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateSettingsRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateSettingsRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
