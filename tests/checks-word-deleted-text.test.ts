/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, WS } from "./support/word-rule-test-xml";
import { wordTextboxDeletedTextCheck, wordMoveFromDeletedTextCheck } from "../src/scripts/office/checks/word-deleted-text";

describe("word-textbox-deleted-text", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:del><w:txbxContent><w:p><w:r><w:delText>x</w:delText></w:r></w:p></w:txbxContent></w:del>"),
        });
        expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-textbox-deleted-text",
                path: "word/document.xml",
                message:
                    "A text box contains deleted text. Microsoft Word offers to recover the file. Remove the w:delText or w:delInstrText from the text box.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                "<w:p><w:r><w:delText>x</w:delText><w:delInstrText>y</w:delInstrText></w:r></w:p><w:txbxContent><w:del><w:p><w:r><w:delText>x</w:delText><w:delInstrText>y</w:delInstrText></w:r></w:p></w:del></w:txbxContent>",
            ),
        });
        expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordTextboxDeletedTextCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordTextboxDeletedTextCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-move-from-deleted-text", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:p><w:moveFrom><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:p>"),
        });
        expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-move-from-deleted-text",
                path: "word/document.xml",
                message:
                    "A move source contains deleted text. Microsoft Word offers to recover the file. Use w:t inside w:moveFrom, not w:delText.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document("<w:p><w:moveFrom><w:r><w:t>x</w:t></w:r></w:moveFrom><w:r><w:delText>y</w:delText></w:r></w:p>"),
        });
        expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordMoveFromDeletedTextCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordMoveFromDeletedTextCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it("distinguishes deletion within a text box from an outside deletion and ordinary body text", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(
            `<w:txbxContent><w:p><w:r><w:delInstrText>x</w:delInstrText></w:r></w:p></w:txbxContent><w:p><w:r><w:delText>body</w:delText></w:r></w:p>`,
        ),
    });
    expect(await wordTextboxDeletedTextCheck.run(ctx)).toHaveLength(1);
    expect(await wordMoveFromDeletedTextCheck.run(ctx)).toEqual([]);
});
it("retains the historical tolerance for fallback text boxes", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(
            `<mc:AlternateContent><mc:Choice Requires="wps" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:p/></mc:Choice><mc:Fallback><w:txbxContent><w:delText>x</w:delText></w:txbxContent></mc:Fallback></mc:AlternateContent>`,
        ),
    });
    expect(await wordTextboxDeletedTextCheck.run(ctx)).toEqual([]);
});
/*
 * The checks walk each part once and carry the enclosing text box, deletion,
 * mc:Fallback and move source down the tree (bot PR #199; #421, #200 and #202
 * queried the containers first instead). These shapes pin the verdicts of the
 * per-node ancestor walks they replace, including the mixed-namespace and
 * nested-container cases reviewers raised on #200 and #202.
 */
describe("deleted-text checks across namespaces and nested containers", () => {
    const strict = `xmlns:s="${WS}"`;
    const count = async (check: typeof wordTextboxDeletedTextCheck, body: string): Promise<number> =>
        (await check.run(ValidationContext.fromEntries({ "word/document.xml": document(body) }))).length;

    it("reports deleted text whose namespace differs from its text box's", async () => {
        expect(
            await count(
                wordTextboxDeletedTextCheck,
                `<w:txbxContent><w:p><w:r><s:delText ${strict}>x</s:delText></w:r></w:p></w:txbxContent>`,
            ),
        ).toBe(1);
        expect(
            await count(
                wordTextboxDeletedTextCheck,
                `<s:txbxContent ${strict}><w:p><w:r><w:delInstrText>x</w:delInstrText></w:r></w:p></s:txbxContent>`,
            ),
        ).toBe(1);
        expect(
            await count(
                wordTextboxDeletedTextCheck,
                `<s:txbxContent ${strict}><s:del><w:r><w:delText>x</w:delText></w:r></s:del></s:txbxContent>`,
            ),
        ).toBe(0);
    });

    it("judges a nested text box by the deletion inside the innermost box", async () => {
        const inner = (content: string): string =>
            `<w:txbxContent><w:p><w:r><w:pict><w:txbxContent>${content}</w:txbxContent></w:pict></w:r></w:p></w:txbxContent>`;
        expect(await count(wordTextboxDeletedTextCheck, inner(`<w:p><w:r><w:delText>x</w:delText></w:r></w:p>`))).toBe(1);
        expect(await count(wordTextboxDeletedTextCheck, `<w:del>${inner(`<w:p><w:r><w:delText>x</w:delText></w:r></w:p>`)}</w:del>`)).toBe(
            1,
        );
        expect(await count(wordTextboxDeletedTextCheck, inner(`<w:del><w:r><w:delText>x</w:delText></w:r></w:del>`))).toBe(0);
        expect(
            await count(
                wordTextboxDeletedTextCheck,
                `<w:txbxContent><w:del><w:r><w:pict><w:txbxContent><w:p><w:r><w:delText>x</w:delText></w:r></w:p></w:txbxContent></w:pict></w:r></w:del></w:txbxContent>`,
            ),
        ).toBe(1);
    });

    it("reports each offending element once", async () => {
        const body =
            `<w:txbxContent><w:p><w:r><w:delText>a</w:delText><w:delInstrText>b</w:delInstrText><w:delText>c</w:delText></w:r></w:p>` +
            `<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:delText>d</w:delText></w:r></w:p></w:txbxContent></w:pict></w:r></w:p></w:txbxContent>`;
        expect(await count(wordTextboxDeletedTextCheck, body)).toBe(4);
    });

    it("reports move-source deleted text across namespaces and nesting, but not w:delInstrText", async () => {
        expect(
            await count(wordMoveFromDeletedTextCheck, `<w:p><w:moveFrom><w:r><s:delText ${strict}>x</s:delText></w:r></w:moveFrom></w:p>`),
        ).toBe(1);
        expect(
            await count(wordMoveFromDeletedTextCheck, `<w:p><s:moveFrom ${strict}><w:r><w:delText>x</w:delText></w:r></s:moveFrom></w:p>`),
        ).toBe(1);
        expect(
            await count(
                wordMoveFromDeletedTextCheck,
                `<w:p><w:moveFrom><w:moveFrom><w:r><w:delText>x</w:delText><w:delText>y</w:delText></w:r></w:moveFrom></w:moveFrom></w:p>`,
            ),
        ).toBe(2);
        expect(
            await count(wordMoveFromDeletedTextCheck, `<w:p><w:moveFrom><w:r><w:delInstrText>x</w:delInstrText></w:r></w:moveFrom></w:p>`),
        ).toBe(0);
        expect(await count(wordMoveFromDeletedTextCheck, `<w:p><w:moveTo><w:r><w:delText>x</w:delText></w:r></w:moveTo></w:p>`)).toBe(0);
    });

    it("reads Strict documents", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:txbxContent><w:p><w:moveFrom><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:p></w:txbxContent>`,
                WS,
            ),
        });
        expect(await wordTextboxDeletedTextCheck.run(ctx)).toHaveLength(1);
        expect(await wordMoveFromDeletedTextCheck.run(ctx)).toHaveLength(1);
    });
});
