/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * #1101: each Word rule that judges an element's relationship id reads it in
 * both relationships namespaces, and its message names the attribute as the
 * part spells it: the qualified name it was read under, as #1090 does for
 * `rels-id-missing`. Each specimen is the minimal package plus one refused
 * id, spelled three ways: `r` bound to the Transitional namespace, `rs` bound
 * to the Strict one, and `r` rebound to the Strict one on the element, as a
 * Word-saved Strict package binds it. The rows run `validate()` under the two
 * profiles that run the Word rules.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const O = "urn:schemas-microsoft-com:office:office";
const WNE = "http://schemas.microsoft.com/office/word/2006/wordml";

const MAIN = MINIMAL_PARTS["word/document.xml"]!.replace(
    "<w:document ",
    `<w:document xmlns:r="${R}" xmlns:rs="${RS}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wne="${WNE}" `,
);
/** The minimal package with `body` before the body `w:sectPr` and `section` as its first children. */
const docx = (body: string, section = "") =>
    buildDocx({ ...MINIMAL_PARTS, "word/document.xml": MAIN.replace("<w:sectPr>", `${body}<w:sectPr>${section}`) });

const drawing = (graphicData: string) =>
    '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/>' +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Object 1"/><a:graphic>${graphicData}</a:graphic></wp:inline></w:drawing></w:r></w:p>`;

/** `declaration` goes on the judged element; `attribute` is the qualified name its id is spelled with. */
interface Shape {
    readonly code: string;
    readonly build: (declaration: string, attribute: string) => Promise<Buffer>;
    readonly label: (attribute: string) => string;
    readonly advice: (attribute: string) => string;
    readonly localName?: string;
}

const SHAPES: Readonly<Record<string, Shape>> = {
    "a w:hyperlink naming no relationship": {
        code: "word-hyperlink-relationship",
        build: (declaration, attribute) =>
            docx(`<w:p><w:hyperlink${declaration} ${attribute}="rIdX"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`),
        label: (attribute) => `<w:hyperlink> ${attribute}="rIdX" does not reference`,
        advice: (attribute) => `with TargetMode External, or remove the ${attribute}.`,
    },
    "a w:headerReference naming no relationship": {
        code: "word-header-footer-relationship",
        build: (declaration, attribute) => docx("", `<w:headerReference${declaration} w:type="default" ${attribute}="rIdX"/>`),
        label: (attribute) => `<w:headerReference> ${attribute}="rIdX" names no Relationship`,
        advice: (attribute) => `Point the ${attribute} at such a Relationship, or add one`,
    },
    "a w:headerReference naming the styles relationship": {
        code: "word-header-footer-relationship",
        build: (declaration, attribute) => docx("", `<w:headerReference${declaration} w:type="default" ${attribute}="rId1"/>`),
        label: (attribute) => `<w:headerReference> ${attribute}="rId1" names Relationship 'rId1' of Type`,
        advice: (attribute) => `Point the ${attribute} at such a Relationship, or retype`,
    },
    "a wne:toolbarData naming no relationship": {
        code: "word-attached-toolbar-relationship",
        build: (declaration, attribute) => docx(`<wne:toolbarData${declaration} ${attribute}="rIdX"/>`),
        label: (attribute) => `<wne:toolbarData> ${attribute}="rIdX" does not reference`,
        advice: (attribute) => `or point the ${attribute} at one.`,
    },
    "a c:chart naming no relationship": {
        code: "word-part-reference-relationship",
        build: (declaration, attribute) =>
            docx(drawing(`<a:graphicData uri="${C}"><c:chart xmlns:c="${C}"${declaration} ${attribute}="rIdX"/></a:graphicData>`)),
        label: (attribute) => `<c:chart> ${attribute}="rIdX" does not reference`,
        advice: (attribute) => `internal to the package, or point the ${attribute} at one.`,
    },
    "an o:OLEObject naming no relationship": {
        code: "word-part-reference-relationship",
        build: (declaration, attribute) =>
            docx(
                `<w:p><w:r><w:object><o:OLEObject xmlns:o="${O}"${declaration} Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" ` +
                    `DrawAspect="Icon" ObjectID="_1" ${attribute}="rIdX"/></w:object></w:r></w:p>`,
            ),
        label: (attribute) => `<o:OLEObject> ${attribute}="rIdX" does not reference`,
        advice: (attribute) => `/oleObject), or point the ${attribute} at one.`,
    },
    "a dgm:relIds data id naming no relationship": {
        code: "word-part-reference-relationship",
        localName: "dm",
        build: (declaration, attribute) =>
            docx(drawing(`<a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}"${declaration} ${attribute}="rIdX"/></a:graphicData>`)),
        label: (attribute) => `<dgm:relIds> ${attribute}="rIdX" does not reference`,
        advice: (attribute) => `internal to the package, or point the ${attribute} at one.`,
    },
};

/** Each spelling's declaration on the judged element and the prefix its id is written with. */
const SPELLINGS = {
    "r bound to the Transitional namespace": { declaration: "", prefix: "r" },
    "rs bound to the Strict namespace": { declaration: "", prefix: "rs" },
    "r rebound to the Strict namespace": { declaration: ` xmlns:r="${RS}"`, prefix: "r" },
} as const;

/** The messages the shape's rule reports in a full run. */
const messages = async (shape: string, spelling: keyof typeof SPELLINGS, profile: "word-valid" | "strict") =>
    withTempDir(async (dir) => {
        const { code, build, localName = "id" } = SHAPES[shape]!;
        const { declaration, prefix } = SPELLINGS[spelling];
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, await build(declaration, `${prefix}:${localName}`));
        const result = await validate(file, { profile });
        return result.issues.filter((issue) => issue.code === code).map((issue) => issue.message);
    });

const ROWS = Object.keys(SHAPES).flatMap((shape) =>
    (Object.keys(SPELLINGS) as (keyof typeof SPELLINGS)[]).flatMap((spelling) =>
        (["word-valid", "strict"] as const).map((profile) => [shape, spelling, profile] as const),
    ),
);

describe("a Word rule names the refused id attribute as the part spells it (#1101)", () => {
    it.each(ROWS)("%s, %s, under %s", async (shape, spelling, profile) => {
        const { label, advice, localName = "id" } = SHAPES[shape]!;
        const attribute = `${SPELLINGS[spelling].prefix}:${localName}`;
        const found = await messages(shape, spelling, profile);
        expect(found).toHaveLength(1);
        expect(found[0]).toContain(label(attribute));
        expect(found[0]).toContain(advice(attribute));
        // No other spelling of the attribute is left in the prose.
        expect(found[0]!.split(attribute)).toHaveLength(3);
        if (attribute !== `r:${localName}`) expect(found[0]).not.toContain(`r:${localName}`);
    });
});

/**
 * #1109: the messages #1101 left alone. The header/footer rule names an empty
 * id as the part spells it, and keeps the conventional `r:id` for a missing
 * one. The chart's dual-namespace message names each value with the attribute
 * that holds it, whatever the prefixes are bound to. And where an element
 * carries an id in both namespaces, the label names the attribute the rule
 * judged (the first non-empty id in the order it reads them, Transitional
 * first), not the first one in document order.
 */
describe("the remaining relationship-id messages name the attribute as the part spells it (#1109)", () => {
    const PROFILES = ["word-valid", "strict"] as const;
    /** The messages `code` reports in a full run of `bytes`. */
    const reported = async (bytes: Buffer, code: string, profile: "word-valid" | "strict") =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, bytes);
            const result = await validate(file, { profile });
            return result.issues.filter((issue) => issue.code === code).map((issue) => issue.message);
        });

    const EMPTY_ROWS = (Object.keys(SPELLINGS) as (keyof typeof SPELLINGS)[]).flatMap((spelling) =>
        PROFILES.map((profile) => [spelling, profile] as const),
    );
    it.each(EMPTY_ROWS)("an empty w:headerReference id, %s, under %s", async (spelling, profile) => {
        const { declaration, prefix } = SPELLINGS[spelling];
        const attribute = `${prefix}:id`;
        const found = await reported(
            await docx("", `<w:headerReference${declaration} w:type="default" ${attribute}=""/>`),
            "word-header-footer-relationship",
            profile,
        );
        expect(found).toHaveLength(1);
        expect(found[0]).toContain(`<w:headerReference> has an empty ${attribute}; `);
        expect(found[0]).toContain(`Set a non-empty ${attribute} naming such a Relationship`);
        expect(found[0]!.split(attribute)).toHaveLength(3);
        if (attribute !== "r:id") expect(found[0]).not.toContain("r:id");
    });

    it.each(PROFILES)("an empty id in both namespaces, the Strict one first, is named as the one read first under %s", async (profile) => {
        const found = await reported(
            await docx("", '<w:headerReference w:type="default" rs:id="" r:id=""/>'),
            "word-header-footer-relationship",
            profile,
        );
        expect(found).toHaveLength(1);
        expect(found[0]).toContain("<w:headerReference> has an empty r:id; ");
        expect(found[0]).not.toContain("rs:id");
    });

    it.each(PROFILES)("a missing w:headerReference id keeps the conventional r:id under %s", async (profile) => {
        const found = await reported(await docx("", '<w:headerReference w:type="default"/>'), "word-header-footer-relationship", profile);
        expect(found).toHaveLength(1);
        expect(found[0]).toContain("<w:headerReference> has a missing r:id; ");
        expect(found[0]).toContain("Set a non-empty r:id naming such a Relationship");
    });

    /** `declaration` and `ids` on a `c:chart` carrying an id in both namespaces; `named` is what the message must say. */
    const DUAL = {
        "r Transitional, rs Strict": { declaration: "", ids: 'r:id="rIdA" rs:id="rIdB"', named: 'r:id="rIdA" rs:id="rIdB"' },
        "r Strict, rs Transitional": {
            declaration: ` xmlns:r="${RS}" xmlns:rs="${R}"`,
            ids: 'r:id="rIdB" rs:id="rIdA"',
            named: 'rs:id="rIdA" r:id="rIdB"',
        },
        "foo Transitional, bar Strict": {
            declaration: ` xmlns:foo="${R}" xmlns:bar="${RS}"`,
            ids: 'foo:id="rIdA" bar:id="rIdB"',
            named: 'foo:id="rIdA" bar:id="rIdB"',
        },
    } as const;
    const DUAL_ROWS = (Object.keys(DUAL) as (keyof typeof DUAL)[]).flatMap((binding) =>
        PROFILES.map((profile) => [binding, profile] as const),
    );
    it.each(DUAL_ROWS)("a c:chart with an id in both namespaces, %s, names each value's attribute under %s", async (binding, profile) => {
        const { declaration, ids, named } = DUAL[binding];
        const found = await reported(
            await docx(drawing(`<a:graphicData uri="${C}"><c:chart xmlns:c="${C}"${declaration} ${ids}/></a:graphicData>`)),
            "word-part-reference-relationship",
            profile,
        );
        expect(found).toHaveLength(1);
        expect(found[0]).toContain(`<c:chart> carries an id attribute in both relationships namespaces (${named}). `);
    });

    /** An element carrying a dangling id in both namespaces, the Strict one first in document order. */
    const BOTH_IDS = {
        "a w:hyperlink": {
            code: "word-hyperlink-relationship",
            build: (declaration: string, ids: string) =>
                docx(`<w:p><w:hyperlink${declaration} ${ids}><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`),
            label: (attribute: string) => `<w:hyperlink> ${attribute}="rIdTrans" does not reference`,
        },
        "a w:headerReference": {
            code: "word-header-footer-relationship",
            build: (declaration: string, ids: string) => docx("", `<w:headerReference${declaration} w:type="default" ${ids}/>`),
            label: (attribute: string) => `<w:headerReference> ${attribute}="rIdTrans" names no Relationship`,
        },
        "an o:OLEObject": {
            code: "word-part-reference-relationship",
            build: (declaration: string, ids: string) =>
                docx(
                    `<w:p><w:r><w:object><o:OLEObject xmlns:o="${O}"${declaration} Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" ` +
                        `DrawAspect="Icon" ObjectID="_1" ${ids}/></w:object></w:r></w:p>`,
                ),
            label: (attribute: string) => `<o:OLEObject> ${attribute}="rIdTrans" does not reference`,
        },
    } as const;
    /** Each binding's declaration, its ids with the Strict one first, and the name the Transitional id is spelled with. */
    const BOTH_BINDINGS = {
        "r Transitional, rs Strict": { declaration: "", ids: 'rs:id="rIdStrict" r:id="rIdTrans"', attribute: "r:id" },
        "r Strict, rs Transitional": {
            declaration: ` xmlns:r="${RS}" xmlns:rs="${R}"`,
            ids: 'r:id="rIdStrict" rs:id="rIdTrans"',
            attribute: "rs:id",
        },
    } as const;
    const BOTH_ROWS = (Object.keys(BOTH_IDS) as (keyof typeof BOTH_IDS)[]).flatMap((shape) =>
        (Object.keys(BOTH_BINDINGS) as (keyof typeof BOTH_BINDINGS)[]).flatMap((binding) =>
            PROFILES.map((profile) => [shape, binding, profile] as const),
        ),
    );
    it.each(BOTH_ROWS)(
        "%s with an id in both namespaces, %s, names the judged Transitional id under %s",
        async (shape, binding, profile) => {
            const { code, build, label } = BOTH_IDS[shape];
            const { declaration, ids, attribute } = BOTH_BINDINGS[binding];
            const found = await reported(await build(declaration, ids), code, profile);
            expect(found).toHaveLength(1);
            expect(found[0]).toContain(label(attribute));
        },
    );
});
