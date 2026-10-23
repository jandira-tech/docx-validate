/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Shading colour value-space check (Task 8 family 8c; lane issues #210, #215,
 * #220, #221, #225, #229, #238, #242).
 *
 * `wml.xsd` types every colour attribute of `CT_Shd`: `w:color` and `w:fill`
 * are `ST_HexColor` (the union of the `ST_HexColorAuto` enumeration `"auto"`
 * and `s:ST_HexColorRGB`, a 3-byte `hexBinary`, i.e. exactly six hexadecimal
 * digits), `w:themeColor`/`w:themeFill` are the `ST_ThemeColor` enumeration,
 * and `w:themeTint`/`w:themeShade`/`w:themeFillTint`/`w:themeFillShade` are
 * `ST_UcharHexNumber` (a 1-byte `hexBinary`, i.e. exactly two hexadecimal
 * digits). The jubarte 0.13.0 lane writes `<w:shd w:val="clear" w:fill=""/>`;
 * the empty string is in none of those value spaces, so strict schema
 * validation rejects every occurrence while Word opens the file regardless.
 * This check reports each value outside its declared space as a warning.
 *
 * `w:shd` may appear wherever paragraph, run, table, and cell properties go:
 * story parts (document, headers, footers, footnotes, endnotes, comments) and
 * `styles.xml`. The `w:val` pattern is not colour-typed and stays out of scope.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

/** `ST_HexColorRGB`: a 3-byte hexBinary, lexically six hexadecimal digits. */
const HEX_COLOR_RGB = /^[0-9A-Fa-f]{6}$/;
/** `ST_UcharHexNumber`: a 1-byte hexBinary, lexically two hexadecimal digits. */
const UCHAR_HEX = /^[0-9A-Fa-f]{2}$/;
/** `ST_HexColorAuto`: the enumeration's single member. */
const HEX_COLOR_AUTO = "auto";
// xsd:hexBinary fixes its whiteSpace facet to collapse, so the schema
// validates hex colours after trimming XML whitespace; the xsd:string
// enumerations ("auto", theme names) keep theirs.
const collapseHex = (value: string): string => value.replace(/^[ \t\n\r]+|[ \t\n\r]+$/g, "");
/** `ST_ThemeColor` enumeration, wml.xsd lines 3449-3469. */
const THEME_COLORS: ReadonlySet<string> = new Set([
    "dark1",
    "light1",
    "dark2",
    "light2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followedHyperlink",
    "none",
    "background1",
    "text1",
    "background2",
    "text2",
]);

type ColourSpace = {
    /** The simple type the schema declares for the attribute. */
    readonly type: string;
    /** The permitted values, as guidance in the finding message. */
    readonly allowed: string;
    readonly isValid: (value: string) => boolean;
};

/** Colour-typed attributes of `CT_Shd`, by their `ST_` value space. */
const COLOUR_ATTRIBUTES: ReadonlyMap<string, ColourSpace> = new Map(
    Object.entries({
        color: {
            type: "ST_HexColor",
            allowed: `"auto" or exactly six hexadecimal digits (for example "DDDDDD")`,
            isValid: (value): boolean => value === HEX_COLOR_AUTO || HEX_COLOR_RGB.test(collapseHex(value)),
        },
        fill: {
            type: "ST_HexColor",
            allowed: `"auto" or exactly six hexadecimal digits (for example "DDDDDD")`,
            isValid: (value): boolean => value === HEX_COLOR_AUTO || HEX_COLOR_RGB.test(collapseHex(value)),
        },
        themeColor: {
            type: "ST_ThemeColor",
            allowed: `a theme colour name such as "accent1", "dark1", or "background1"`,
            isValid: (value): boolean => THEME_COLORS.has(value),
        },
        themeFill: {
            type: "ST_ThemeColor",
            allowed: `a theme colour name such as "accent1", "dark1", or "background1"`,
            isValid: (value): boolean => THEME_COLORS.has(value),
        },
        themeTint: {
            type: "ST_UcharHexNumber",
            allowed: `exactly two hexadecimal digits (for example "66")`,
            isValid: (value): boolean => UCHAR_HEX.test(collapseHex(value)),
        },
        themeShade: {
            type: "ST_UcharHexNumber",
            allowed: `exactly two hexadecimal digits (for example "66")`,
            isValid: (value): boolean => UCHAR_HEX.test(collapseHex(value)),
        },
        themeFillTint: {
            type: "ST_UcharHexNumber",
            allowed: `exactly two hexadecimal digits (for example "66")`,
            isValid: (value): boolean => UCHAR_HEX.test(collapseHex(value)),
        },
        themeFillShade: {
            type: "ST_UcharHexNumber",
            allowed: `exactly two hexadecimal digits (for example "66")`,
            isValid: (value): boolean => UCHAR_HEX.test(collapseHex(value)),
        },
    } satisfies Record<string, ColourSpace>),
);

/** Story parts plus styles.xml, where `CT_Shd` may legally appear. */
const shadingParts = (storyParts: readonly string[], partNames: readonly string[]): readonly string[] => [
    ...storyParts,
    ...partNames.filter((name) => /(?:^|\/)styles\.xml$/i.test(name)),
];

export const wordShadingColorCheck: CheckDefinition = {
    id: "word-shading-color",
    description: "Every w:shd colour attribute in the Word story parts and styles.xml holds a value its schema type allows.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const part of shadingParts(ctx.storyParts(), ctx.partNames)) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) continue;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const shd of getElementsByTagNameNSAll(parsed.document, ns, "shd")) {
                    issues.push(...shadingIssues(shd, part));
                }
            }
        }
        return issues;
    },
};

/** One issue per colour attribute whose value is outside its declared space. */
const shadingIssues = (shd: Element, part: string): ValidationIssue[] => {
    const prefix = shd.prefix ? `${shd.prefix}:` : "";
    const issues: ValidationIssue[] = [];
    for (let i = 0; i < shd.attributes.length; i += 1) {
        const attr = shd.attributes[i]!;
        // CT_Shd declares its attributes qualified in the element's own
        // namespace, so only same-namespace attributes are the declared ones.
        if (attr.namespaceURI !== shd.namespaceURI) continue;
        const space = COLOUR_ATTRIBUTES.get(attr.localName ?? "");
        if (space === undefined || space.isValid(attr.value ?? "")) continue;
        issues.push({
            severity: "warning",
            code: "word-shading-color",
            path: part,
            message:
                `${prefix}shd carries ${prefix}${attr.localName}="${attr.value}", which the schema does not allow: ` +
                `${space.type} accepts ${space.allowed}. Omit the optional attribute when no colour is set.`,
        });
    }
    return issues;
};
