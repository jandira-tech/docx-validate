/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The default registry of named checks. Registration order is dependency
 * order: a check appears after every check it requires.
 */

import { type CheckDefinition, CheckRegistry } from "../../../lib/check-registry";
import { wordCellMarginOrderCheck } from "./cell-margin-order";
import { wordGridColumnAttributesCheck } from "./grid-column-attributes";
import { mathRunDeletionCheck } from "./math-run-deletion";
import { mathRunInsertionCheck } from "./math-run-insertion";
import { zipBytesAfterEocdCheck, zipEocdCommentOverrunCheck } from "./package";
import { pageMarginAttributesCheck } from "./page-margin-attributes";
import { wordShadingColorCheck } from "./shading-color";
import { wordTablePropertiesRequiredCheck } from "./table-properties-required";
import { wordRowWithoutCellCheck } from "./tables";
import { textboxContentNamespaceCheck } from "./textbox-content-namespace";
import { wordCustomPropertyFiletimeEmptyCheck } from "./word-custom-properties";
import { wordDrawingExtentMissingCheck } from "./word-drawing";
import { wordPictureNvpicprMissingCheck, wordPictureSpprMissingCheck } from "./word-picture";
import { xmlSyntaxCheck } from "./xml";

/**
 * Word rules that `DOCXSchemaValidator` runs through the registry under the
 * `word-valid` and `strict` profiles. Lenient callers never see them, so
 * adding a rule here cannot change legacy output. Each rule skips parts that
 * do not parse; XML syntax is reported by its own check.
 */
export const WORD_RULE_CHECKS: readonly CheckDefinition[] = [
    mathRunInsertionCheck,
    mathRunDeletionCheck,
    wordShadingColorCheck,
    wordGridColumnAttributesCheck,
    wordCellMarginOrderCheck,
    wordTablePropertiesRequiredCheck,
    pageMarginAttributesCheck,
    textboxContentNamespaceCheck,
    wordCustomPropertyFiletimeEmptyCheck,
    wordDrawingExtentMissingCheck,
    wordPictureNvpicprMissingCheck,
    wordPictureSpprMissingCheck,
];

export function createDefaultCheckRegistry(): CheckRegistry {
    return new CheckRegistry([
        zipBytesAfterEocdCheck,
        zipEocdCommentOverrunCheck,
        xmlSyntaxCheck,
        wordRowWithoutCellCheck,
        ...WORD_RULE_CHECKS,
    ]);
}
