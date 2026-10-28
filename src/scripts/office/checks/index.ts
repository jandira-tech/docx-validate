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
import { zipBytesAfterEocdCheck, zipBytesBeforeArchiveCheck, zipEocdCommentOverrunCheck } from "./package";
import { opcEquivalentPartNamesCheck } from "./opc-equivalent-part-names";
import { opcRelationshipsStructureCheck } from "./opc-relationships-structure";
import { opcRelationshipsTargetModeCheck } from "./opc-relationships-target-mode";
import { opcPartNameWordRefusedCheck } from "./opc-part-name-word-refused";
import { opcItemNameRefusedCheck } from "./opc-item-name-refused";
import { ctOverridePartnameWordRefusedCheck } from "./ct-override-partname-word-refused";
import { opcPartNameEscapedSlashCheck } from "./opc-part-name-escaped-slash";
import { opcPartNameEscapedUnreservedCheck } from "./opc-part-name-escaped-unreserved";
import { pageMarginAttributesCheck } from "./page-margin-attributes";
import { wordShadingColorCheck } from "./shading-color";
import { wordTablePropertiesRequiredCheck } from "./table-properties-required";
import { wordRowWithoutCellCheck } from "./tables";
import { textboxContentNamespaceCheck } from "./textbox-content-namespace";
import { wordCustomPropertyFiletimeEmptyCheck } from "./word-custom-properties";
import { wordDrawingExtentMissingCheck } from "./word-drawing";
import { wordDrawingExtentMisplacedCheck } from "./word-drawing-extent-position";
import { wordPictureNvpicprMissingCheck, wordPictureSpprMissingCheck } from "./word-picture";
import { wordParagraphPropertiesMisplacedCheck } from "./word-paragraph-properties";
import { wordRunMisplacedCheck, wordRunPropertiesMisplacedCheck, wordUnderlineMisplacedCheck } from "./word-run-placement";
import { wordSectionPropertiesMisplacedCheck } from "./word-section-properties";
import { wordTablePropertiesMisplacedCheck } from "./word-table-properties";
import { wordLinkWithoutNamespaceCheck, wordUnknownElementCheck } from "./word-unknown-elements";
import { wordMarkupCompatibilityCheck } from "./word-markup-compatibility";
import { wordDocumentRootCheck, wordMainDocumentPartCheck } from "./word-document-root";
import { ctDuplicateDeclarationCheck, wordContentTypeInvalidCheck } from "./word-content-types";
import { ctOverridePartnameInvalidCheck } from "./ct-override-partname";
import { wordPartContentTypeGenericCheck } from "./word-part-content-type-generic";
import { wordDrawingScalarWhitespaceCheck, wordDrawingArcRadiusCheck, wordDrawingExtentNegativeCheck } from "./word-drawing-values";
import { wordEmptyDrawingCheck, wordAnchorMissingAttributeCheck, wordPositionOffsetAndAlignCheck } from "./word-drawing-anchors";
import { wordCellEndsWithTableCheck } from "./word-cell-ending";
import { wordTextboxDeletedTextCheck, wordMoveFromDeletedTextCheck } from "./word-deleted-text";
import { wordTextContainsParagraphCheck } from "./word-text-paragraph";
import { wordHyperlinkRelationshipCheck } from "./word-hyperlink";
import {
    wordDuplicateThemeRelationshipCheck,
    wordDuplicateFontTableRelationshipCheck,
    wordDuplicateWebSettingsRelationshipCheck,
    wordDuplicateSettingsRelationshipCheck,
    wordDuplicateStylesRelationshipCheck,
    wordDuplicateNumberingRelationshipCheck,
    wordDuplicateEndnotesRelationshipCheck,
    wordDuplicateFootnotesRelationshipCheck,
} from "./word-duplicate-relationships";
import { wordChoiceRequiresUndeclaredCheck } from "./word-choice-requires";
import { wordSdtIdEmptyCheck } from "./word-sdt-id";
import { wordAttachedToolbarRelationshipCheck, wordAttachedToolbarUtf8Check } from "./word-attached-toolbar";
import { wordNumberingPictureBulletCheck } from "./word-numbering-picture-bullet";
import { wordFootnoteReferenceMissingCheck, wordEndnoteReferenceMissingCheck } from "./word-footnote-reference";
import { wordMathParseCheck, wordMathSpreBodyCheck } from "./word-body-math";
import { commentMarkersCheck, commentThreadingCheck } from "./comments";
import { xmlSyntaxCheck } from "./xml";
import { ctMarkupCompatibilityCheck } from "./ct-markup-compatibility";
import { wordPartContentTypeMismatchCheck } from "./word-part-content-type-mismatch";
import { wordPartContentTypeInexactCheck } from "./word-part-content-type-inexact";
import { wordHeaderFooterRelationshipCheck } from "./word-header-footer";
import { wordPartContentTypeUnrecognizedCheck } from "./word-part-content-type-unrecognized";
import { wordPartReferenceRelationshipCheck } from "./word-part-reference";

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
    wordDrawingExtentMisplacedCheck,
    wordPictureNvpicprMissingCheck,
    wordPictureSpprMissingCheck,
    wordRunMisplacedCheck,
    wordRunPropertiesMisplacedCheck,
    wordUnderlineMisplacedCheck,
    wordParagraphPropertiesMisplacedCheck,
    wordSectionPropertiesMisplacedCheck,
    wordTablePropertiesMisplacedCheck,
    wordUnknownElementCheck,
    wordLinkWithoutNamespaceCheck,
    wordMarkupCompatibilityCheck,
    wordDocumentRootCheck,
    wordMainDocumentPartCheck,
    opcRelationshipsStructureCheck,
    opcRelationshipsTargetModeCheck,
    wordRowWithoutCellCheck,
    wordMathParseCheck,
    wordMathSpreBodyCheck,
    wordFootnoteReferenceMissingCheck,
    wordEndnoteReferenceMissingCheck,
    wordNumberingPictureBulletCheck,
    wordAttachedToolbarRelationshipCheck,
    wordAttachedToolbarUtf8Check,
    wordSdtIdEmptyCheck,
    wordChoiceRequiresUndeclaredCheck,
    wordDuplicateThemeRelationshipCheck,
    wordDuplicateFontTableRelationshipCheck,
    wordDuplicateWebSettingsRelationshipCheck,
    wordDuplicateSettingsRelationshipCheck,
    wordDuplicateStylesRelationshipCheck,
    wordDuplicateNumberingRelationshipCheck,
    wordDuplicateEndnotesRelationshipCheck,
    wordDuplicateFootnotesRelationshipCheck,
    wordHyperlinkRelationshipCheck,
    wordHeaderFooterRelationshipCheck,
    wordPartReferenceRelationshipCheck,
    wordTextContainsParagraphCheck,
    wordTextboxDeletedTextCheck,
    wordMoveFromDeletedTextCheck,
    wordCellEndsWithTableCheck,
    wordEmptyDrawingCheck,
    wordAnchorMissingAttributeCheck,
    wordPositionOffsetAndAlignCheck,
    wordDrawingScalarWhitespaceCheck,
    wordDrawingArcRadiusCheck,
    wordDrawingExtentNegativeCheck,
    wordContentTypeInvalidCheck,
];

export function createDefaultCheckRegistry(): CheckRegistry {
    return new CheckRegistry([
        zipBytesAfterEocdCheck,
        zipBytesBeforeArchiveCheck,
        zipEocdCommentOverrunCheck,
        opcEquivalentPartNamesCheck,
        opcPartNameEscapedUnreservedCheck,
        opcPartNameEscapedSlashCheck,
        opcPartNameWordRefusedCheck,
        opcItemNameRefusedCheck,
        xmlSyntaxCheck,
        ctDuplicateDeclarationCheck,
        ctOverridePartnameInvalidCheck,
        ctOverridePartnameWordRefusedCheck,
        ctMarkupCompatibilityCheck,
        wordPartContentTypeGenericCheck,
        wordPartContentTypeInexactCheck,
        wordPartContentTypeMismatchCheck,
        wordPartContentTypeUnrecognizedCheck,
        commentMarkersCheck,
        commentThreadingCheck,
        ...WORD_RULE_CHECKS,
    ]);
}
