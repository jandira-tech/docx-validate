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

/**
 * Pure, deterministic derivation of a content-descriptive fixture name from a
 * docx content fingerprint. See
 * docs/superpowers/specs/2026-05-27-repo-cleanup-fixture-naming-design.md.
 */

export interface FixtureFingerprint {
    /** Distinct, sorted strict-profile error codes. */
    strictErrorCodes: string[];
    /** Distinct, sorted lenient-profile error codes. */
    lenientErrorCodes: string[];
    insCount: number;
    delCount: number;
    commentCount: number;
    firstCommentText: string | null;
    tableCount: number;
    hasTextBox: boolean;
    hasHeaderFooter: boolean;
    /** dc:title or first non-empty paragraph text, if any. */
    titleText: string | null;
    /** sha256 of word/document.xml, for dedup. */
    contentHash: string;
}

export interface DerivedName {
    /** "broken" / "working" by default, or an explicit `fixedCategory`. */
    category: string;
    subjectSlug: string;
    descriptor: string;
    /** `${subjectSlug}.${descriptor}.docx` */
    fileName: string;
}

/**
 * How the descriptor is chosen:
 * - "error-first" (default): lead with the strict error code; fall back to the
 *   content feature, then "plain-paragraphs". Suits curated synthetic fixtures
 *   where the validation break is the point.
 * - "content-first": lead with the distinguishing content feature
 *   (tracked-changes / comment / table / text-box / header-footer); fall back to
 *   the first error code (strict, else lenient), then "plain-paragraphs". Suits
 *   real-world specimens that share a benign, near-universal validation quirk —
 *   the content is what differentiates them.
 */
export type DescriptorMode = "error-first" | "content-first";

export interface DeriveOptions {
    descriptorMode?: DescriptorMode;
    /** Force the category instead of deriving "broken"/"working" from errors. */
    fixedCategory?: string;
}

/** Lowercase kebab slug, capped at `maxWords` hyphen-delimited words. */
export function slugify(text: string, maxWords = 4): string {
    const words = text
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "")
        .split("-")
        .filter(Boolean);
    return words.slice(0, maxWords).join("-");
}

/** The distinguishing content feature, or null when the doc has none. */
function contentDescriptor(fp: FixtureFingerprint): string | null {
    if (fp.insCount > 0 && fp.delCount > 0) return "suggesting-mixed-edits";
    if (fp.insCount > 0) return "suggesting-insertions";
    if (fp.delCount > 0) return "suggesting-deletions";
    if (fp.commentCount > 0) {
        const gist = slugify(fp.firstCommentText ?? "", 4);
        return gist ? `comment-${gist}` : "comment";
    }
    if (fp.tableCount > 0) return "table";
    if (fp.hasTextBox) return "text-box";
    if (fp.hasHeaderFooter) return "header-footer";
    return null;
}

/** First strict error code, else first lenient error code, else null. */
function firstErrorCode(fp: FixtureFingerprint): string | null {
    if (fp.strictErrorCodes.length > 0) return [...fp.strictErrorCodes].sort()[0];
    if (fp.lenientErrorCodes.length > 0) return [...fp.lenientErrorCodes].sort()[0];
    return null;
}

function descriptorFor(fp: FixtureFingerprint, mode: DescriptorMode): string {
    if (mode === "content-first") {
        return contentDescriptor(fp) ?? firstErrorCode(fp) ?? "plain-paragraphs";
    }
    // error-first: strict error wins, then content feature, then plain.
    if (fp.strictErrorCodes.length > 0) return [...fp.strictErrorCodes].sort()[0];
    return contentDescriptor(fp) ?? "plain-paragraphs";
}

export function deriveName(fp: FixtureFingerprint, opts: DeriveOptions = {}): DerivedName {
    const mode = opts.descriptorMode ?? "error-first";
    const category = opts.fixedCategory ?? (fp.strictErrorCodes.length > 0 ? "broken" : "working");
    const subjectSlug = (fp.titleText && slugify(fp.titleText)) || "document";
    const descriptor = descriptorFor(fp, mode);
    return { category, subjectSlug, descriptor, fileName: `${subjectSlug}.${descriptor}.docx` };
}
