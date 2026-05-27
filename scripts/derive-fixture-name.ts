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
    category: "broken" | "working";
    subjectSlug: string;
    descriptor: string;
    /** `${subjectSlug}.${descriptor}.docx` */
    fileName: string;
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

function descriptorFor(fp: FixtureFingerprint): string {
    if (fp.strictErrorCodes.length > 0) {
        return [...fp.strictErrorCodes].sort()[0];
    }
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
    return "plain-paragraphs";
}

export function deriveName(fp: FixtureFingerprint): DerivedName {
    const category: "broken" | "working" = fp.strictErrorCodes.length > 0 ? "broken" : "working";
    const subjectSlug = (fp.titleText && slugify(fp.titleText)) || "document";
    const descriptor = descriptorFor(fp);
    return { category, subjectSlug, descriptor, fileName: `${subjectSlug}.${descriptor}.docx` };
}
