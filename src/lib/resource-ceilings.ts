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
 * Named resource ceilings for hostile packages: a decompression bomb, an
 * entry flood, an attribute bomb, deep nesting, an entity bomb, a layout
 * bomb. A named ceiling is the strongest refusal a validator can give: the
 * finding names the limit it enforced, and the package is refused before the
 * work the bomb targets runs — ZIP ceilings before any entry is inflated,
 * XML ceilings in one linear pre-scan before any DOM is built.
 *
 * The defaults are deliberately generous: they sit far above every real
 * Office package (the fixture corpus tops out at 74 entries, 2.4 MB per
 * part, 68 elements deep, 21 attributes per element, no DTD) and just above
 * the format's own hard limits — a deflate stream cannot exceed ~1032:1, and
 * libxml2 without `XML_PARSE_HUGE` refuses deeper than 256 elements on its
 * own (#1159), so `maxDepth` mirrors the engine rather than tightening it.
 *
 * Callers override individual ceilings with `ceilings` on the validate
 * options, alongside `deadlineMs`.
 */

import type { ValidationIssue } from "./types";
import type { ZipCentralRecord } from "./zip-central-directory";

/** The resource ceilings of one run; every field is required once resolved. */
export type ResourceCeilings = {
    /** File entries the archive's central directory may declare (entry flood). */
    maxEntries: number;
    /** Declared uncompressed size of one part, in bytes (decompression bomb). */
    maxPartBytes: number;
    /** Declared uncompressed/compressed ratio of one entry (a lie past deflate's ~1032:1 ceiling). */
    maxCompressionRatio: number;
    /** Cumulative declared uncompressed size of the whole archive, in bytes. */
    maxTotalUncompressedBytes: number;
    /** Element nesting depth of one XML part. Mirrors libxml2's own limit without `XML_PARSE_HUGE`. */
    maxDepth: number;
    /** Attributes on one XML element (attribute bomb). */
    maxAttributes: number;
    /** `<!ENTITY` declarations in one XML part (entity bomb backstop; the parser never expands them). */
    maxEntityDeclarations: number;
    /** `w:gridSpan` value of one table cell (Word's table column limit). */
    maxGridSpan: number;
    /** Field markers (`w:fldChar`, `w:instrText`, `w:fldSimple`) in one story part. */
    maxFieldMarkers: number;
};

/** The caller-supplied subset of {@link ResourceCeilings}; omitted fields keep their defaults. */
export type ResourceCeilingsInput = Partial<ResourceCeilings>;

/** The ceilings every run enforces when the caller does not override them. */
export const DEFAULT_RESOURCE_CEILINGS: ResourceCeilings = {
    maxEntries: 10_000,
    maxPartBytes: 268_435_456,
    maxCompressionRatio: 1_100,
    maxTotalUncompressedBytes: 2_147_483_648,
    maxDepth: 256,
    maxAttributes: 1_024,
    maxEntityDeclarations: 64,
    maxGridSpan: 63,
    maxFieldMarkers: 10_000,
};

/** Ceilings that must be whole numbers when given. */
const INTEGER_CEILINGS: ReadonlySet<string> = new Set([
    "maxEntries",
    "maxPartBytes",
    "maxTotalUncompressedBytes",
    "maxDepth",
    "maxAttributes",
    "maxEntityDeclarations",
    "maxGridSpan",
    "maxFieldMarkers",
]);

/**
 * Reject a ceilings object whose values are not finite numbers greater than
 * zero (and whole numbers where the ceiling counts things).
 *
 * @throws {RangeError}
 */
export const assertResourceCeilings: (value: unknown) => asserts value is ResourceCeilingsInput = (value: unknown): void => {
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
        throw new RangeError(`ceilings must be an object of ceiling names to numbers; got ${String(value)}`);
    }
    for (const [key, ceiling] of Object.entries(value as Record<string, unknown>)) {
        if (typeof ceiling !== "number" || !Number.isFinite(ceiling) || ceiling <= 0) {
            throw new RangeError(`ceilings.${key} must be a finite number greater than 0; got ${String(ceiling)}`);
        }
        if (INTEGER_CEILINGS.has(key) && !Number.isInteger(ceiling)) {
            throw new RangeError(`ceilings.${key} must be a whole number; got ${String(ceiling)}`);
        }
    }
};

/** Merge caller-supplied ceilings over the defaults, validating what was given. */
export const resolveResourceCeilings = (input?: ResourceCeilingsInput): ResourceCeilings => {
    if (input !== undefined) {
        assertResourceCeilings(input);
    }
    return { ...DEFAULT_RESOURCE_CEILINGS, ...input };
};

/** Code of the entry-count ceiling finding. */
export const ZIP_ENTRY_COUNT_CEILING = "zip-entry-count-ceiling";
/** Code of the per-part declared-size ceiling finding. */
export const ZIP_PART_SIZE_CEILING = "zip-part-size-ceiling";
/** Code of the per-entry compression-ratio ceiling finding. */
export const ZIP_COMPRESSION_RATIO_CEILING = "zip-compression-ratio-ceiling";
/** Code of the cumulative declared-size ceiling finding. */
export const ZIP_TOTAL_UNCOMPRESSED_CEILING = "zip-total-uncompressed-ceiling";

/** How many entry names a ceiling finding lists before counting the rest. */
const NAMED_ENTRIES = 8;

const nameEntries = (records: readonly ZipCentralRecord[]): string => {
    const named = records.slice(0, NAMED_ENTRIES).map((record) => JSON.stringify(record.name));
    const more = records.length - named.length;
    return more > 0 ? `${named.join(", ")}; and ${more} more` : named.join(", ");
};

/**
 * The ZIP ceilings, judged from the central directory's declared fields
 * alone — before any entry is inflated. `undefined` records (an envelope the
 * directory reader refuses) produce nothing here: the archive fails to open
 * and is reported as such. Every finding is an error and names the ceiling
 * it enforced.
 */
export const zipCeilingIssues = (records: readonly ZipCentralRecord[] | undefined, ceilings: ResourceCeilings): ValidationIssue[] => {
    if (records === undefined) {
        return [];
    }
    const issues: ValidationIssue[] = [];
    const files = records.filter((record) => !record.directory);
    if (files.length > ceilings.maxEntries) {
        issues.push({
            severity: "error",
            code: ZIP_ENTRY_COUNT_CEILING,
            message: `Archive declares ${files.length} file entries, over the entry-count ceiling of ${ceilings.maxEntries} (maxEntries). Refused before any entry is inflated.`,
        });
    }
    const oversized = files.filter((record) => record.uncompressedSize > ceilings.maxPartBytes);
    if (oversized.length > 0) {
        issues.push({
            severity: "error",
            code: ZIP_PART_SIZE_CEILING,
            message: `${oversized.length} ${oversized.length === 1 ? "entry declares" : "entries declare"} an uncompressed size over the per-part ceiling of ${ceilings.maxPartBytes} bytes (maxPartBytes): ${nameEntries(oversized)}. Refused before any entry is inflated.`,
        });
    }
    const inflated = files.filter(
        (record) => record.compressedSize > 0 && record.uncompressedSize / record.compressedSize > ceilings.maxCompressionRatio,
    );
    if (inflated.length > 0) {
        const worst = inflated.reduce((a, b) => (a.uncompressedSize / a.compressedSize > b.uncompressedSize / b.compressedSize ? a : b));
        issues.push({
            severity: "error",
            code: ZIP_COMPRESSION_RATIO_CEILING,
            message: `${inflated.length} ${inflated.length === 1 ? "entry declares" : "entries declare"} a compression ratio over the ceiling of ${ceilings.maxCompressionRatio}:1 (maxCompressionRatio), worst ${Math.floor(worst.uncompressedSize / Math.max(1, worst.compressedSize))}:1: ${nameEntries(inflated)}. A real deflate stream tops out near 1032:1, so the declared sizes lie. Refused before any entry is inflated.`,
        });
    }
    const total = files.reduce((sum, record) => sum + record.uncompressedSize, 0);
    if (total > ceilings.maxTotalUncompressedBytes) {
        issues.push({
            severity: "error",
            code: ZIP_TOTAL_UNCOMPRESSED_CEILING,
            message: `Archive declares ${total} uncompressed bytes in total, over the cumulative ceiling of ${ceilings.maxTotalUncompressedBytes} bytes (maxTotalUncompressedBytes). Refused before any entry is inflated.`,
        });
    }
    return issues;
};
