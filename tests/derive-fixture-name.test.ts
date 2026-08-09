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

import { describe, expect, it } from "vitest";

import { type FixtureFingerprint, deriveName, slugify } from "../scripts/derive-fixture-name";

const base: FixtureFingerprint = {
  strictErrorCodes: [],
  lenientErrorCodes: [],
  insCount: 0,
  delCount: 0,
  commentCount: 0,
  firstCommentText: null,
  tableCount: 0,
  hasTextBox: false,
  hasHeaderFooter: false,
  titleText: null,
  contentHash: "deadbeef",
};

describe("slugify", () => {
  it("lowercases, hyphenates, and caps word count", () => {
    expect(slugify("Q4 2025 NPS Survey Results Extra", 4)).toBe("q4-2025-nps-survey");
  });
  it("returns empty string for punctuation-only input", () => {
    expect(slugify("!!! ???")).toBe("");
  });
});

describe("deriveName", () => {
  it("routes a file with strict errors to broken/ using the first error code", () => {
    const d = deriveName({ ...base, strictErrorCodes: ["tables-broken-rels", "another-code"] });
    expect(d.category).toBe("broken");
    expect(d.descriptor).toBe("another-code"); // alphabetically first
    expect(d.fileName).toBe("document.another-code.docx");
  });

  it("describes insertions-only tracked changes", () => {
    const d = deriveName({ ...base, insCount: 3 });
    expect(d.category).toBe("working");
    expect(d.fileName).toBe("document.suggesting-insertions.docx");
  });

  it("describes deletions-only tracked changes", () => {
    const d = deriveName({ ...base, delCount: 2 });
    expect(d.fileName).toBe("document.suggesting-deletions.docx");
  });

  it("describes mixed tracked changes", () => {
    const d = deriveName({ ...base, insCount: 1, delCount: 1 });
    expect(d.fileName).toBe("document.suggesting-mixed-edits.docx");
  });

  it("describes a comment using its slugified gist", () => {
    const d = deriveName({
      ...base,
      commentCount: 1,
      firstCommentText: "Please review this clause",
    });
    expect(d.fileName).toBe("document.comment-please-review-this-clause.docx");
  });

  it("falls back to a structural descriptor for a clean table doc", () => {
    const d = deriveName({ ...base, tableCount: 2 });
    expect(d.fileName).toBe("document.table.docx");
  });

  it("uses the slugified title as the subject when present", () => {
    const d = deriveName({ ...base, titleText: "Master Services Agreement", insCount: 1 });
    expect(d.fileName).toBe("master-services-agreement.suggesting-insertions.docx");
  });

  it("falls back to plain-paragraphs for an otherwise featureless clean doc", () => {
    expect(deriveName(base).fileName).toBe("document.plain-paragraphs.docx");
  });
});

describe("deriveName — content-first descriptor mode", () => {
  const opts = { descriptorMode: "content-first" as const };

  it("prefers the content feature over a present error code", () => {
    // strict error AND a table: content-first names it for the table, not the error
    const d = deriveName(
      { ...base, strictErrorCodes: ["id-paraid-overflow"], tableCount: 2 },
      opts,
    );
    expect(d.descriptor).toBe("table");
  });

  it("prefers tracked changes over a present error code", () => {
    const d = deriveName(
      { ...base, strictErrorCodes: ["id-paraid-overflow"], insCount: 2, delCount: 1 },
      opts,
    );
    expect(d.descriptor).toBe("suggesting-mixed-edits");
  });

  it("falls back to the first error code when there is no content signal", () => {
    const d = deriveName(
      { ...base, strictErrorCodes: ["style-default-missing", "id-paraid-overflow"] },
      opts,
    );
    expect(d.descriptor).toBe("id-paraid-overflow"); // alphabetically first strict code
  });

  it("falls back to a lenient error code when there is no strict error or content", () => {
    const d = deriveName({ ...base, lenientErrorCodes: ["only-lenient"] }, opts);
    expect(d.descriptor).toBe("only-lenient");
  });

  it("falls back to plain-paragraphs when there is neither content nor error", () => {
    expect(deriveName(base, opts).descriptor).toBe("plain-paragraphs");
  });
});

describe("deriveName — fixedCategory override", () => {
  it("uses the fixed category regardless of error state", () => {
    const broken = deriveName({ ...base, strictErrorCodes: ["x"] }, { fixedCategory: "eigen" });
    const clean = deriveName({ ...base, insCount: 1 }, { fixedCategory: "eigen" });
    expect(broken.category).toBe("eigen");
    expect(clean.category).toBe("eigen");
  });
});
