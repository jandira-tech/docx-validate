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

import { resolveWordOutcome } from "../scripts/update-manifest";

describe("resolveWordOutcome", () => {
  const prior = new Map<string, string>([["broken/a.docx", "clean-open"]]);

  it("prefers a fresh probe outcome", () => {
    expect(resolveWordOutcome("broken/a.docx", "recovered", prior)).toBe("recovered");
  });

  it("falls back to the prior manifest value when no probe", () => {
    expect(resolveWordOutcome("broken/a.docx", undefined, prior)).toBe("clean-open");
  });

  it("is 'unknown' when neither probe nor prior value exists", () => {
    expect(resolveWordOutcome("broken/new.docx", undefined, prior)).toBe("unknown");
  });
});
