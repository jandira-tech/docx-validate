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

import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";

import JSZip from "jszip";

import { NS } from "../src/lib/types";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import { validate } from "../src/scripts/office/validate";
import type { FixtureFingerprint } from "./derive-fixture-name";

const DC = "http://purl.org/dc/elements/1.1/";

async function errorCodes(file: string, profile: "strict" | "lenient"): Promise<string[]> {
  const result = await validate(file, { profile });
  return Array.from(
    new Set(
      result.issues
        .filter((i) => i.severity === "error")
        .map((i) => i.code)
        .filter((c): c is string => Boolean(c)),
    ),
  ).sort();
}

function firstParagraphText(doc: Document): string | null {
  const paras = getElementsByTagNameNSAll(doc, NS.W, "p");
  for (const p of paras) {
    const text = getElementsByTagNameNSAll(p, NS.W, "t")
      .map((t) => t.textContent ?? "")
      .join("")
      .trim();
    if (text) return text;
  }
  return null;
}

export async function fingerprint(docxPath: string): Promise<FixtureFingerprint> {
  const buf = await fs.readFile(docxPath);
  const zip = await JSZip.loadAsync(buf);

  const docXml = (await zip.file("word/document.xml")?.async("string")) ?? "";
  const contentHash = createHash("sha256").update(docXml).digest("hex");

  // Some broken fixtures have malformed XML (e.g. undeclared namespace prefixes) that
  // parseXml will reject. Gracefully fall back to raw-string heuristics in that case.
  let doc: Document | null = null;
  try {
    doc = parseXml(docXml);
  } catch {
    // leave doc null; structural counts fall back to 0 / string-scan below
  }

  const insCount = doc ? getElementsByTagNameNSAll(doc, NS.W, "ins").length : 0;
  const delCount = doc ? getElementsByTagNameNSAll(doc, NS.W, "del").length : 0;
  const tableCount = doc ? getElementsByTagNameNSAll(doc, NS.W, "tbl").length : 0;
  const hasTextBox = docXml.includes("txbxContent") || docXml.includes("textbox");
  const hasHeaderFooter = zip.file(/word\/(header|footer)\d*\.xml$/).length > 0;

  let commentCount = 0;
  let firstCommentText: string | null = null;
  const commentsXml = await zip.file("word/comments.xml")?.async("string");
  if (commentsXml) {
    try {
      const cdoc = parseXml(commentsXml);
      const comments = getElementsByTagNameNSAll(cdoc, NS.W, "comment");
      commentCount = comments.length;
      if (comments[0]) {
        const text = getElementsByTagNameNSAll(comments[0], NS.W, "t")
          .map((t) => t.textContent ?? "")
          .join("")
          .trim();
        firstCommentText = text || null;
      }
    } catch {
      // malformed comments.xml — skip
    }
  }

  let titleText: string | null = null;
  const coreXml = await zip.file("docProps/core.xml")?.async("string");
  if (coreXml) {
    try {
      const core = parseXml(coreXml);
      const title = getElementsByTagNameNSAll(core, DC, "title")[0]?.textContent?.trim();
      if (title) titleText = title;
    } catch {
      // malformed core.xml — skip title extraction
    }
  }
  if (!titleText) titleText = doc ? firstParagraphText(doc) : null;

  const [strictErrorCodes, lenientErrorCodes] = await Promise.all([
    errorCodes(docxPath, "strict"),
    errorCodes(docxPath, "lenient"),
  ]);

  return {
    strictErrorCodes,
    lenientErrorCodes,
    insCount,
    delCount,
    commentCount,
    firstCommentText,
    tableCount,
    hasTextBox,
    hasHeaderFooter,
    titleText,
    contentHash,
  };
}
