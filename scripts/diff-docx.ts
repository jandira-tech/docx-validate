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

import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";

import { runCli, withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { unpack } from "../src/scripts/office/unpack";
import { collectDocxSemanticInventory } from "../src/scripts/office/validators/docx-diagnostics";
import {
  diffDocxInventories,
  formatInventoryDiffMarkdown,
  inventoryDiffToIssues,
} from "../src/scripts/office/validators/docx-inventory-diff";

async function isDirectory(p: string): Promise<boolean> {
  try {
    return (await fs.stat(p)).isDirectory();
  } catch {
    return false;
  }
}

async function inventoryOf(
  input: string,
  profile: Profile,
): Promise<ReturnType<typeof collectDocxSemanticInventory>> {
  if (await isDirectory(input)) return collectDocxSemanticInventory(input, profile);
  // packed file → unpack into a temp dir, then collect
  return withTempDir(async (tmp) => {
    const out = path.join(tmp, "unpacked");
    const res = await unpack(input, out);
    if (!res.ok) throw new Error(res.message);
    return collectDocxSemanticInventory(out, profile);
  });
}

export async function runDiffDocx(
  args: readonly string[],
): Promise<{ code: number; markdown: string }> {
  const cmd = new Command();
  cmd
    .name("diff-docx")
    .description(
      "Symmetric semantic fingerprint diff between two DOCX inputs (packed file or unpacked dir).",
    )
    .argument("<a>", "first document (file or unpacked dir)")
    .argument("<b>", "second document (file or unpacked dir)")
    .option("--profile <profile>", "lenient | strict | word-valid", "lenient")
    .allowExcessArguments(false);
  cmd.exitOverride();
  cmd.parse(args as string[], { from: "user" });
  const opts = cmd.opts<{ profile: string }>();
  if (opts.profile !== "lenient" && opts.profile !== "strict" && opts.profile !== "word-valid") {
    const bad = String(opts.profile);
    return {
      code: 1,
      markdown: `Invalid --profile: ${bad}. Must be 'lenient', 'strict', or 'word-valid'.`,
    };
  }
  const profile = opts.profile as Profile;
  const [a, b] = cmd.args;

  const diff = diffDocxInventories(await inventoryOf(a, profile), await inventoryOf(b, profile));
  const markdown = formatInventoryDiffMarkdown(diff);
  const hasError = inventoryDiffToIssues(diff).some((i) => i.severity === "error");
  return { code: hasError ? 1 : 0, markdown };
}

runCli(import.meta.url, async () => {
  const { code, markdown } = await runDiffDocx(process.argv.slice(2));
  if (code !== 0) {
    process.stderr.write(`${markdown}\n`);
    return code;
  }
  process.stdout.write(`${markdown}\n`);
  return code;
});
