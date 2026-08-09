/**
 * Batch word-validity validator — validates many .docx files in ONE process.
 *
 * Loads the libxmljs2 native binding exactly once and validates a list of
 * files sequentially. This avoids the per-file `npx tsx` spawning that races
 * npm's node_modules self-repair and wipes the native binding under load.
 *
 * Usage (run with the local tsx, NOT npx):
 *   node_modules/.bin/tsx scripts/batch-validate.ts <list.json> <out.json> [profile]
 *
 * <list.json> : JSON array of { "id": string, "path": string }
 * <out.json>  : written as JSON array of
 *               { "id": string, "valid": boolean, "errors": [{part,message}] }
 */
import { readFileSync, writeFileSync } from "node:fs";

import { validate } from "../src/scripts/office/validate";
import type { Profile } from "../src/lib/types";

type Item = { id: string; path: string };

async function main(): Promise<void> {
  const [, , listPath, outPath, profileArg] = process.argv;
  if (!listPath || !outPath) {
    process.stderr.write("usage: batch-validate.ts <list.json> <out.json> [profile]\n");
    process.exit(2);
  }
  const profile = (profileArg ?? "word-valid") as Profile;
  const items: Item[] = JSON.parse(readFileSync(listPath, "utf8"));
  const out: Array<{
    id: string;
    valid: boolean;
    errors: Array<{ part: string; message: string }>;
  }> = [];
  for (const it of items) {
    try {
      const res = await validate(it.path, { profile });
      const errors = (res.issues ?? [])
        .filter((i) => i.severity === "error")
        .map((i) => ({ part: i.path ?? "<unknown>", message: i.message }));
      out.push({ id: it.id, valid: Boolean(res.valid) && errors.length === 0, errors });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      out.push({ id: it.id, valid: false, errors: [{ part: "<cli>", message }] });
    }
  }
  writeFileSync(outPath, JSON.stringify(out));
}

main().catch((e: unknown) => {
  process.stderr.write(`${e instanceof Error ? e.stack : String(e)}\n`);
  process.exit(2);
});
