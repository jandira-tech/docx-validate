// Walk a dir of *.docx, run docx-validate word-valid profile on each.
// Usage: node --import tsx scripts/sweep-word-valid.ts <dir> [--json out.json]
import * as fs from "node:fs";
import * as path from "node:path";
import { validate } from "../src/scripts/office/validate.ts";

const args = process.argv.slice(2);
const dir = path.resolve(args[0]!);
const jsonIdx = args.indexOf("--json");
const jsonOut = jsonIdx >= 0 ? path.resolve(args[jsonIdx + 1]!) : "";

function* walk(d: string): Generator<string> {
  for (const n of fs.readdirSync(d)) {
    const abs = path.join(d, n);
    if (fs.statSync(abs).isDirectory()) yield* walk(abs);
    else if (n.toLowerCase().endsWith(".docx")) yield abs;
  }
}

const files = [...walk(dir)].sort();
let ok = 0, fail = 0, err = 0;
const errors: string[] = [];
for (const f of files) {
  try {
    const r: any = await validate(f, { profile: "word-valid" });
    if (r.valid) ok++;
    else {
      fail++;
      const e = (r.issues || []).filter((i: any) => (i.severity ?? "error") === "error")[0];
      errors.push(`${path.basename(f)}: ${e?.code ?? ""} ${e?.message?.slice(0, 90) ?? "?"}`);
    }
  } catch (e: any) {
    err++;
    errors.push(`${path.basename(f)}: THREW ${String(e?.message || e).slice(0, 90)}`);
  }
}
console.log(`word-valid: ${ok}/${files.length} ok | ${fail} invalid | ${err} threw`);
console.log("--- first 25 failures ---");
for (const e of errors.slice(0, 25)) console.log("  " + e);
if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify({ dir, total: files.length, ok, fail, err, errors }, null, 2));
