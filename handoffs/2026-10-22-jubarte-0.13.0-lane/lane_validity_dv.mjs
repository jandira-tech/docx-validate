// Lane validity with docx-validate (XSD, libxml2-wasm): error kinds in each redline minus those either source has.
// Run from the bench root: node lane_validity_dv.mjs [lane]
import { validate } from "/Users/arthrod/temp/T/docx-validate/dist/index.mjs";
import fs from "node:fs";
const R = "results/redlines_0929_full", W = "corpus/word";
const LANE = process.argv[2] ?? "jubarte-0.13.0";
const renamed = {};
const man = `${R}/${LANE}/pdf_by_word/manifest.jsonl`;
if (fs.existsSync(man)) for (const l of fs.readFileSync(man, "utf8").split("\n")) {
  if (!l.trim()) continue; const r = JSON.parse(l);
  if (r.kind === "rename" && r.status === "renamed") renamed[r.original_filename] = r.converted_filename;
}
const laneDocx = (k) => { const n = `${k}_${LANE}.docx`; return `${R}/${LANE}/docx/${renamed[n] ?? n}`; };
const norm = (m) => m.replace(/^\s*-\s*/, "").replace(/'[^']*'(?=: This element is not expected)/, (s) => s)
  .replace(/line \d+/g, "line N").replace(/\d+ new error\(s\)/, "N new error(s)").slice(0, 220);
const cache = new Map();
async function kinds(p) {
  if (cache.has(p)) return cache.get(p);
  const got = new Map();
  if (!fs.existsSync(p)) got.set("ABSENT", 1);
  else {
    try {
      const r = await validate(p);
      for (const i of r.issues) if (i.severity === "error" && i.code !== "xsd-summary") {
        const k = `${i.code}\t${i.path ?? ""}\t${norm(i.message)}`; got.set(k, (got.get(k) ?? 0) + 1);
      }
    } catch (e) { got.set(`THROW\t\t${String(e).slice(0, 200)}`, 1); }
  }
  cache.set(p, got); return got;
}
const lines = fs.readFileSync("results/release_0.13.0_evidence/gen_pairs_sample.csv", "utf8").trim().split("\n");
const head = lines[0].split(","); const rows = lines.slice(1).map((l) => Object.fromEntries(l.split(",").map((v, i) => [head[i], v])));
let flagged = 0; const own = []; const table = new Map();
for (const r of rows) {
  const mine = await kinds(laneDocx(r.key));
  if (!mine.size) continue; flagged++;
  const src = new Set([...(await kinds(`${W}/${r.base}.docx`)).keys(), ...(await kinds(`${W}/${r.next}.docx`)).keys()]);
  const novel = [...mine].filter(([k]) => !src.has(k));
  if (novel.length) { own.push([r.key, novel]); for (const [k] of novel) table.set(k, (table.get(k) ?? 0) + 1); }
}
console.log(`${LANE} (docx-validate): ${rows.length} redlines; ${flagged} with validator errors; ${own.length} with a kind neither source has`);
for (const [k, n] of [...table].sort((a, b) => b[1] - a[1])) console.log(`  ${String(n).padStart(4)} redlines  ${k}`);
for (const [key, novel] of own) console.log("OWN", key, JSON.stringify(novel));
