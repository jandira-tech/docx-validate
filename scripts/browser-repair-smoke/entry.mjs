import { repairDocxInMemory } from "/Users/arthrod/temp/T/docx-validate/dist/index.mjs";
import JSZip from "jszip";
globalThis.runRepair = async (base64) => {
  const bin = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const zip = await JSZip.loadAsync(bin);
  const parts = [];
  for (const name of Object.keys(zip.files)) {
    const f = zip.files[name];
    if (f.dir) continue;
    parts.push([name, await f.async("uint8array")]);
  }
  const { parts: repaired, repairs } = await repairDocxInMemory(parts);
  const dec = new TextDecoder("utf-8");
  const doc = repaired.find(([n]) => n === "word/document.xml");
  const docText = doc ? dec.decode(doc[1]) : "";
  return { ok: true, repairs, partCount: repaired.length,
    xmlSpaceCount: (docText.match(/xml:space="preserve"/g) || []).length,
    bufferDefined: typeof Buffer !== "undefined", processDefined: typeof process !== "undefined" };
};
globalThis.runRepairReady = true;
