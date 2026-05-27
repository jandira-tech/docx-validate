import JSZip from "jszip";
import { promises as fs } from "node:fs";
async function main() {
  for (const [src, out] of [
    [".drift-run/copies/external/open-xml-sdk/mcdoc.docx", ".drift-run/mcdoc-orig.xml"],
    [".drift-run/repaired/external/open-xml-sdk/mcdoc.docx", ".drift-run/mcdoc-rep.xml"],
  ]) {
    const z = await JSZip.loadAsync(await fs.readFile(src));
    await fs.writeFile(out, await z.file("word/document.xml")!.async("string"));
  }
}
main();
