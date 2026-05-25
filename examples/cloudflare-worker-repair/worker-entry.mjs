import { repairDocxInMemory } from "/Users/arthrod/temp/T/docx-validate/dist/index.mjs";
import JSZip from "jszip";
const HTML = `<!doctype html><meta charset=utf-8><h1>docx repair (Cloudflare Worker / no-Node)</h1>
<input type=file id=f multiple accept=.docx><pre id=out>drop .docx files…</pre>
<script>
f.onchange=async()=>{out.textContent="";for(const file of f.files){const b=await file.arrayBuffer();
const r=await fetch("/api/repair",{method:"POST",body:b});const j=await r.json();
out.textContent+=file.name+" -> repairs="+j.repairs+" parts="+j.partCount+" buffer="+j.bufferDefined+"\\n";}}
</script>`;
export default {
  async fetch(req) {
    const url = new URL(req.url);
    if (req.method === "POST" && url.pathname === "/api/repair") {
      try {
        const bytes = new Uint8Array(await req.arrayBuffer());
        const zip = await JSZip.loadAsync(bytes);
        const parts = [];
        for (const name of Object.keys(zip.files)) { const e = zip.files[name]; if (e.dir) continue; parts.push([name, await e.async("uint8array")]); }
        const { parts: repaired, repairs } = await repairDocxInMemory(parts);
        const out = new JSZip();
        for (const [n, c] of repaired) out.file(n, c);
        const repackedB64 = await out.generateAsync({ type: "base64" });
        return Response.json({ ok: true, repairs, partCount: repaired.length, bufferDefined: typeof Buffer !== "undefined", repackedB64 });
      } catch (e) { return Response.json({ ok: false, error: String(e && e.stack || e) }, { status: 500 }); }
    }
    return new Response(HTML, { headers: { "content-type": "text/html;charset=utf-8" } });
  },
};
