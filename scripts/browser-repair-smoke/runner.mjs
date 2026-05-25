import { chromium } from "playwright";
import { readFileSync } from "node:fs";
const b64 = readFileSync("/tmp/browser-repair/sample.docx").toString("base64");
const browser = await chromium.launch();
const page = await browser.newPage();
page.on("console", (m) => console.log("[page]", m.text()));
page.on("pageerror", (e) => console.log("[pageerror]", e.message));
await page.goto("http://localhost:8799/index.html");
await page.waitForTimeout(2500);
const diag = await page.evaluate(() => ({ ready: window.runRepairReady, hasFn: typeof window.runRepair, err: window.__err }));
console.log("DIAG:", JSON.stringify(diag).slice(0,400));
if (diag.hasFn === "function") {
  const result = await page.evaluate(async (b) => { try { return await window.runRepair(b); } catch (e) { return { ok:false, error:String(e&&e.stack||e) }; } }, b64);
  console.log("BROWSER RESULT:", JSON.stringify(result).slice(0,400));
}
await browser.close();
