// instantiateWasm is wired statically inside ./lib/libxml2.mjs (inline bytes),
// so no globalThis/ordering setup is needed here.
export * from "./lib/index.mjs";
