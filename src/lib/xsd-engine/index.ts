/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Runtime engine selection: native libxmljs2 in Node, WASM libxml2-wasm in the
// browser/worker. Both engine modules are imported, but each loads its heavy
// dependency lazily (createRequire / dynamic import), so only the selected
// engine's library is ever loaded.

import { createXsdEngine as createBrowserEngine } from "./browser";
import { createXsdEngine as createNodeEngine } from "./node";
import type { XsdEngine, XsdEngineConfig } from "./types";

export type { XsdEngine, XsdEngineConfig, XsdValidationOutcome } from "./types";
export { createNodeEngine, createBrowserEngine };

/**
 * Node when there is a real `process.versions.node` AND a global `Buffer`
 * (absent in browsers and in Workers without nodejs_compat). Otherwise the
 * WASM engine, which runs everywhere.
 */
export function isNodeRuntime(): boolean {
    return (
        typeof process !== "undefined" &&
        !!process.versions &&
        typeof process.versions.node === "string" &&
        process.versions.node.length > 0 &&
        typeof Buffer !== "undefined"
    );
}

export function createXsdEngine(config: XsdEngineConfig): XsdEngine {
    return isNodeRuntime() ? createNodeEngine(config) : createBrowserEngine(config);
}
