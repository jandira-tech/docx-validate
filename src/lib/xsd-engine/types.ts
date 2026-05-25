/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Pluggable XSD validation engine. Two backends share this contract:
//   - node.ts    → libxmljs2 (native libxml2 binding; fast; Node-only)
//   - browser.ts → libxml2-wasm (WebAssembly libxml2; runs in browser/worker)
// The validator picks one at runtime by environment, so XSD validation works in
// Node AND the browser without changing the rest of the pipeline.

export interface XsdValidationOutcome {
    /** `null` = no schema configured / engine unavailable (caller skips). */
    valid: boolean | null;
    errors: Set<string>;
}

export interface XsdEngine {
    /** One-time async init (WASM load). No-op for the native backend. */
    init(): Promise<void>;
    /** Throw a descriptive error if XSD validation cannot run on this host. */
    assertAvailable(): void;
    /**
     * Validate already-preprocessed XML against the schema identified by
     * `schemaRelKey` — a path relative to the schema root, e.g.
     * `"ISO-IEC29500-4_2016/wml.xsd"`. The engine loads and caches the schema,
     * resolving `xs:include`/`xs:import` from its own schema source (disk for
     * node; an in-memory bundle for browser).
     */
    validate(cleanedXml: string, schemaRelKey: string): XsdValidationOutcome;
}

export interface XsdEngineConfig {
    /** Schema root directory (native backend reads schemas from here). */
    schemasDir?: string;
    /**
     * In-memory schema bundle keyed by path relative to the schema root
     * (browser backend). When omitted, the browser engine uses its generated
     * default bundle.
     */
    schemaBundle?: Record<string, string>;
}
