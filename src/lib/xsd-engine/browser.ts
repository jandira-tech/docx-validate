/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Browser/worker XSD engine — libxml2-wasm (WebAssembly libxml2). No native
// binding, no Node `Buffer`. Schemas come from an in-memory bundle keyed by
// path relative to the schema root; xs:include/import resolve through libxml2's
// input-provider seeded with that bundle (proven against the full OOXML
// wml.xsd graph). libxml2-wasm is dynamically imported so Node never loads the
// WASM when it selects the native engine.

import type { XsdEngine, XsdEngineConfig, XsdValidationOutcome } from "./types";

type Libxml2 = typeof import("libxml2-wasm");

export function createXsdEngine(config: XsdEngineConfig): XsdEngine {
    let lib: Libxml2 | null = null;
    let providerRegistered = false;
    const bundle = config.schemaBundle ?? {};
    // Cache one reusable XsdValidator per schema (parsed once).
    const validators = new Map<string, ReturnType<Libxml2["XsdValidator"]["fromDoc"]>>();

    return {
        async init() {
            if (!lib) lib = await import("libxml2-wasm");
            if (!providerRegistered && Object.keys(bundle).length > 0) {
                const enc = new TextEncoder();
                const buffers: Record<string, Uint8Array> = {};
                for (const key of Object.keys(bundle)) buffers[key] = enc.encode(bundle[key]);
                lib.xmlRegisterInputProvider(new lib.XmlBufferInputProvider(buffers));
                providerRegistered = true;
            }
        },
        assertAvailable() {
            if (!lib) throw new Error("libxml2-wasm not initialized — call init() before validate()");
        },
        validate(cleanedXml: string, schemaRelKey: string): XsdValidationOutcome {
            if (!lib) return { valid: null, errors: new Set() };
            const schemaSource = bundle[schemaRelKey];
            if (!schemaSource) return { valid: null, errors: new Set() };
            try {
                let validator = validators.get(schemaRelKey);
                if (!validator) {
                    const xsdDoc = lib.XmlDocument.fromString(schemaSource, { url: schemaRelKey });
                    validator = lib.XsdValidator.fromDoc(xsdDoc);
                    validators.set(schemaRelKey, validator);
                }
                const doc = lib.XmlDocument.fromString(cleanedXml);
                try {
                    validator.validate(doc);
                    return { valid: true, errors: new Set() };
                } catch (e) {
                    return { valid: false, errors: new Set([String((e as { message?: string }).message ?? e)]) };
                } finally {
                    doc.dispose();
                }
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                return { valid: false, errors: new Set([message]) };
            }
        },
    };
}
