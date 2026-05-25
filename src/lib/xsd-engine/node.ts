/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Native XSD engine — libxmljs2 (compiled libxml2). Fast; Node-only. The import
// is TYPE-only + a lazy createRequire loader so a browser bundle that selects
// the browser engine never pulls the native binding.

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

import type * as LibXmlJs from "libxmljs2";
import type { XsdEngine, XsdEngineConfig, XsdValidationOutcome } from "./types";

let _libxmljs: typeof LibXmlJs | null = null;
function libxmljs(): typeof LibXmlJs {
    if (!_libxmljs) {
        const require = createRequire(import.meta.url);
        _libxmljs = require("libxmljs2") as typeof LibXmlJs;
    }
    return _libxmljs;
}

// Process-wide parsed-XSD cache (the OOXML bundle is ~1 MB; reused per file).
const xsdCache = new Map<string, LibXmlJs.Document>();

export function createXsdEngine(config: XsdEngineConfig): XsdEngine {
    const schemasDir = config.schemasDir ?? "";

    function loadXsd(absPath: string): LibXmlJs.Document {
        const abs = path.resolve(absPath);
        const hit = xsdCache.get(abs);
        if (hit) return hit;
        const content = readFileSync(abs, "utf-8");
        // baseUrl lets <xs:include>/<xs:import> resolve sibling schemas on disk.
        const doc = libxmljs().parseXml(content, { baseUrl: abs });
        xsdCache.set(abs, doc);
        return doc;
    }

    return {
        async init() {},
        assertAvailable() {
            try {
                const xsd =
                    '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="r" type="xs:string"/></xs:schema>';
                const xsdDoc = libxmljs().parseXml(xsd);
                const doc = libxmljs().parseXml("<r>ok</r>");
                const ok = doc.validate(xsdDoc);
                if (ok !== true) throw new Error(`validate() returned ${String(ok)} on a known-good doc`);
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                throw new Error(`libxmljs2 required for XSD validation: ${message}`);
            }
        },
        validate(cleanedXml: string, schemaRelKey: string): XsdValidationOutcome {
            try {
                const xsdDoc = loadXsd(path.join(schemasDir, schemaRelKey));
                const xmlLibDoc = libxmljs().parseXml(cleanedXml);
                const valid = xmlLibDoc.validate(xsdDoc);
                if (valid) return { valid: true, errors: new Set() };
                const errors = new Set<string>();
                const errs =
                    (xmlLibDoc as unknown as { validationErrors?: Array<{ message: string }> }).validationErrors ?? [];
                for (const e of errs) errors.add((e.message ?? "").trim());
                return { valid: false, errors };
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                return { valid: false, errors: new Set([message]) };
            }
        },
    };
}
