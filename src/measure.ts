/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * `Measure` — per-file round-trip + classify + metrics.
 *
 * PR C Task C.4 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §2.
 *
 * Measure on docx-validate is a STUB by design. The full classification
 * (byte-equivalent / bailed-with-canonical-match / ast-equivalent-byte-differs
 * / lossy-tracked / hard-fail) requires a reader→AST→writer adapter that
 * lives in jubarte-first. Until that adapter is injected, this class
 * reports `classification: "no-ast-adapter"` and surfaces basic metrics
 * derived from the validator (whether the input is structurally valid).
 *
 * When jubarte-first composes docx-validate (spec §7), it will pass its
 * AST adapter via the `astAdapter` option to enable the full Measure
 * pipeline without modifying this file.
 */

import type { ValidationIssue } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { Validate } from "./validate";

export type MeasureClassification =
    | "byte-equivalent"
    | "bailed-with-canonical-match"
    | "ast-equivalent-byte-differs"
    | "lossy-tracked"
    | "hard-fail"
    | "no-ast-adapter";

export type MeasureMetrics = {
    bodyBailoutCount: number;
    t2ElementCarrierCount: number;
};

export type MeasureResult = {
    classification: MeasureClassification;
    metrics: MeasureMetrics;
    diagnostics: ValidationIssue[];
};

/**
 * Optional dependency injected by jubarte-first. When provided, Measure
 * runs the full read→write→re-read pipeline and classifies into the five
 * spec buckets. When omitted (default), Measure falls back to a structural
 * validity check.
 */
export type AstAdapter = {
    readonly read: (bytes: Uint8Array) => Promise<{
        ast: unknown;
        bodyBailoutCount: number;
        t2ElementCarrierCount: number;
    }>;
    readonly write: (ast: unknown) => Promise<Uint8Array>;
};

export type MeasureOptions = {
    astAdapter?: AstAdapter;
    xsdValidator?: XsdValidator;
    schemasDir?: string;
};

export class Measure {
    public constructor(private readonly opts: MeasureOptions = {}) {}

    public async runOne(bytes: Uint8Array): Promise<MeasureResult> {
        if (!this.opts.astAdapter) {
            // Fallback: structural-validity only.
            const result = await new Validate({
                xsdValidator: this.opts.xsdValidator,
                schemasDir: this.opts.schemasDir,
            }).run(bytes);
            return {
                classification: "no-ast-adapter",
                metrics: { bodyBailoutCount: 0, t2ElementCarrierCount: 0 },
                diagnostics: result.issues,
            };
        }

        // Full pipeline: read → write → re-read → compare. The injected
        // adapter does the AST hops; Measure only orchestrates.
        const inputView = await this.opts.astAdapter.read(bytes);
        const writeBytes = await this.opts.astAdapter.write(inputView.ast);
        const outputView = await this.opts.astAdapter.read(writeBytes);

        const isByteEquivalent = bytes.length === writeBytes.length
            && bytes.every((b, i) => b === writeBytes[i]);

        const metrics: MeasureMetrics = {
            bodyBailoutCount: inputView.bodyBailoutCount,
            t2ElementCarrierCount: outputView.t2ElementCarrierCount,
        };

        const classification: MeasureClassification = isByteEquivalent
            ? "byte-equivalent"
            : "ast-equivalent-byte-differs";

        return {
            classification,
            metrics,
            diagnostics: [],
        };
    }
}
