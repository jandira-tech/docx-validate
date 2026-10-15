/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { publicationPlan, verifyPackedFiles, verifyPackedPackage } from "../scripts/prerelease";

describe("release publication guards", () => {
    it("keeps prereleases off the latest dist-tag", () => {
        expect(publicationPlan("0.53.0-rc.0", "v0.53.0-rc.0", "prerelease").distTag).toBe("next");
        expect(publicationPlan("0.53.0", "v0.53.0", "stable").distTag).toBe("latest");
    });

    it.each(["0.53", "00.53.0", "0.53.0-rc.01", "0.53.0;echo bad", "v0.53.0"])("rejects invalid version %s", (version) => {
        expect(() => publicationPlan(version)).toThrow("Invalid release version");
    });

    it("rejects mismatched release tags and release status", () => {
        expect(() => publicationPlan("0.53.0-rc.0", "v0.52.0")).toThrow("does not match");
        expect(() => publicationPlan("0.53.0-rc.0", undefined, "stable")).toThrow("prerelease status");
        expect(() => publicationPlan("0.53.0", undefined, "prerelease")).toThrow("prerelease status");
    });

    it("requires the built entry, declarations, and bundled schemas", () => {
        const files = ["dist/index.mjs", "dist/index.d.mts", "dist/schemas/ISO-IEC29500-4_2016/wml.xsd"];
        expect(() => verifyPackedFiles(files)).not.toThrow();
        for (const file of files) expect(() => verifyPackedFiles(files.filter((entry) => entry !== file))).toThrow("Package is missing");
        expect(() => verifyPackedFiles([...files, "tests/fixture.docx"])).toThrow("unexpectedly contains");
    });

    it("checks package identity and contents with npm 11 and npm 12 metadata", () => {
        const pkg = {
            name: "docx-validate",
            version: "0.53.0-rc.1",
            files: ["dist/index.mjs", "dist/index.d.mts", "dist/schemas/ISO-IEC29500-4_2016/wml.xsd"].map((path) => ({ path })),
        };
        for (const metadata of [[pkg], { "docx-validate": pkg }]) {
            const json = JSON.stringify(metadata);
            expect(() => verifyPackedPackage(json, pkg.name, pkg.version)).not.toThrow();
            expect(() => verifyPackedPackage(json, "other-package", pkg.version)).toThrow("identity");
            expect(() => verifyPackedPackage(json, pkg.name, "0.52.0")).toThrow("identity");
        }
        expect(() => verifyPackedPackage(JSON.stringify([pkg, pkg]), pkg.name, pkg.version)).toThrow("identity");
        expect(() => verifyPackedPackage(JSON.stringify({ [pkg.name]: { ...pkg, files: [] } }), pkg.name, pkg.version)).toThrow(
            "Package is missing",
        );
    });
});
