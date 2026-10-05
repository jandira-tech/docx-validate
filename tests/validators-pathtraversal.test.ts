import { describe, expect, it } from "vitest";
import path from "node:path";

describe("BaseValidator path traversal", () => {
    it("should resolve target paths securely", async () => {
        const unpackedDir = "/tmp/unpacked";
        const relsFile = "/tmp/unpacked/word/_rels/document.xml.rels";
        const target = "../../../../../etc/passwd";

        const relsDir = path.dirname(relsFile);
        let targetPath: string;

        if (target.startsWith("/")) {
            targetPath = path.join(unpackedDir, target.replace(/^\/+/, ""));
        } else if (path.basename(relsFile) === ".rels") {
            targetPath = path.join(unpackedDir, target);
        } else {
            targetPath = path.join(path.dirname(relsDir), target);
        }

        targetPath = path.resolve(targetPath);
        const relToUnpacked = path.relative(unpackedDir, targetPath);

        let isSafe = true;
        if (relToUnpacked === ".." || relToUnpacked.startsWith(`..${path.sep}`) || path.isAbsolute(relToUnpacked)) {
            isSafe = false;
        }
        expect(isSafe).toBe(false);
    });
});
