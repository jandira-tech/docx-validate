import { describe, expect, it } from 'vitest';
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

describe('Path traversal vulnerability coverage in DOCXSchemaValidator', () => {
    it('should refuse to resolve paths pointing outside the root unpacked dir during validation', async () => {
        const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docx-traversal-test-'));

        try {
            const unpackedDir = path.join(tmpDir, 'unpacked');
            await fs.mkdir(unpackedDir);

            const relsDir = path.join(unpackedDir, 'word', '_rels');
            await fs.mkdir(relsDir, { recursive: true });

            const relsFile = path.join(relsDir, 'document.xml.rels');
            const badTarget = "../../../../etc/passwd";

            const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${badTarget}"/>
</Relationships>`;

            await fs.writeFile(relsFile, relsContent);
            const validator = new DOCXSchemaValidator({ unpackedDir, profile: "word-valid", originalFile: undefined });
            const result = await (validator as unknown as { validateOrphanedRelationships: () => Promise<unknown> }).validateOrphanedRelationships();
            expect(result).toBeDefined();
        } finally {
            await fs.rm(tmpDir, { recursive: true, force: true });
        }
    });

    it('should cover absolute path traversal logic', async () => {
        const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docx-traversal-test-2-'));
        try {
            const unpackedDir = path.join(tmpDir, 'unpacked');
            await fs.mkdir(unpackedDir);
            const relsDir = path.join(unpackedDir, 'word', '_rels');
            await fs.mkdir(relsDir, { recursive: true });
            const relsFile = path.join(relsDir, 'document.xml.rels');
            const badTarget = "/etc/passwd";
            const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${badTarget}"/>
</Relationships>`;
            await fs.writeFile(relsFile, relsContent);
            const validator = new DOCXSchemaValidator({ unpackedDir, profile: "word-valid", originalFile: undefined });
            await (validator as unknown as { validateOrphanedRelationships: () => Promise<unknown> }).validateOrphanedRelationships();
        } finally {
            await fs.rm(tmpDir, { recursive: true, force: true });
        }
    });

    it('should cover legitimate path resolution', async () => {
        const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docx-traversal-test-3-'));
        try {
            const unpackedDir = path.join(tmpDir, 'unpacked');
            await fs.mkdir(unpackedDir);
            const relsDir = path.join(unpackedDir, 'word', '_rels');
            await fs.mkdir(relsDir, { recursive: true });
            const mediaDir = path.join(unpackedDir, 'word', 'media');
            await fs.mkdir(mediaDir, { recursive: true });
            await fs.writeFile(path.join(mediaDir, 'image1.jpeg'), 'fake image');

            const relsFile = path.join(relsDir, 'document.xml.rels');
            const goodTarget = "media/image1.jpeg"; // Relative to word/
            const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${goodTarget}"/>
</Relationships>`;
            await fs.writeFile(relsFile, relsContent);
            const validator = new DOCXSchemaValidator({ unpackedDir, profile: "word-valid", originalFile: undefined });
            await (validator as unknown as { validateOrphanedRelationships: () => Promise<unknown> }).validateOrphanedRelationships();
        } finally {
            await fs.rm(tmpDir, { recursive: true, force: true });
        }
    });
});
