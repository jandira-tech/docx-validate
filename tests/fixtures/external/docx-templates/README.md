# docx-templates corpus

Real-world Word/Office 365 templates copied from the
[docx-templates](https://github.com/guigrpa/docx-templates) test fixture set
(`src/__tests__/fixtures/`). docx-templates uses these as inputs for its own
template-rendering engine; from an OOXML perspective they are all
structurally valid documents (with two cataloged exceptions — see
`tests/fixtures-docx-templates.test.ts`).

The Office lockfile `~$fice365.docx` from the upstream tree is intentionally
omitted (it is not a real DOCX, just a 162-byte byte-stream Word writes
while a document is open).
