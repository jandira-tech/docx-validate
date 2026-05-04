# Acknowledgment — mammoth.js

The DOCX fixtures in this directory are borrowed from
[mwilliamson/mammoth.js](https://github.com/mwilliamson/mammoth.js)'s own
test corpus.

- Upstream repo: <https://github.com/mwilliamson/mammoth.js>
- Upstream source path: `test/test-data/`
- License: BSD-2-Clause — see `LICENSE` in this directory.

## Files

- `strict-format.docx` — ISO OOXML Strict conformance class with a single
  body paragraph. Used by `tests/strict-namespace-paragraph-count.test.ts`
  to verify Strict-namespace detection and paragraph counting.
- `text-box.docx` — DOCX with a text box; the inner paragraphs must not be
  counted as body paragraphs. Used by `tests/textbox-paragraph-count.test.ts`.

Thanks to Michael Williamson and the mammoth.js contributors for keeping
their fixtures public.
