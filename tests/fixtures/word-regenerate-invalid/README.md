# Word-Regeneration Invalid Corpus

This folder is the tracked experimental corpus for DOCX files whose manifest
Word outcome is not `clean-open`.

- `original/` contains verbatim copies from `tests/fixtures/`, preserving the
  original relative paths.
- `regenerated/` is produced by `bun run test:fixtures:regenerate-word-invalid`.
- `regeneration-results.jsonl` records unpack, repair, pack, content-signature,
  and `word-valid` validator outcomes for each regenerated file.
- `word-probe-results.jsonl` records the Microsoft Word oracle result for the
  current regenerated corpus.

The pass criterion for this experiment is empirical: Microsoft Word opens the
regenerated file cleanly and the text-bearing Word content is preserved.

Current checkpoint:

- Regenerated DOCX files probed in Word: 21
- Word clean opens: 10
- Validator false positives: 0
- Validator false negatives: 1
