# word-invalid-fixtures — files real Microsoft Word refuses to open

A **Word-labeled regression corpus**: every `.docx` here was opened in *real*
Microsoft Word (macOS) by `scripts/word_oracle.py` and classified by the dialog
Word showed. These are the ground-truth negatives for the "word-valid"
openability bar — if a conversion/redline pipeline produces a file like one of
these, Mr. Smith's Word will reject it.

Verdicts here were **re-confirmed against real Word** (batch probe + per-file
re-probe of every indeterminate result; all TIMEOUTs resolved). The directory a
file sits in IS its current real-Word verdict — the folder reflects reality
exactly.

## Layout (by Word's verdict)

| dir | meaning | count |
|-----|---------|-------|
| `OPEN_ERROR/` | "Word experienced an error trying to open the file." (Word cannot load it at all) | 31 |
| `UNREADABLE_CONTENT/` | "Word found unreadable content in X. Do you want to recover…" (repair prompt) | 43 |
| `PASSWORD/` | Word prompts for a password (encrypted package, key required to open) | 2 |

Total: **76** files. None open cleanly in real Word.

> Note on encrypted packages: `encrypted-advanced-text.docx` triggers a real
> password prompt → `PASSWORD/`. `encrypted-hello.docx` uses an encryption type
> Word can't handle ("the encryption type used is not available") → it's a hard
> `OPEN_ERROR/`, not a password prompt.

Filenames are the **origin path flattened** (`/`→`__`) so provenance is visible
and there are no collisions, e.g.
`UNREADABLE_CONTENT/tests__fixtures__critic-variations__docx__13-table-with-changes.docx`.

- `manifest.jsonl` — one row per file: `{file, verdict, origin, duration_s, word_dialog}` (`word_dialog` = the exact Word dialog text captured during the re-confirm probe).
- `probe-reconfirm-summary.json` / `probe-reconfirm-results.jsonl` — the full re-confirm probe run (verdict + dialog messages per file) for reproducibility.

## How it was generated / re-confirmed

```bash
# real-Word open-and-observe over the corpus (one Word driver at a time)
python3 scripts/word_oracle.py probe word-invalid-fixtures --glob '*.docx' --out /tmp/reprobe
# any TIMEOUT (a sandbox-staging race, not a real verdict) is re-probed individually
python3 scripts/word_oracle.py probe <that-file>.docx --out /tmp/reprobe/single
```

`word_oracle.py` stages each file inside Word's sandbox container (so the flaky
macOS grant panel never appears), opens it, and reads Word's UI via the
Accessibility API — see `docs/word-validity-oracle.md` for the full tool +
hard-won macOS-Word automation notes.

## Defect classes found (why Word rejects these)

- **Misplaced content-model elements** (most `OPEN_ERROR` + table `UNREADABLE`):
  a `<w:r>` / `<w:rPr>` / `<w:sectPr>` / drawing `<wp:extent>` or a table-property
  element (`<w:tblBorders>`, `<w:tblCellSpacing>`) sitting where the schema
  forbids it. `docx-validate --profile word-valid` flags these as errors
  (blocking allowlist in `WORD_BLOCKING_MISPLACED_LOCALS`).
- **Encrypted packages** (`PASSWORD/` + the unsupported-encryption `OPEN_ERROR`).

NB — signals that look like defects but Word TOLERATES (intentionally NOT
word-valid errors, to keep the profile false-positive-free):

- **Duplicate `<w:pPr>`** in one paragraph — Word silently merges/ignores the
  second; `pPr` is therefore NOT in the blocking allowlist.
- **Broken `media/` relationships** (image Target absent) — Word shows a
  missing-image placeholder and opens cleanly.
- **Orphan comment-range marker** — `comments.unmatched-comment-marker_from_html.docx`
  carries the exact same orphan and opens cleanly.

## Using it as a regression bench

The goal: `docx-validate --profile word-valid` should ERROR on `OPEN_ERROR/` and
`UNREADABLE_CONTENT/` (Word-faithful), and stay clean on Word-clean files. The
sibling docx-validate repo locks this with `tests/word-valid-openability.test.ts`.
Verified over the full 1380-file labeled probe: **`CLEAN-FP 0/1304`,
`BAD-COVERED 61/70`**. Re-label after Word/pipeline changes by re-running the
probe above.
