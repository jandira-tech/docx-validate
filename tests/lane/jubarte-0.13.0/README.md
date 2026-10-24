# jubarte 0.13.0 lane specimens

Twenty-four redlines produced by jubarte-redlines 0.13.0 from the public
`arthrod/neurotic_docx_bench-fixtures` corpus (no private documents). They come
from the 2026-10-22 lane run recorded in
`handoffs/2026-10-22-jubarte-0.13.0-lane/HANDOFF.md`. GitHub issues #208–#263
report the docx-validate findings on these files, one issue per finding.

- `manifest.csv` gives each file's lane kind, its sha256, its bench path, and
  whether Microsoft Word opened and exported it during the lane run. All 24 were
  opened and exported. That is a historical observation bound to the listed
  hash, not a fresh Word verification.
- The files live outside `tests/fixtures/` because the `fixtures-all-*` suites
  pin every fixture there, and several of these redlines take tens of seconds
  to validate. `tests/lane-jubarte-0.13.0.integration.test.ts` exercises them.
- Do not edit these files. A changed hash invalidates the Word observation.
