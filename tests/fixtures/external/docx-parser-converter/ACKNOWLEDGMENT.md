# Acknowledgment — docx-parser-converter (omer-go)

The DOCX fixtures in this directory are borrowed from
[omer-go/docx-parser-converter](https://github.com/omer-go/docx-parser-converter)'s
own test corpus.

- Upstream repo: <https://github.com/omer-go/docx-parser-converter>
- Upstream source paths: `fixtures/test_docx_files/` (13), `fixtures/tagged_tests/` (5), `docx_parser_converter_ts/tests/fixtures/` (1)
- License: MIT (Copyright (c) 2024 omer-go) — see `LICENSE` in this directory.

## `test_docx_files/` — demo DOCX

Curated demos that exercise common WordprocessingML structures: inline
formatting (`docx_inline_formatting_demo`), lists with various numbering/
styles (`docx_list_numbering_text_styling_demo`, `lists_demo`,
`docx_list_formatting_demo`), tables (`tables_demo`, `table_advanced_demo`),
formatting and run effects (`formatting_and_styles_demo`, `run_effects_demo`,
`underline_styles_examples`, `docx_formatting_demo_combinations_paragraphs_fonts`,
`fonts_and_sizes_demo`), paragraph controls (`paragraph_control_demo`),
and a broad `comprehensive_docx_demo`.

## `tagged_tests/` — targeted feature tests

Smaller per-feature fixtures: `list_tests`, `formatting_tests`,
`table_tests_v2`, `margin_tests`, `image_tests`.

## `minimal_for_test.docx`

A minimal valid DOCX scaffold used by the upstream TS test suite.

Expected validator outcomes are pinned in `tests/fixtures-all.manifest.json`
(regenerated when these were added).

Thanks to omer-go and the docx-parser-converter contributors for keeping their
fixtures public.
