# Acknowledgment — docx2html (PolicyStat)

The DOCX fixtures and WordprocessingML template snippets in this directory are
borrowed from [PolicyStat/docx2html](https://github.com/PolicyStat/docx2html)'s
own test corpus.

- Upstream repo: <https://github.com/PolicyStat/docx2html>
- Upstream source paths: `docx2html/fixtures/` (`.docx`), `docx2html/tests/templates/` (`.xml`)
- License: BSD-3-Clause (Copyright (c) 2013, PolicyStat LLC.) — see `LICENSE` in this directory.

## Fixtures (`.docx`)

Real-world DOCX exercising structures a converter/validator must handle:
lists (`simple_lists`, `nested_lists`, `lists_with_styles`, `tables_in_lists`,
`list_in_table`, `upper_alpha_all_bold`), tables (`nested_tables`,
`table_col_row_span`, `nested_table_rowspan`), headings/headers (`headers`,
`split_header`, `headers_with_full_line_styles`, `fake_headings_by_length`,
`bigger_font_size_to_header`, `convert_p_to_h`, `list_to_header`, `has_title`),
images (`has_image`, `resized_image`, `has_missing_image`, `attachment_is_tiff`),
inline/formatting (`inline_tags`, `special_chars`, `greek_alphabet`,
`shift_enter`), and tracked changes (`track_changes_on`).

## Templates (`templates/*.xml`)

Minimal WordprocessingML fragment snippets (`p`, `r`, `t`, `tr`, `tc`, `table`,
`style`, `styles`, `sectPr`, `hyperlink`, `insert`, `pict`, `drawing`,
`smart_tag`, `base`) used upstream to assemble document XML for tests.

Expected validator outcomes for the `.docx` are pinned in
`tests/fixtures-all.manifest.json` (regenerated when these were added).

Thanks to PolicyStat and the docx2html contributors for keeping their fixtures public.
