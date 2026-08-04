# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.12.0] - 2026-08-04

### Added
- **PDF comment support in `add_comment()`** — PyMuPDF sticky-note text annotations with author name (appears as the annotation title). New `page_num` parameter. The same method now covers DOCX, XLSX, PPTX, and PDF.

## [3.11.3] - 2026-08-03

### Fixed
- **`_resolve_file()` detected file type from the on-disk path instead of the DB `filename`
  column** — files this tool itself creates via `_save_and_link()` are stored under their bare
  file UUID with no extension, so every function routed through `_resolve_file()` (`add_comment`,
  `document_stats`, and others sharing the same helper) returned "Unsupported format" for any
  file the tool had previously created — a common two-step flow ("create a doc, then comment on
  it"). Now falls back to the DB's `filename` column, which always carries the real extension,
  whenever path-based detection comes back `unknown`.
- **`add_comment()` on DOCX called `Document.add_comment()` with the wrong argument** — passed
  the comment text as the first positional arg (`runs`), which python-docx expects to be a `Run`
  or sequence of `Run`s to anchor the comment to, not a string. Every DOCX comment call raised
  `AttributeError: 'str' object has no attribute 'mark_comment_range'`. Now anchors to the target
  paragraph's runs (adding an empty run first if the paragraph has none) and passes the comment
  body via the `text=` keyword, matching the real API.
- **`requirements:` header was missing 8 of the tool's actual runtime dependencies**
  (`docx-revisions, lxml, PyMuPDF, Pillow, pytesseract, qrcode, google-api-python-client,
  google-auth`) — present in the code (track changes, PDF/OCR, QR codes, Google Drive upload)
  but never listed, so a fresh install without them already present would silently break those
  features on first use.

Both `_resolve_file` and DOCX `add_comment` were verified against real files (create → comment →
reopen and confirm the comment landed) on a live Open WebUI instance, not just import-tested.

## [3.11.2] - 2026-08-03

### Fixed
- **Duplicate PPTX comment ext URI on repeat `add_comment()` calls** — the slide `ext` with URI
  `{6950BFC3-D8DA-4A85-94F7-54DA5524770B}` and its `commentRel` child are now reused instead of
  appended on every call, and the slide `.rels` `Relationship` pointing at the same
  `commentModern{n}.xml` target is deduplicated — repeat calls no longer corrupt the slide XML.

### Changed
- **Modular `src/` package split** — `tool.py` was split into `src/constants.py` (paths +
  namespaces), `src/utils.py` (shared helpers), and `src/tool.py` (the `Tools` class). The root
  `tool.py` is now a thin wrapper importing `Tools` from `src/`.

## [3.11.1] - 2026-08-03

### Added
- **`add_comment()` for PPTX is now fully implemented** — writes modern PowerPoint comments
  via lxml (`ppt/comments/commentModern{n}.xml` + `ppt/commentsAuthors.xml`) with author names,
  correcting the premature v3.11.0 claim where the PPTX block was missing from the released file.
  The Excel comments and track-changes error handling from v3.11.0 remain unchanged.

## [3.9.2] - 2026-07-30

### Fixed
- **`create_file()` for DOCX crashed with `NameError: RGBColor not defined` when rendering
  links** — added `RGBColor` to the `docx.shared` import. (BUG-1)
- **`_format_text()` split sentences on abbreviations** like "Dr.", "e.g.", "U.S.A." — added
  a comprehensive abbreviation exclusion list with negative lookbehind regex. (BUG-2)
- **`_parse_inline_md()` link regex broke on URLs with parentheses** (e.g. Wikipedia) — the
  `.+?` pattern stopped at the first `)`. Replaced with a balanced-parentheses-aware regex.
  (BUG-3)
- **`_parse_inline_md()` link label extraction used fragile `split('](')`** — switched to
  storing the regex group `(label, url)` tuple directly. (BUG-4)

## [3.9.0] - 2026-07-28

### Fixed
- **`_save_and_link()` was missing a `__user__` parameter that 14 functions already called it
  with** (`create_odf`, `mail_merge`, `add_chart`, `add_watermark`, `edit_metadata`,
  `add_alt_text`, `export_to_markdown`, `version_file`, `add_speaker_notes`, `add_qr_code`,
  `add_data_validation`, `add_named_range`, `add_slide_transitions`, `export_to_html`) —
  every one of these crashed with `TypeError` on its very first call. One-line signature fix.
- **`self._read_xlsx`/`_read_xls`/`_read_docx`/`_read_pptx` were called from 13 functions but
  never defined** (`convert_format`, `preview_file`, `compare_documents`, `export_to_markdown`,
  `ai_summarize`, `document_stats`, `export_to_html`, `ai_analyze`, `smart_fill`,
  `grammar_check`, `translate_document`, `classify_document`, `version_diff`) — every call was
  a guaranteed `AttributeError`. Implemented all four, mirroring the already-working
  `_read_odf()` pattern.
- **`create_odf(format="odp")` was completely broken, template or not** — it tried to add
  `text:p`/`text:h` elements directly under `office:presentation`, which ODF's schema doesn't
  allow (every call raised `IllegalChild`). Rebuilt using proper `draw:page`/`draw:frame`/
  `draw:text-box` slide structure with a master page.
- `update_cells` and `add_content` (xlsx) could crash with `AttributeError: 'MergedCell' object
  attribute 'value' is read-only` when writing into a non-anchor cell of a merged range. Added
  the same guard `create_file`'s template mode already had.
- `modify_rows` (insert/delete) didn't shift existing merged-cell ranges, silently corrupting
  layout on any sheet with merged headers. Now re-anchors ranges cleanly above/below the
  affected rows, and warns (instead of silently corrupting) when a range straddles them.
- `tracked_change`, `merge_sheets`, `merge_pdfs`, `split_pdf`, `manage_revisions` resolved
  their source file via a broken inline lookup (`meta.get("path", file_id)` — `meta` never
  actually has a `"path"` key) instead of the already-correct `self._resolve_file()` helper
  used everywhere else. Switched all 5 to use it.
- `manage_revisions`'s `action="list"` referenced `RevisionParagraph` without importing it —
  `NameError` on every call. Added the missing import.
- `create_file` (docx) used `RGBColor` for markdown link styling without importing it —
  `NameError` on any content containing a `[link](url)`. Added the missing import.
- `add_comment` misused python-docx's comment API (passed comment text as the anchor argument;
  called a `Paragraph.add_comment` method that doesn't exist) — crashed on every call. Fixed to
  anchor the comment on the paragraph's runs per the real API.
- `add_pivot_table` never built a pivot table — it created an empty "Pivot" sheet and always
  reported success. Now computes a real cross-tab aggregation (sum/count/average/min/max) in
  Python (openpyxl can't create native Excel PivotTable objects).
- `generate_spreadsheet` parsed CSV by hand (only checked for tab/comma) — broke on quoted
  fields containing the delimiter and never recognized semicolon-delimited data, the standard
  Excel export format in PT/FR/DE locales. Switched to `csv.reader` with delimiter sniffing.
- `mail_merge` had no error handling, loaded the data workbook without `data_only=True` (so
  formula cells inserted literal formula text like `=SUM(F13:F17)` instead of the computed
  value), and only substituted `{{field}}` in paragraphs, never in tables. Fixed all three.
- `document_assembly` forwarded its `template_name` into `mail_merge`'s uploaded-file lookup
  instead of the saved-templates store (`save_template`/`use_template` actually use) — a saved
  template could never be found. Reimplemented to read from the correct store and generate one
  document per data row.
- `import_from_api` silently wrote zero rows (while still reporting "Imported N records" as
  success) when the API response was a list of non-dict items. Now wraps scalar items into a
  usable row instead of silently dropping them.
- `_xls_to_xlsx` never carried over merged-cell ranges from the source `.xls`, silently
  dropping merge formatting on every legacy-format conversion.
- `docx-revisions`, `lxml`, `odfpy`, `PyMuPDF`, `Pillow`, `pytesseract`, `qrcode`,
  `google-api-python-client`, `google-auth` were imported by working features but missing from
  `requirements.txt`.
- `batch_process` discarded the actual result of each per-file operation and always reported a
  hardcoded success line — a failed file looked identical to a succeeded one. Now reports each
  file's real outcome and a failure count.

### Added
- `template_file_id` support extended to `create_file`'s docx and pptx branches, and added to
  `create_odf` (odt/ods/odp) — same principle as the xlsx support shipped in v3.8.0: load the
  real template instead of building blank, so new content inherits its styles/theme. docx and
  pptx inherit the template's full style/theme by loading the real file; ODF style reuse is at
  the document/style level (including per-column cell styles for .ods) rather than a full
  per-cell copy, since odfpy's API is flatter than openpyxl/python-docx/python-pptx.
- `debug_errors` valve (default off): include the full Python traceback in error responses only
  when explicitly enabled, instead of always. Normal-operation error payloads are now much
  shorter across all ~19 functions that had this pattern.
- `read_file` now caps and clearly flags truncation for docx (`max_paragraphs`, default 200) and
  pptx (`max_slides`, default 50) — previously unbounded. The existing xlsx row cap now also
  reports `truncated`/`returned_rows`/`total_rows` explicitly instead of leaving it implicit.
- `compare_documents`'s diff output is now capped at 100 lines (with a "N more differences not
  shown" note) instead of unbounded.

### Changed
- Added `data_only=True` to workbook loads that read *displayed* values rather than formulas —
  the 4 new `_read_xxx` helpers, `file_search`, and (via a separate read-only load, so the
  formulas in the saved file aren't destroyed) `add_pivot_table`'s aggregation pass.
- Trimmed `create_file`'s `template_header_row`/`template_data_row` docstring, which repeated
  the same caveat three times and was ~3x longer than any other docstring in the file.
- `read_file`'s main JSON response no longer pretty-prints with `indent=2` (LLM-consumed, not
  human-read) — compact separators only. Left `export_to_json`-style file *content* (actually
  downloaded by the user) pretty-printed, since that's genuinely human-facing.

### Removed
- Deleted `build.py` and the fragmented `src/` module split (`00_header.py`, `01_helpers.py`
  through `13_advanced.py`). Two separate multi-file split schemes were attempted across this
  repo's history and both silently rotted out of sync with the real `tool.py` — `src/
  00_header.py` was ~700 lines stale (missing the entire v3.5–3.7.2 feature set), and `src/
  01_helpers.py` through `13_advanced.py` were fully-dead duplicates from an abandoned earlier
  scheme. `src/tool.py` is kept as a single-file mirror of the real `tool.py` (updated in the
  same commit whenever `tool.py` changes) — no fragmented split, no build step, no risk of
  the two drifting apart piece by piece the way the old scheme did.

## [3.8.0] - 2026-07-28

### Added
- `create_file` (.xlsx) now accepts an optional `template_file_id`: instead of always creating a
  blank workbook, it loads the referenced file and copies its fonts, fills, borders, number
  formats, merged cell ranges, and column widths onto the new file. The template file itself is
  never modified. Fixes new xlsx files losing all formatting when a user attaches a reference
  file and expects the output to match it.
- `template_header_row` / `template_data_row` (default 1 / 2) let you point at which template
  rows to copy style from, for templates where row 1/2 aren't a plain header+data table (e.g. a
  title bar or metadata row) — using the wrong row can otherwise carry over a misleading
  number_format (e.g. a date format landing on a plain number).
- `template_file_id` is not yet supported for .docx/.pptx or `create_odf` (.odt/.ods/.odp) —
  calling with a non-xlsx type now returns a clear error instead of silently ignoring it.

### Added
- Added support for VPS / server installations: 
Now tool use Open WebUI API for file storage (`/api/v1/files/{id}/content`) instead of local export directory + custom HTTP file server. 
- Added `base_url` Valve for overriding download link base URL
- Added support for download link generation behind CDN/Proxy. Proxy url can be passed via `X-Original-Host` header and will be used to generate download link
- Bump tool.py meta to 3.0.0 (since prev downloading logic was replaced with more webui-native approach)
- `requirements.txt` for dependency management
- `CHANGELOG.md` to track project history
- Expanded `.gitignore` with test file exceptions

### Changed
- Updated `.gitignore` to exclude generated Office files and compiled Python

## [3.4.0] - 2026-07-26

### Added
- **AI Summarize** — `ai_summarize()` extracts document text for LLM summarization
- **Speaker Notes** — `add_speaker_notes()` adds speaker notes to PowerPoint slides
- **Document Stats** — `document_stats()` shows word count, reading time, complexity
- **QR Codes** — `add_qr_code()` generates QR codes in DOCX/PPTX documents
- **Bulk Folder Ops** — `bulk_folder_ops()` lists, deletes, and shows stats on all uploads
- **File Search** — `file_search()` full-text search across all generated files
- **Data Validation** — `add_data_validation()` adds dropdown lists and validation rules to Excel
- **Named Ranges** — `add_named_range()` defines named ranges in Excel workbooks
- **Slide Transitions** — `add_slide_transitions()` adds fade/push/wipe/split transitions to PPTX
- **HTML Export** — `export_to_html()` exports any Office file to a styled HTML page

## [3.3.0] - 2026-07-26

### Added
- **Document Comparison** — `compare_documents()` shows differences between two files
- **Markdown Export** — `export_to_markdown()` converts any Office file to Markdown
- **URL Import** — `import_from_url()` fetches web pages and converts to Word documents
- **File Versioning** — `version_file()` saves timestamped copies before editing
- **Google Drive** — `upload_to_drive()` uploads files to Google Drive
- **OCR** — `ocr_extract()` extracts text from images in PDFs
- **i18n** — `translate_errors()` sets error message language (en, pt, es, fr, de)

## [3.2.0] - 2026-07-26

### Added
- **ODF Write** — `create_odf()` creates .odt, .ods, .odp files from scratch
- **Format Conversion** — `convert_format()` converts between all Office formats
- **Template System** — `save_template()`, `use_template()`, `list_templates()` for reusable templates
- **Scheduled Cleanup** — `schedule_cleanup()` for automatic file cleanup
- **Mail Merge** — `mail_merge()` generates personalized documents from template + data
- **Charts** — `add_chart()` adds bar, line, pie, scatter charts to Excel
- **Watermark** — `add_watermark()` adds diagonal watermarks to DOCX/PDF
- **Password Protection** — `protect_file()` encrypts Excel and Word files
- **File Preview** — `preview_file()` shows text preview before downloading
- **Metadata Editing** — `edit_metadata()` edits author, title, subject, keywords
- **Accessibility** — `check_accessibility()` and `add_alt_text()` for document accessibility
- **Progress Indicators** — `_progress()` emits status updates for long operations

### Changed
- Tool now has 34 functions (up from 18)

## [3.1.0] - 2026-07-25

### Added
- **LibreOffice ODF Read** — `.odt`, `.ods`, `.odp` read support via `odfpy`
- **File Cleanup** — `cleanup_files(days_old=30)` removes old generated files
- `odfpy` dependency

### Fixed
- `_get_owui_data_dir` NameError on import (moved before constants)
- `_read_odf` error handling for invalid files
- `_detect_type` for ODF formats

## [3.0.0] - 2026-07-25

### Added
- **Native File API** — files saved to uploads dir, served via `/api/v1/files/{id}/content`
- `base_url` Valve with auto-detection from headers/env
- `pydantic` dependency

### Changed
- **Breaking:** Removed `export_dir` and `file_server_url` Valves
- **Breaking:** Removed `_save_file_sync()` helper
- `file_server.py` no longer required

### Contributors
- @skorphil — PR #1

## [2.4.0] - 2026-07-24

### Added
- **Cross-platform support** — tool now works on Windows, Mac, and Linux without configuration:
  - Windows: `%APPDATA%\open-webui\data\`
  - Mac: `~/Library/Application Support/open-webui/data/`
  - Linux/Docker: `$OPEN_WEBUI_DATA_DIR/data/`
- Helper functions `_get_owui_data_dir()` and `_get_owui_uploads_dir()` that auto-detect the OS

### Removed
- All hardcoded Windows paths (`C:\Users\...`, `AppData\Roaming\`, `LOCALAPPDATA`)

## [2.3.0] - 2026-07-24

### Added
- **Text formatting** — `_format_text()` function that automatically normalizes all generated document text:
  - Replaces em dashes (—) with regular hyphens (-)
  - Applies sentence case (first letter of each sentence uppercase, rest lowercase)
  - Preserves acronyms (API, PDF, HTML, CSS, JSON, SQL, AI, UI, UX, URL, etc.)
  - Capitalizes first letter after each period
  - Preserves Excel formulas (values starting with `=`)
- Applied to all three generators: `generate_document`, `generate_slides`, `generate_spreadsheet`

## [2.2.0] - 2026-07-23

### Added
- `merge_pdfs` — Merge multiple PDFs into one using PyMuPDF
- `split_pdf` — Split PDF into parts by page count
- `tool_stats` — Dashboard showing tools, functions, models, and exports count
- `merge_sheets` — Merge XLSX files preserving styles
- `batch_process` — Apply operations to multiple files at once
- `auto_backup` — Timestamped database snapshot for safety
- Office Templates KB — CV Europass, Cover Letter PT, Invoice, Proposal

## [1.2.0] - 2026-07-22

### Added
- **Track Changes** — `tracked_change()` function for Word document redlines with custom author names
  - Replace mode: swap text while preserving OOXML `w:ins` / `w:del` elements
  - Insert mode: add new paragraphs as tracked insertions
  - Delete mode: mark paragraphs for deletion
- **Manage Revisions** — `manage_revisions()` function to list, accept, or reject all tracked changes
- `docx-revisions` library integration for standards-compliant track changes
- Track changes visibility documentation for Microsoft Word, LibreOffice, and Google Docs

### Changed
- Enhanced `tool.py` with Word-specific revision handling via OOXML manipulation

## [1.1.0] - 2026-07-20

### Added
- `replace_text()` function for find-and-replace across all supported formats
- `create_file()` function to generate new Office files from scratch
- Professional styling templates for created files
- Highlight, bold, and italic detection in DOCX reads

### Changed
- Improved `read_file()` to return structured JSON with formatting metadata
- Enhanced `add_content()` to preserve original file styles when appending

## [1.0.0] - 2026-07-18

### Added
- Initial release
- `read_file()` — Read .xlsx, .xls, .docx, .pptx files
- `add_content()` — Append content while preserving formatting
- Excel support via `openpyxl` and `xlrd`
- Word support via `python-docx`
- PowerPoint support via `python-pptx`
- HTTP file server (`file_server.py`) on port 9000
- Export directory management

## [0.1.0] - 2026-07-15

### Added
- Project scaffolding and initial prototype
- Basic Excel read/write functionality
