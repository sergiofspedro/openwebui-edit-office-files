# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.0.0] - 2026-08-15

Full bug-fix sweep of the whole file (~5900 lines, all 70 functions), triggered by a real
workflow failing three sessions in a row: placing excerpt-anchored review comments across a
156-page PDF. Root cause traced to two compounding bugs -- `add_comments` required a `page_num`
the model had no way to discover, and `export_to_markdown`/`read_file` extracted PDF text
perfectly then returned only a download link the model couldn't open. Fixing that led to a
full audit that surfaced ~60 real defects. Breaking change in one place: several functions that
previously reported false "success" now return an explicit error instead (see below) -- this is
the point of the release, not an oversight.

### Added
- `find_text(file_id, query, ...)` -- locate an excerpt's page/paragraph/cell before commenting,
  without reading the whole document. PDF matching shares the same word-level logic as
  `add_comment`/`add_comments`, so a page `find_text` reports is guaranteed to be the page a
  comment lands on.
- `page_num` is now optional for PDF in `add_comment`/`add_comments` -- when omitted, every page
  is searched in order and the comment lands wherever the excerpt is found; `match_index` then
  counts occurrences across the whole document instead of one page.
- `read_file` gained a `pdf` branch (with `page_start`/`page_end` paging) -- previously the only
  way to get PDF text into a response was the much slower `ocr_extract`.

### Fixed -- comment-anchoring correctness (the bugs behind the original failure)
- `_pdf_find_excerpt_quads` mis-indexed matches: PyMuPDF's `search_for` returns one rect per
  *line fragment*, not per occurrence, so a quote wrapping a line broke `match_index` counting
  and only half-highlighted the match. Rewritten on word-level matching (`page.get_text("words")`)
  so occurrences are grouped correctly and a wrapped quote is highlighted in full.
- DOCX batch comments with an unmatched excerpt no longer silently stack on paragraph 0 (the
  document title) -- they're now reported as a per-entry error instead.
- Fixed `read_file` raising `NameError: WD_COLOR_INDEX` on any DOCX containing a highlighted run
  (missing import) -- this broke reading exactly the kind of "comments in docx" file the failing
  workflow used.
- `_docx_split_run_at` no longer deletes inline images/line breaks when splitting a run to anchor
  a comment to part of its text.

### Fixed -- functions that reported false success
- `export_to_markdown`, `read_file`, `sql_to_spreadsheet`, `import_from_api`, `import_from_url`,
  `convert_data`, `add_pivot_table` now return real content inline (bounded, with a link to the
  full file), not just a download link the calling model can't open.
- `compare_documents` gained a PDF branch and now errors on unsupported formats instead of
  silently comparing two empty strings and reporting "0 differences".
- `document_assembly` now returns the generated documents' links (previously built and then
  discarded); capped at 200 rows.
- `add_alt_text` did nothing (`alt_text` isn't a real python-pptx property) and `check_accessibility`
  crashed reading the same non-existent property on any deck with a picture -- both now use the
  underlying XML element directly.
- `version_diff` could never find a version (it globbed a filename pattern that never exists on
  disk) -- now looks versions up by their real DB record.
- `add_pivot_table` silently wrote a second run to a sheet named "Pivot2" while still reporting
  "Pivot"; text-formatted numeric values were coerced to 0.0 in aggregates. Both fixed.
- `add_slide_transitions` raised `NameError` on every second call to the same file.
- `tracked_change` and `batch_process` reported success on entries that changed nothing.
- `generate_document` silently dropped a KPI block or pull quote if it was the last thing in the
  content, or followed by a blank line.
- `sql_to_spreadsheet` never committed -- any write statement was silently rolled back.
- `protect_file`'s xlsx sheet protection was set but never enabled (`protection.sheet` stayed
  `False`) -- every "protected" sheet opened fully editable.
- `fill_pdf_form` reported success with 0 fields filled; now returns the requested vs. actual
  field names when nothing matches.
- `replace_text` destroyed numbers and dates (substring-tested but whole-value-overwrote
  non-string cells) and inflated its reported count when a match spanned a DOCX/PPTX formatting
  boundary (no run actually changed). Both fixed; a cross-run miss is now reported, not hidden.
- `mail_merge` gained a try/except per row, now scans tables (not just top-level paragraphs),
  handles cross-run placeholders, stops treating `0`/`False` data values as empty, and is capped
  at 200 rows.

### Fixed -- crashes on ordinary input
- `add_content` (`NameError: Pt` on any content with inline code), `_read_docx` (`int()` crash on
  a `"Heading"` style with no number), `add_watermark`'s PDF branch (invalid PyMuPDF API usage --
  never worked), `add_data_validation` (crash on its own documented default call), `create_odf`
  (unreachable `ImportError` handler; `.odp` output produced an empty/broken presentation --
  rebuilt on real `draw:page` slides), `add_chart` (no format guard, plotted the label column as
  a data series instead of categories).

### Security
- `import_from_url`, `import_from_api`, `webhook_trigger` now reject non-http(s) schemes and
  URLs resolving to loopback/link-local/private-network addresses (blocks `file://` reads and
  SSRF to internal/cloud-metadata targets).
- `bulk_folder_ops` rejects a `pattern` containing `..` or a path separator, and double-checks
  the resolved path stays inside the uploads directory before `delete_old` removes anything.

### Fixed -- other
- Base64-encoded filenames (an internal storage detail) were never decoded back before being
  reused in 9 functions, so chained edits grew filenames like `cmVwb3J0Lnhsc3g_edited.xlsx`.
- CSV values with a leading zero ("02134") or an underscore ("1_000") are kept as text instead of
  being parsed as a different number.
- Markdown heading level in `generate_document`/`add_content` used `line.count('#')`, so
  `"# Bug #42"` became a level-2 heading instead of level-1.
- Markdown table separator rows using alignment syntax (`|:---|---:|`) rendered as a data row
  instead of being recognized as a separator.
- A progress-bar value outside 0-100 no longer produces a negative table-cell width.
- `read_file` kept only the last table on a PPTX slide with 2+ tables, and returned empty
  `"headers"` for every page after the first when paging with `row_start > 1`.
- `update_cells` used `re.match` (matched a prefix) instead of `fullmatch`, so `"A1:B10"` wrote
  one cell and reported success; missing `"value"` and unknown sheet names are now reported
  instead of silently mis-handled.
- `merge_sheets` disambiguates output sheet names that would otherwise collide after 15-char
  truncation (e.g. two files both truncating to `"Quarterly_Repor"`).
- `add_named_range` now quotes sheet names containing spaces and produces a fully absolute
  range (`$A$1:$B$10`), and validates the given name isn't a bare cell reference.
- `ocr_extract`'s page rendering (`get_pixmap`, ~100-300ms of blocking C code) now runs in the
  worker thread alongside Tesseract, instead of blocking the event loop before handing off.
- `_read_file_bytes` now caps input at 200MB instead of buffering an arbitrarily large file.
- `auto_backup` uses SQLite's own online backup API instead of a plain file copy, which could
  produce a torn/unrecoverable backup if it landed mid-write on the live database.

### Known remaining issues (not fixed in this pass, lower severity)
`_format_text` collapses newlines within a block and over-matches some short acronyms; some OOXML
elements are appended without strict schema ordering. None of these caused a confirmed live
failure; documented here so they're tracked rather than lost. (5 other items from this list were
fixed in v4.0.1 -- see below.)

## [4.0.1] - 2026-08-15

Fixes the 5 items v4.0.0 left on its "known remaining issues" list.

### Fixed
- **`_read_odf` (ODS)** now reads `table:number-columns-repeated` and `table:number-rows-repeated`
  -- a compacted ODS (LibreOffice collapses runs of identical/blank cells into one element with a
  repeat count) previously misaligned every column after the first repeated run. Repeats are
  capped at 20 to avoid exploding output on filler rows/columns encoding "the rest of the sheet
  is empty"; a fully-blank repeated row is skipped entirely rather than emitted as hundreds of
  blank lines.
- **`_read_odf` (ODP)** now numbers slides by real slide boundaries (`draw:page`, the same API
  already used correctly in `create_odf`'s writer) instead of incrementing once per paragraph --
  a single 5-bullet slide no longer reports as "Slide 1" through "Slide 5".
- **`generate_document`'s progress-bar detection** is now anchored to the whole line
  (`^Label:\s*NN%$`) instead of "contains a `:` and a `%` anywhere" -- prose like "Note: sales
  grew by 12% last quarter" is no longer misread as a progress bar.
- **`generate_document`'s timeline (colon/dash) detection** now validates `M/D`-style dates as a
  plausible month (1-12) and day (1-31), and rejects a matched description longer than 8 words --
  real timeline/agenda entries are short phrases, not full sentences, so "2024: was a great year
  for the company because it grew fast" is no longer misread as a timeline entry. This reduces
  (does not claim to eliminate -- full semantic disambiguation isn't achievable with regex alone)
  false positives from ordinary sentences that happen to start with a number or date-like token.
- **`modify_rows`** now shifts merged-cell ranges to match an insert/delete instead of leaving
  them pointing at the wrong cells (openpyxl's `insert_rows`/`delete_rows` move cell values only).
  Formula references are still not rewritten (would need a real formula parser, out of scope for
  this fix) -- the response now includes an explicit warning when the sheet contains formulas,
  so the risk is visible instead of silent.
- **`manage_revisions(action="list")`** now uses `docx_revisions`'s `RevisionDocument.all_paragraphs`
  (confirmed via the installed package's source) instead of `.paragraphs`, which explicitly
  excludes table content by design -- tracked changes inside a table cell were previously missing
  from the listing entirely. Also removed a redundant re-wrap (`all_paragraphs` already returns
  ready-to-use objects) and now reports a `paragraphs_with_errors` count instead of silently
  swallowing per-paragraph failures.

## [4.0.0] - 2026-08-15

Full bug-fix sweep of the whole file (~5900 lines, all 70 functions), triggered by a real

Full audit pass across all 70 functions for the same class of bug as 3.15.6: docstrings that
don't accurately convey real behavior, causing a calling model to misuse a tool, avoid it, or
trust a wrong result.

### Fixed (real logic bugs)
- **`smart_template`** crashed with `NameError` on the "no matching template" branch (referenced
  an undefined variable). Now returns a proper guidance message instead.
- **`compliance_check`** and **`check_accessibility`** silently returned a false-positive "passed,
  no issues found" for unsupported formats instead of running no checks. Both now return an
  explicit "not supported for this format" error.
- **`convert_data`** wrote unconverted raw text into a file under the wrong extension for
  unsupported source/target pairs (e.g. xml→csv), reporting success. Now returns an explicit
  unsupported-pair error instead of writing garbage output.
- **`add_slide_transitions`**'s documented `"random"` transition type had no code branch and
  silently produced no transition. Now picks a real transition (fade/push/wipe/split) at random.

### Fixed (missing docstrings — functions were invisible to the calling model)
- Added real docstrings to `merge_sheets`, `batch_process`, `auto_backup`, `merge_pdfs`,
  `split_pdf`, `tool_stats`.

### Changed (docstring accuracy — no behavior change)
- Corrected or clarified docstrings for 24 functions where wording didn't match real behavior:
  `read_file` (row_start/row_end are xlsx/xls-only), `add_content`/`create_file` (pptx uses
  `---`-separated blocks, not one slide per line), `protect_file` (honest structural edit-lock,
  not password encryption — no msoffcrypto-tool involved), `generate_document` (documented all 9
  auto-detected rich-content trigger patterns), `tracked_change` (DOCX-only, insert always
  appends at the end), `convert_format` (cross-category conversions are lossy), `mail_merge`
  (`{{field}}` placeholder syntax), `add_watermark` (DOCX branch is a centered banner, not
  diagonal), `ocr_extract` (PDF-only), `ai_summarize`/`ai_analyze`/`grammar_check`/
  `translate_document`/`classify_document` (these return raw text + a prompt for the calling LLM,
  not a finished result themselves), `add_pivot_table` (no-`rows_field` discovery mode),
  `version_diff` (heuristic line-count diff, not positional; PDF unsupported), `import_from_api`
  (added Args section, `data_path` dot-notation syntax), `document_assembly` (added Args section;
  non-xlsx data files fall back to a raw CSV-decode attempt), `use_template` (placeholders come
  from `**kwargs`), `import_from_url` (crude tag-stripping extraction, 50k-char cap),
  `bulk_folder_ops` (`delete_old` is permanent and irreversible), `file_search` (`pdf` filter
  doesn't actually work), `sql_to_spreadsheet` (flagged that `query` runs arbitrary SQL, not just
  SELECT, against the app's own database).

## [3.15.6] - 2026-08-14

### Changed
- **`export_to_markdown`'s docstring didn't mention its PDF support at all.** For PDF it already
  does fast, direct text extraction per page (via PyMuPDF, no OCR) with each page labeled
  "--- Page N ---" and sparse/likely-scanned pages flagged — exactly what's needed to confirm
  real page numbers before calling `add_comment`/`add_comments`. Because the docstring only said
  "Export any Office file to Markdown format," a calling model had no way to discover this and
  would reach for `ocr_extract` instead — a full, slow OCR pass over every page (156 in one real
  case) to answer a question `export_to_markdown` could answer in a single fast call. Docstring
  now describes the PDF behavior and explicitly recommends it over `ocr_extract` except for pages
  actually flagged as scanned images. No logic changed.

## [3.15.5] - 2026-08-11

### Changed
- **Strengthened the tool-selection wording for `add_comment` vs. `add_comments`.** `add_comments`
  (batch, one call = one output file with all comments) already existed and worked correctly, but
  a user reported the calling model chose to call `add_comment` 16 times instead — producing 16
  separate single-comment files — when asked to add 16 comments in one message. Root cause: the
  guidance to use `add_comments` for 2+ comments was on the 3rd line of `add_comment`'s docstring,
  not the first, and the plugin's top-level `description:` (the most global text Open WebUI shows a
  model) didn't mention it at all. Moved the warning to the first line of both `add_comment`'s and
  `add_comments`' docstrings, and added a short mention to the plugin `description:` metadata. No
  logic changed -- `add_comments` already did the right thing; this only makes it more likely a
  tool-selecting model picks it for multi-comment requests.

## [3.15.4] - 2026-08-11

### Fixed
- **`add_comments`'s docstring said `page_num` (PDF) was only required "if excerpt is
  omitted or not found," but the code always required it** — PDF excerpt search is scoped
  to a single page (unlike DOCX, which scans the whole document), so `page_num` was never
  actually optional. Corrected the docstring for both `page_num` (PDF, always required) and
  `paragraph_index` (DOCX, genuinely optional/fallback-only) to match actual behavior.
  (Flagged by Sourcery review on #8.)
- Typo in README: "mis-placing" -> "misplacing". (Flagged by Sourcery review on #8.)

### Changed
- Consolidated the duplicated quote/dash replacement map in `_normalize_match_text` and
  `_normalize_chars` into a single `_QUOTE_DASH_REPLACEMENTS` dict, with `_normalize_chars`
  as the core 1:1 (offset-preserving) primitive and `_normalize_match_text` as a thin
  wrapper that additionally collapses whitespace. No behavior change. (Flagged by Sourcery
  review on #8.)

## [3.15.3] - 2026-08-11

### Added
- **`add_comment`/`add_comments` accept an optional `excerpt` (+ `match_index`) to anchor a
  comment to exact quoted text instead of a fixed page number (PDF) or whole paragraph
  (DOCX).** PDF: highlights the matched text and attaches the comment to that highlight
  (falls back to the previous fixed-position sticky note, with a warning, if the excerpt
  isn't found on the page). DOCX: splits the matching run so the comment covers just the
  quoted span — or the minimal set of overlapping runs if the quote crosses a formatting
  boundary — falling back to `paragraph_index` with a warning if not found. Handles `"..."`
  inside excerpts (marks text omitted between two quoted spans; matches the text before it
  and reports a partial match) and curly vs. straight quote differences between the excerpt
  and the source document. All new parameters are optional and default to the prior
  behavior, so existing calls are unaffected.
- **`add_comments` (batch) now supports DOCX**, not just PDF — opens the document once,
  applies every comment (each entry may use `excerpt` or `paragraph_index`), and saves once,
  same as the existing PDF batch path.

## [3.15.2] - 2026-08-07

### Fixed
- **`add_content()` on `.xlsx`/`.xls` inserted new rows after a large gap of empty rows instead
  of right after the existing data.** It used `ws.max_row` (openpyxl's absolute last row in the
  sheet, which counts trailing rows that only carry formatting or blank cells) as the insertion
  point. On a sheet with such trailing rows, new content landed dozens or hundreds of rows below
  the real data instead of immediately after it. Added a `_last_populated_row(ws)` helper that
  scans from the bottom of the sheet for the last row with an actual non-empty value, and used it
  for both the insertion row and the reference-style row (font/fill/border/alignment copied from
  the last real row, not a blank one) in both the xlsx and xls branches.

## [3.15.1] - 2026-08-04

### Fixed
- **Chaining `add_comment()` calls (passing one call's returned `file_id` into the next, the
  documented way to accumulate comments one at a time) corrupted the filename on every hop** —
  `.pdf` growing to `.pdf.pdf.pdf...` and (before an intermediate fix in this same release) the
  name itself getting re-base64-encoded each time. Root cause: `_resolve_file()`'s fallback path
  (used for files this tool creates, which are stored on disk under a bare UUID) read the
  DB-stored filename directly without decoding it first — `_encode_filename()` base64-encodes
  the *whole* name including its extension, so passing that encoded string straight back into
  `_save_and_link()` (which encodes again) compounded on every chained call.
  `_decode_filename()` existed for exactly this but was never actually called anywhere in the
  file until now. Fixing its call site surfaced a second bug in `_decode_filename()` itself: it
  re-appended the extension on top of the already-decoded name (which already contains it),
  duplicating it once per call. Both fixed together. Verified with a real 3-hop chained
  `add_comment()` test — filename now stays `chain_test.pdf` on every hop, and page targeting
  (checked in the same investigation, prompted by a report that pages weren't landing correctly
  under real usage) was confirmed already correct — not reproducible, likely a
  transcription/context-confusion issue in an extremely long chat rather than a tool bug.

## [3.15.0] - 2026-08-04

Full bug/dead-feature audit of the file, prompted by the `_format_text` crash flagged in
3.14.2. All findings verified with real calls, not just read-and-assumed.

### Fixed
- **`_format_text()` crashed on every call with `raw_text=False` (the default)** — its
  sentence-splitting regex used a variable-length lookbehind, which Python's `re` rejects.
  Affected `create_file`, `add_content`, `generate_spreadsheet`, `create_odf`, and any other
  caller not explicitly passing `raw_text=True`. Rewritten to split on a fixed-width boundary
  and merge fragments back together in plain Python when the preceding fragment ends with a
  mini-abbreviation or a word with an internal period — same intended behavior, valid regex.
- **`manage_revisions(file_id, "list")` always returned zero revisions** — used
  `RevisionParagraph` without importing it (only `RevisionDocument` was imported), so every
  paragraph raised a `NameError` that was silently swallowed by an inner `except`. Import fixed.
- **~20 more functions never passed `__user__` to `_save_and_link()`** (same bug fixed for
  `add_comment` in 3.14.0) — `add_content`, `replace_text`, `update_cells`, `modify_rows`,
  `protect_file`, `create_file`, `generate_document`, `generate_slides`, `generate_spreadsheet`,
  `tracked_change`, `merge_sheets`, `merge_pdfs`, `split_pdf`, `manage_revisions`,
  `add_pivot_table`, `sql_to_spreadsheet`, `fill_pdf_form`, `convert_data`,
  `conditional_format`, `import_from_api`. Every file these created was unowned in the database.
- **~8 more functions didn't enforce the output file extension** (same bug fixed for
  `create_file` in 3.14.2) — a caller-supplied `output_filename` without the right suffix
  produced `application/octet-stream` files instead of the correct Office/PDF content type.
- **New opposite bug: `create_odf`, `sql_to_spreadsheet`, `import_from_api` doubled the
  extension** if the caller's `output_filename` already included it (e.g. `results.xlsx` →
  `results.xlsx.xlsx`). Both this and the missing-extension bug above are now fixed by one
  shared `_ensure_ext()` helper used everywhere output filenames are computed.

### Changed
- **`translate_errors()` now actually does something.** Previously it stored a language
  preference that nothing ever read — every error message stayed English regardless. The three
  shared, high-traffic messages ("File not found", "Could not save file", "Unsupported format" —
  which together appear ~65 times across the file) now route through a `_err()` lookup that
  respects the stored language. Function-specific error text (anything with a filename or
  reason embedded) is unchanged/still English — full i18n of every bespoke message was out of
  scope for what's meant to be a small convenience.
- **`schedule_cleanup()` / `retention_policy(policy="set", ...)` no longer claim automatic
  background execution.** There is no scheduler in this plugin (it's invoked per chat request,
  not a persistent process) — these functions only ever stored a policy that had to be applied
  manually via `cleanup_files()` / `retention_policy(policy="apply")`. The return messages now
  say so plainly instead of implying it happens on its own.

## [3.14.2] - 2026-08-04

### Fixed
- **`create_file()` didn't enforce the file extension on `output_filename`** — if the caller
  passed a name without the correct suffix (e.g. `"Review_Comments"` instead of
  `"Review_Comments.docx"`), the saved file kept the bare name, so `_save_and_link()`'s
  extension-based content-type lookup fell through to `application/octet-stream` and browsers/OS
  had no way to know it was a Word/Excel/PowerPoint file. Now appends the correct `.{ftype}`
  extension automatically when missing. Verified with a real call using an extension-less
  filename, confirming the saved record's filename and content-type are now correct.

### Known issue (not fixed in this release)
- `_format_text()`'s sentence-splitting regex uses a variable-length lookbehind
  (`(?<!\b(?:Dr|Mrs|Prof|...))`), which Python's `re` module rejects
  (`look-behind requires fixed-width pattern`) — found while testing this release.
  **`create_file()` and any other function routing text through `_format_text()` with
  `raw_text=False` (the default) currently crashes for docx output.** `raw_text=True` bypasses
  it as a workaround. Needs its own fix (fixed-width alternation, or a non-lookbehind rewrite of
  the sentence-boundary check).

## [3.14.1] - 2026-08-04

### Fixed
- **Every file this tool ever created was unreachable via Open WebUI's own `/content` download
  endpoint** — `_save_and_link()` wrote a plain preview string (e.g.
  `"[Word document: 45637 bytes]"`) directly into the `file` table's `data` column, which
  SQLAlchemy treats as a JSON type and auto-deserializes on every read. A non-JSON string there
  throws `JSONDecodeError`, which Open WebUI's `Files.get_file_by_id()` silently swallows
  (`except Exception: return None`) — so every download link this tool ever produced 404'd with
  a generic "not found" even though the file existed fine on disk the whole time. Root-caused by
  reproducing the exact failing request with a real admin JWT and tracing the swallowed
  exception directly, not guessed. Fixed by JSON-encoding the preview
  (`json.dumps({"content_preview": preview})`), matching the shape Open WebUI's own native
  upload code uses. A one-time repair migration was run against the live database to fix all
  pre-existing affected records (44 of 122 existing files needed repair; verified two
  specifically-reported broken links now return `200`).

## [3.14.0] - 2026-08-04

### Added
- **`add_comments()` — batch PDF comments in one call.** `add_comment()` always starts from the
  target file's *original* content and saves an independent copy, so calling it repeatedly
  (e.g. 28 times for a 28-comment review) produced 28 separate single-comment PDFs instead of
  one PDF with all comments — the caller would need to manually chain each call's returned
  `file_id` into the next to accumulate, which the docstring never explained. `add_comments()`
  takes a list of `{page_num, text, author}` entries, opens the PDF once, applies every comment,
  and saves once. Verified with a real 3-comment call: single output file, all 3 annotations
  present with correct page/text/author.

### Fixed
- **`add_comment()` never passed `__user__` to `_save_and_link()`** in any of its four format
  branches (DOCX/XLSX/PPTX/PDF), despite the function signature already accepting it —
  `_save_and_link()` defaults `user_id` to `""` when `__user__` is missing, so every file
  `add_comment()` created was silently unowned in the database. Same class of bug fixed for 14
  other functions in v3.11.x; this one was missed. Fixed all four branches; verified via a real
  call that the resulting file record's `user_id` is correctly populated.

## [3.13.0] - 2026-08-04

### Fixed
- **`ocr_extract()` blocked the entire server on multi-page PDFs** — looped over every page
  synchronously calling `pytesseract.image_to_string()` (a blocking subprocess call) inside an
  `async def`, with no page cap. Since Open WebUI runs a single-worker event loop, one OCR call
  on a large PDF could freeze the server for *all* users for many minutes — confirmed as the root
  cause of repeated live freezes via a `py-spy` stack dump taken mid-incident. Now runs each
  page's OCR in a background executor thread (`loop.run_in_executor`) so the event loop stays
  responsive, adds a `max_pages` cap (default 25, reports pages processed vs. skipped), and a
  60s per-page timeout as a second safety net. Verified with a real concurrency test: OCR ran for
  112s while a concurrent unrelated task completed in ~1s, unaffected.

### Added
- **Native PDF support in `export_to_markdown()`** — previously had no PDF branch at all, so the
  only way to get text out of a PDF via this tool was `ocr_extract()`, even for PDFs with a real
  embedded text layer (e.g. InDesign/Distiller-produced files), where OCR is unnecessary and far
  slower. Now extracts text directly via PyMuPDF's `page.get_text()` — no subprocess calls, near
  -instant even on large documents. Pages with little or no extractable text are flagged in the
  output as likely scanned images, with a pointer to `ocr_extract` for those specific pages.

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
