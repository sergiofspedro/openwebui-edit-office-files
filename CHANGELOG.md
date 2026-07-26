# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
