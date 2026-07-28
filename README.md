# Edit Office Files — Open WebUI Tool

Create, read, edit and export Office files (.docx, .xlsx, .xls, .pptx, .odt, .ods, .odp) directly from Open WebUI chats. Preserves original formatting and styles. Supports Word track changes (redlines) with custom author names.

**Cross-platform:** Works on Windows, Mac, Linux, and VPS/Docker — no configuration needed. **LibreOffice ODF** read support included.

## What's New

| Version | Feature |
|---|---|
| **v3.9.0** | `template_file_id` now works for all 6 creatable formats (docx, pptx, odt, ods, odp — not just xlsx). Fixed ~19 real bugs, including 27 functions that were guaranteed to crash on first use (`_save_and_link`/`_read_xxx` issues), `create_odf(format="odp")` being completely broken, and merged-cell crashes. Added a `debug_errors` valve and response caps/truncation to keep normal-operation output shorter. Removed the unmaintained `src/` module split — `tool.py` is now the single source of truth. See CHANGELOG.md for the full list. |
| **v3.8.0** | `create_file` can now match an existing file's formatting: pass `template_file_id` (xlsx only for now) to reuse its fonts, fills, borders, number formats, merged cells, and column widths in the new file instead of a blank default style. Optional `template_header_row`/`template_data_row` for templates where row 1/2 aren't the real header/data rows. |
| **v3.7.2** | Bug fixes: Linux data directory detection fix (`~/.open-webui/data/`), improved DB path fallback — 71 functions total |
| **v3.7.1** | Bug fixes: link rendering (blue + underline) in DOCX, variable shadowing fix — 71 functions total |
| **v3.7.0** | Markdown rendering in DOCX (headings, bold, italic, code), heading case preservation, `raw_text` parameter for all generate/create functions — 71 functions total |
| **v3.6.0** | AI analysis, smart fill, grammar check, translation, classification, smart templates, pivot tables, SQL→Excel, PDF forms, data conversion, compliance, audit log, retention, scheduled reports, document assembly, conditional formatting, comments, version diff, webhooks, API import |
| **v3.5.0** | 18 visual improvements: emojis, cards, KPI dashboards, progress bars, timelines, pull quotes, comparison tables, step guides, status badges, visual separators, 10 color palettes, 4 typography presets |
| **v3.4.0** | AI summarize, speaker notes, document stats, QR codes, bulk ops, file search, data validation, named ranges, slide transitions, HTML export |
| **v3.3.0** | Document comparison, Markdown export, URL import, versioning, Google Drive, OCR, i18n |
| **v3.2.0** | ODF write (.odt/.ods/.odp), format conversion, templates, mail merge, charts, watermark, password protection, preview, metadata, accessibility |
| **v3.1.0** | LibreOffice ODF read + File cleanup |
| **v3.0.0** | Native file API — no file server needed, works on VPS/Docker |
| **v2.4.0** | Cross-platform auto-detection (Windows, Mac, Linux) |
| **v2.3.0** | Automatic text formatting (sentence case, no em dashes) |

## Features

| # | Function | Formats | Description |
|---|---|---|---|
| 1 | `read_file` | .xlsx .xls .docx .pptx .odt .ods .odp | Read any Office file and return contents as structured JSON. Detects highlights, bold, italic in DOCX. LibreOffice ODF support. |
| 2 | `add_content` | .xlsx .xls .docx .pptx | Add new content while preserving ALL original formatting. CSV rows for Excel, text for Word, slides for PowerPoint. |
| 3 | `replace_text` | .xlsx .xls .docx .pptx | Find and replace text across the entire file preserving fonts, styles, and cell formatting. |
| 4 | `create_file` | .xlsx .docx .pptx | Create a brand new Office file from scratch with professional styling. DOCX supports markdown rendering. Optional `template_file_id` matches an existing file's formatting (fonts/styles/theme; xlsx also copies per-column styles and merged cells) instead of using default styling. |
| 5 | `tracked_change` | .docx | Apply Word track changes (redlines) with custom author name. Supports replace, insert, and delete modes. |
| 6 | `manage_revisions` | .docx | List all tracked changes, accept all, or reject all revisions in a Word document. |
| 7 | `merge_pdfs` | .pdf | Merge multiple PDFs into one using PyMuPDF. |
| 8 | `split_pdf` | .pdf | Split PDF into parts by page count. |
| 9 | `merge_sheets` | .xlsx | Merge Excel files preserving styles. |
| 10 | `batch_process` | All | Apply operation to multiple files at once. |
| 11 | `auto_backup` | - | Timestamped database backup for safety. |
| 12 | `tool_stats` | - | Show tool usage dashboard with counts. |
| 13 | `generate_document` | .docx | Generate professional Word documents with cover page, TOC, callouts, signatures, headers/footers. |
| 14 | `generate_slides` | .pptx | Generate PowerPoint presentations with 13 layouts, 5 chart types, 6 themes. |
| 15 | `generate_spreadsheet` | .xlsx | Generate Excel workbooks with tables, formulas, conditional formatting, multi-sheet. |
| 16 | `cleanup_files` | - | Remove generated Office files older than N days from storage and database. Default 30 days. |
| 17 | `create_odf` | .odt .ods .odp | Create new LibreOffice/OpenOffice files from scratch. Optional `template_file_id` reuses an existing ODF file's styles (document/style-level, plus per-column cell styles for .ods). |
| 18 | `convert_format` | All | Convert between formats (docx↔odt, xlsx↔ods, pptx↔odp). |
| 19 | `save_template` | - | Save document templates with {placeholders} for reuse. |
| 20 | `use_template` | .docx | Generate a document from a saved template. |
| 21 | `list_templates` | - | List all saved templates. |
| 22 | `schedule_cleanup` | - | Schedule automatic file cleanup at intervals. |
| 23 | `mail_merge` | .docx | Generate personalized documents from template + CSV/Excel data. |
| 24 | `add_chart` | .xlsx | Add bar, line, pie, or scatter charts to Excel. |
| 25 | `add_watermark` | .docx .pdf | Add diagonal watermark (DRAFT, CONFIDENTIAL) to documents. |
| 26 | `protect_file` | .xlsx .docx | Password-protect Excel and Word files. |
| 27 | `preview_file` | All | Show text preview of any file before downloading. |
| 28 | `edit_metadata` | .xlsx .docx .pptx | Edit author, title, subject, keywords in document properties. |
| 29 | `check_accessibility` | .docx .pptx | Check for heading hierarchy, missing alt text, structure issues. |
| 30 | `add_alt_text` | .pptx | Add alt text to images in PowerPoint slides. |
| 31 | `compare_documents` | All | Compare two documents and show differences. |
| 32 | `export_to_markdown` | All | Export any Office file to Markdown format. |
| 33 | `import_from_url` | .docx | Fetch a web page and convert it to a Word document. |
| 34 | `version_file` | All | Save a timestamped version before editing. |
| 35 | `upload_to_drive` | All | Upload files to Google Drive (requires credentials). |
| 36 | `ocr_extract` | .pdf | Extract text from images in PDFs using OCR. |
| 37 | `translate_errors` | - | Set error message language (en, pt, es, fr, de). |
| 38 | `ai_summarize` | All | Extract document text for LLM summarization. |
| 39 | `add_speaker_notes` | .pptx | Add speaker notes to PowerPoint slides. |
| 40 | `document_stats` | All | Word count, reading time, complexity analysis. |
| 41 | `add_qr_code` | .docx .pptx | Generate QR codes in documents. |
| 42 | `bulk_folder_ops` | - | List, delete, stats on all files in uploads folder. |
| 43 | `file_search` | All | Full-text search across all generated files. |
| 44 | `add_data_validation` | .xlsx | Dropdown lists, numeric/date validation in Excel. |
| 45 | `add_named_range` | .xlsx | Define named ranges in Excel workbooks. |
| 46 | `add_slide_transitions` | .pptx | Add fade, push, wipe, split transitions to slides. |
| 47 | `export_to_html` | All | Export any Office file to a styled HTML page. |
| 48 | `generate_document` *(updated)* | .docx | Markdown rendering — headings, bold, italic, code now parse correctly |
| 49 | `generate_slides` *(updated)* | .pptx | Now accepts `raw_text=True` to skip auto-formatting |
| 50 | `generate_spreadsheet` *(updated)* | .xlsx | Now accepts `raw_text=True` to skip auto-formatting |
| 51 | `create_file` *(updated)* | .docx | Markdown parsing for headings, bullets, inline formatting |
| 52 | `create_odf` *(updated)* | .odt .ods .odp | Now accepts `raw_text=True` to skip auto-formatting |

**Total: 71 functions (52 documented core functions + 19 internal helpers)**

### LibreOffice ODF Support (v3.1.0)

Read LibreOffice and OpenOffice files natively:
- **.odt** — Writer documents
- **.ods** — Calc spreadsheets
- **.odp** — Impress presentations

Requires `odfpy` dependency. Graceful error handling for invalid files.

### File Cleanup (v3.1.0)

Manage storage with `cleanup_files(days_old=30)`:
- Removes generated Office files older than N days
- Cleans both disk and database
- Uses `source: "office-plugin"` metadata to identify generated files
- Example: `cleanup_files(days_old=7)` — removes files older than 1 week

### Markdown Rendering (v3.7.0)

`create_file("docx", ...)` and all `generate_*` functions now parse markdown in text content:

- **Headings** — `# H1`, `## H2`, `### H3` render as Word headings with **original capitalization preserved** (no more forced lower case)
- **Bold** — `**text**` renders as bold
- **Italic** — `*text*` renders as italic
- **Inline code** — `` `code` `` renders as code-styled text
- **Code blocks** — Triple-backtick blocks render as monospace paragraphs
- **Bullet lists** — `- item` / `* item` render as bulleted lists

### Formatting Control (v3.7.0)

All `generate_document`, `generate_slides`, `generate_spreadsheet`, `create_file`, and `create_odf` now accept:

- **`raw_text=True`** — skip all auto-formatting (sentence case, em dash replacement) and keep your text exactly as written
- Existing document edits (`add_content`, `replace_text`) always preserve original formatting

### Bug Fixes (v3.7.2)

- **Linux data directory detection** — Fixed path detection for `~/.open-webui/data/` on Linux environments
- **DB path fallback** — Improved database path resolution when `OPEN_WEBUI_DATA_DIR` is not set, ensuring correct fallback to default paths

### Bug Fixes (v3.7.1)

- **Link rendering** — Hyperlinks in DOCX files now render with proper blue color and underline styling
- **Variable shadowing** — Fixed variable shadowing conflict that could cause unexpected behavior in document processing

### Text Formatting (v2.3.0)

All generated documents (DOCX, PPTX, XLSX) are automatically formatted with:
- **No em dashes** — replaced with regular hyphens
- **Sentence case** — only first letter of each sentence is uppercase
- **Acronym preservation** — API, PDF, HTML, CSS, JSON, SQL, etc. kept uppercase
- **Excel formulas preserved** — values starting with `=` are not modified
- **Use `raw_text=True`** to bypass all automatic formatting

### Track Changes (v1.2.0)

Built on [docx-revisions](https://github.com/balalofernandez/docx-revisions) library. Writes standard OOXML `w:ins` / `w:del` elements — 100% compatible with Microsoft Word.

```python
# Replace text with track changes
tracked_change(file_id, change_type="replace", content="old_text|||new_text", author="Sergio Pedro")

# Insert new text as tracked change
tracked_change(file_id, change_type="insert", content="New paragraph text", author="Reviewer")

# Mark paragraph for deletion
tracked_change(file_id, change_type="delete", content="3", author="Editor")

# List all revisions
manage_revisions(file_id, action="list")

# Accept all changes
manage_revisions(file_id, action="accept_all")
```

**Track changes visibility by program:**

| Program | Track changes visible? |
|---|---|
| Microsoft Word | Yes — 100% native support |
| LibreOffice | Yes — good OOXML revision support |
| Google Docs | Partial — opens .docx but may drop metadata |

### Format Support

| Format | Read | Edit (preserve style) | Create | Notes |
|---|---|---|---|---|
| .xlsx | Yes | Yes | Yes | Full support via openpyxl |
| .xls | Yes | Yes (saves as .xlsx) | — | Legacy format via xlrd |
| .docx | Yes | Yes + Track Changes | Yes | Full support via python-docx + docx-revisions |
| .pptx | Yes | Yes | Yes | Full support via python-pptx |
| .doc | No | No | No | Suggest converting to .docx |
| .ppt | No | No | No | Suggest converting to .pptx |
| .odt | Yes | No | Yes | LibreOffice Writer via odfpy |
| .ods | Yes | No | Yes | LibreOffice Calc via odfpy |
| .odp | Yes | No | Yes | LibreOffice Impress via odfpy |

## Installation

### Method 1: Open WebUI Community
Search for "Edit Office Files" in the Open WebUI Community tools.

### Method 2: Manual Install
1. Download `tool.py` from this repo
2. In Open WebUI: Workspace > Tools > New Tool
3. Paste the code and save
4. Install dependencies (see [Dependencies](#dependencies) below):
```bash
pip install openpyxl python-docx python-pptx xlrd docx-revisions lxml odfpy PyMuPDF Pillow pytesseract qrcode google-api-python-client google-auth
```

### Method 3: Batch Install
Use the Batch Install Plugins tool in Open WebUI pointing to this repo.

### Development
`tool.py` at the repo root is the single source of truth. `src/tool.py` is kept as an exact
mirror (same content, updated in the same commit) for tooling that expects a `src/` layout —
it is not a separate build artifact and there is no build step.

## Cross-Platform & VPS/Docker

The tool auto-detects your operating system and works on all platforms:

| OS | Database path | Uploads path |
|---|---|---|
| **Windows** | `%APPDATA%\open-webui\data\webui.db` | `%APPDATA%\open-webui\data\uploads` |
| **Mac** | `~/Library/Application Support/open-webui/data/webui.db` | `~/Library/Application Support/open-webui/data/uploads` |
| **Linux / Docker** | `$OPEN_WEBUI_DATA_DIR/data/webui.db` | `$OPEN_WEBUI_DATA_DIR/data/uploads` |

### VPS & Server Deployments

Since v3.0.0, the tool uses Open WebUI's **native file API** — no separate file server needed. Generated files are saved to the uploads directory and served via `/api/v1/files/{id}/content`. This works out of the box on VPS, Docker, and behind reverse proxies.

For **Docker users**, no extra configuration is needed. The tool auto-detects the environment.

## Usage Examples

**Read a file:**
```
"Read this Excel file and show me the data"
"Show me what's in this Word document"
"What slides are in this presentation?"
```

**Add content:**
```
"Add these rows to the Excel keeping the same style:
Name,Age,City
Ana,30,Lisbon"
"Add this paragraph to the end of the Word document"
```

**Replace text:**
```
"Replace 'N/A' with 'Not Available' in this file"
"Change all '2025' to '2026' in the spreadsheet"
```

**Track changes (Word only):**
```
"Replace 'old contract' with 'new contract' in this Word doc as a tracked change, author=Sergio"
"Insert this clause as a redline, author=Legal Team"
"List all tracked changes in this document"
"Accept all revisions"
```

**Create new file:**
```
"Create an Excel with columns Name, Age, City and 5 rows of data"
"Create a PowerPoint with 3 slides about Q3 results"
```


## Base URL and Proxy Setup

When the tool saves generated files, it creates download links. The base URL for these links is resolved in this order:

1. **Valve setting** — Set `base_url` in the tool configuration in Open WebUI (e.g., `https://your-domain.com`)
2. **X-Original-Host header** — If running behind a reverse proxy (e.g., Nginx, Caddy), the tool reads the `X-Original-Host` header from the incoming request
3. **WEBUI_URL env var** — Falls back to the `WEBUI_URL` environment variable
4. **Default** — `http://localhost:3000`

### Reverse Proxy Configuration

If you run Open WebUI behind a reverse proxy, ensure the proxy passes the original host header:

**Nginx:**
```nginx
proxy_set_header X-Original-Host $host;
```

**Caddy:**
```
header_up X-Original-Host {host}
```

**Apache:**
```apache
ProxyPreserveHost On
```

Without this header, download links will use `localhost:3000` and won't work from other machines.

## Other Valves
- `debug_errors` (default `false`) — include the full Python traceback in error responses.
  Leave off for normal use; turn on when debugging a specific failure.
- `templates` / `cleanup_schedule` / `language` — see `save_template`/`schedule_cleanup`/
  `translate_errors` above.

## Dependencies
- `openpyxl` — Excel .xlsx read/write
- `python-docx` — Word .docx read/write
- `python-pptx` — PowerPoint .pptx read/write
- `xlrd` — Legacy Excel .xls read
- `docx-revisions` — Word track changes (redlines)
- `odfpy` — LibreOffice ODF read/write (.odt, .ods, .odp)
- `lxml` — XML processing (dependency of docx-revisions)
- `PyMuPDF` — PDF merge/split
- `Pillow` — image handling (OCR, QR codes, images in documents)
- `pytesseract` — OCR text extraction (requires the `tesseract` binary on the host)
- `qrcode` — QR code generation
- `google-api-python-client`, `google-auth` — Google Drive import

## License
MIT

## Author
giofsp — [GitHub](https://github.com/sergiofspedro)
