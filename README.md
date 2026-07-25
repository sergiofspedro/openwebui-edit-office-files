# Edit Office Files — Open WebUI Tool

Create, read, edit and export Office files (.docx, .xlsx, .xls, .pptx, .odt, .ods, .odp) directly from Open WebUI chats. Preserves original formatting and styles. Supports Word track changes (redlines) with custom author names.

**Cross-platform:** Works on Windows, Mac, Linux, and VPS/Docker — no configuration needed. **LibreOffice ODF** read support included.

## What's New

| Version | Feature |
|---|---|
| **v3.1.0** | LibreOffice ODF support (.odt, .ods, .odp) + File cleanup |
| **v3.0.0** | Native file API — no file server needed, works on VPS/Docker |
| **v2.4.0** | Cross-platform auto-detection (Windows, Mac, Linux) |
| **v2.3.0** | Automatic text formatting (sentence case, no em dashes) |

## Features

| # | Function | Formats | Description |
|---|---|---|---|
| 1 | `read_file` | .xlsx .xls .docx .pptx .odt .ods .odp | Read any Office file and return contents as structured JSON. Detects highlights, bold, italic in DOCX. LibreOffice ODF support. |
| 2 | `add_content` | .xlsx .xls .docx .pptx | Add new content while preserving ALL original formatting. CSV rows for Excel, text for Word, slides for PowerPoint. |
| 3 | `replace_text` | .xlsx .xls .docx .pptx | Find and replace text across the entire file preserving fonts, styles, and cell formatting. |
| 4 | `create_file` | .xlsx .docx .pptx | Create a brand new Office file from scratch with professional styling. |
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

### Text Formatting (v2.3.0)

All generated documents (DOCX, PPTX, XLSX) are automatically formatted with:
- **No em dashes** — replaced with regular hyphens
- **Sentence case** — only first letter of each sentence is uppercase
- **Acronym preservation** — API, PDF, HTML, CSS, JSON, SQL, etc. kept uppercase
- **Excel formulas preserved** — values starting with `=` are not modified

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
| .odt | Yes | No | No | LibreOffice Writer via odfpy |
| .ods | Yes | No | No | LibreOffice Calc via odfpy |
| .odp | Yes | No | No | LibreOffice Impress via odfpy |

## Installation

### Method 1: Open WebUI Community
Search for "Edit Office Files" in the Open WebUI Community tools.

### Method 2: Manual Install
1. Download `tool.py` from this repo
2. In Open WebUI: Workspace > Tools > New Tool
3. Paste the code and save
4. Install dependencies:
```bash
pip install openpyxl python-docx python-pptx xlrd docx-revisions odfpy pydantic
```

### Method 3: Batch Install
Use the Batch Install Plugins tool in Open WebUI pointing to this repo.

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

## Dependencies
- `openpyxl` — Excel .xlsx read/write
- `python-docx` — Word .docx read/write
- `python-pptx` — PowerPoint .pptx read/write
- `xlrd` — Legacy Excel .xls read
- `docx-revisions` — Word track changes (redlines)
- `odfpy` — LibreOffice ODF read (.odt, .ods, .odp)
- `lxml` — XML processing (dependency of docx-revisions)

## License
MIT

## Author
giofsp — [GitHub](https://github.com/sergiofspedro)
