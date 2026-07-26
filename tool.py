"""
title: Edit Office Files
author: giofsp
author_url: https://github.com/sergiofspedro
description: Unified tool to read, edit, and create Office files (.xlsx, .xls, .docx, .pptx) preserving original formatting and styles. Detects highlights, bold, italic formatting. Detects legacy .doc and .ppt. Note: Track changes are not supported.
version: 3.0.0
requirements: openpyxl, python-docx, python-pptx, xlrd, odfpy
"""

import io
import json
import os
import platform
import re
import sqlite3
import sys
import traceback
from copy import copy
from typing import Any, Dict, Optional

from pydantic import BaseModel, Field


def _get_owui_data_dir() -> str:
    """Return the Open WebUI data directory for the current OS."""
    data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
    if data_dir:
        return data_dir
    home = os.path.expanduser("~")
    if platform.system() == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "open-webui", "data")
    else:
        return os.path.join(home, "Library", "Application Support", "open-webui", "data")

def _get_owui_uploads_dir() -> str:
    """Return the Open WebUI uploads directory for the current OS."""
    data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
    if data_dir:
        return os.path.join(data_dir, "data", "uploads")
    home = os.path.expanduser("~")
    if platform.system() == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "open-webui", "data", "uploads")
    else:
        return os.path.join(home, "Library", "Application Support", "open-webui", "data", "uploads")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
_DB_PATH = os.path.join(
    os.environ.get("OPEN_WEBUI_DATA_DIR", ""),
    "data", "webui.db",
)
if not os.path.isfile(_DB_PATH):
    _DB_PATH = os.path.join(_get_owui_data_dir(), "webui.db")

_UPLOAD_DIR = os.path.join(
    os.environ.get("OPEN_WEBUI_DATA_DIR", ""),
    "data", "uploads",
)
if not os.path.isdir(_UPLOAD_DIR):
    _UPLOAD_DIR = _get_owui_uploads_dir()

_EXPORT_DIR = os.environ.get("OWUI_EXPORTS_DIR", os.path.join(os.path.expanduser("~"), "open-webui", "exports"))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _resolve_file_path(file_id: str) -> Optional[str]:
    """Resolve an Open WebUI file UUID to an absolute disk path."""
    try:
        conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
        row = conn.execute(
            "SELECT path FROM file WHERE id = ?", (file_id,)
        ).fetchone()
        conn.close()
    except Exception as exc:
        print(f"[office] DB lookup failed for {file_id}: {exc}", file=sys.stderr)
        return None

    if not row or not row[0]:
        # Fallback: try by filename
        try:
            conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
            row = conn.execute(
                "SELECT path FROM file WHERE filename LIKE ?",
                (f"%{file_id}%",),
            ).fetchone()
            conn.close()
        except Exception:
            pass

    if not row or not row[0]:
        print(f"[office] No path for file_id {file_id}", file=sys.stderr)
        return None

    path = row[0]
    if os.path.isfile(path):
        return path

    # Fallback: uploads directory
    candidate = os.path.join(_UPLOAD_DIR, os.path.basename(path))
    if os.path.isfile(candidate):
        return candidate

    # Last resort: UUID prefix match in uploads
    prefix = file_id.split("-")[0] if "-" in file_id else file_id[:8]
    if os.path.isdir(_UPLOAD_DIR):
        for name in os.listdir(_UPLOAD_DIR):
            if name.startswith(prefix):
                candidate = os.path.join(_UPLOAD_DIR, name)
                if os.path.isfile(candidate):
                    return candidate

    print(f"[office] File not found on disk: {path}", file=sys.stderr)
    return None


def _read_file_bytes(file_id: str) -> Optional[bytes]:
    """Return raw bytes for an Open WebUI file."""
    path = _resolve_file_path(file_id)
    if path is None:
        return None
    # Path traversal guard: ensure resolved path stays inside allowed directories
    try:
        _abs = os.path.realpath(path)
        _allowed = False
        for _base in (_UPLOAD_DIR, _EXPORT_DIR,
                       os.path.join(os.environ.get("OPEN_WEBUI_DATA_DIR", ""), "data"),
                       _get_owui_data_dir()):
            if _base and os.path.realpath(_base) in _abs:
                _allowed = True
                break
        if not _allowed:
            print(f"[office] Path traversal blocked: {path}", file=sys.stderr)
            return None
    except Exception:
        pass
    try:
        with open(path, "rb") as fh:
            return fh.read()
    except Exception as exc:
        print(f"[office] Read failed for {path}: {exc}", file=sys.stderr)
        return None


def _detect_type(filename: str) -> str:
    """Detect file type from extension."""
    ext = os.path.splitext(filename)[1].lower()
    if ext in (".xlsx",):
        return "xlsx"
    if ext in (".xls",):
        return "xls"
    if ext in (".docx",):
        return "docx"
    if ext in (".pptx",):
        return "pptx"
    if ext in (".ppt",):
        return "ppt"
    if ext in (".odt",):
        return "odt"
    if ext in (".ods",):
        return "ods"
    if ext in (".odp",):
        return "odp"
    return "unknown"


def _cell_value(cell):
    """Extract a JSON-safe value from an openpyxl cell."""
    import datetime

    v = cell.value
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def _xls_to_xlsx(xls_data: bytes) -> bytes:
    """Convert .xls bytes to .xlsx bytes using xlrd + openpyxl.

    Returns an in-memory .xlsx workbook as bytes.
    """
    import xlrd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    xls_book = xlrd.open_workbook(file_contents=xls_data)
    xlsx_wb = openpyxl.Workbook()
    # Remove the default sheet; we'll add one per xls sheet
    xlsx_wb.remove(xlsx_wb.active)

    for sheet_idx in range(xls_book.nsheets):
        xls_sheet = xls_book.sheet_by_index(sheet_idx)
        ws = xlsx_wb.create_sheet(title=xls_sheet.name[:31])  # Excel 31-char limit

        for rx in range(xls_sheet.nrows):
            for cx in range(xls_sheet.ncols):
                cell = xls_sheet.cell(rx, cx)
                value = cell.value

                # xlrd date handling: if cell type is XL_CELL_DATE, convert
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt_tuple = xls_book.datemode, int(value)
                        import datetime as _dt
                        value = _dt.datetime(*xlrd.xldate_as_tuple(value, xls_book.datemode))
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)

                ws.cell(row=rx + 1, column=cx + 1, value=value)

        # Auto-fit column widths (rough estimate)
        for col_cells in ws.columns:
            max_len = 0
            for cell in col_cells:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            letter = openpyxl.utils.get_column_letter(col_cells[0].column)
            ws.column_dimensions[letter].width = min(max_len + 2, 50)

    out = io.BytesIO()
    xlsx_wb.save(out)
    xlsx_wb.close()
    xls_book.release_resources()
    out.seek(0)
    return out.read()


def _format_text(text: str) -> str:
    """Apply consistent text formatting: normalize dashes, sentence case, preserve acronyms."""
    if not isinstance(text, str) or not text or text.startswith('='):
        return text
    
    # 1. Normalize dashes
    text = text.replace('\u2014', '-').replace('\u2015', '-').replace('\u2013', '-')
    
    # 2. Split into sentences and apply sentence case
    acronyms = {'API', 'PDF', 'HTML', 'CSS', 'JSON', 'SQL', 'AI', 'UI', 'UX', 'ID', 'URL', 'HTTP', 'HTTPS', 'FTP', 'SSH', 'DNS', 'IP', 'TCP', 'UDP', 'SSL', 'TLS', 'REST', 'CRUD', 'YAML', 'XML', 'SVG', 'PNG', 'JPG', 'GIF', 'CSV', 'DOCX', 'PPTX', 'XLSX', 'DOC', 'PPT', 'XLS', 'ISO', 'RGB', 'CMYK', 'PHP', 'NPM', 'YARN', 'CLI', 'GUI', 'IDE', 'SDK', 'JDK', 'JRE', 'JVM', 'DB', 'SQLite', 'MySQL', 'NoSQL', 'OAuth', 'JWT', 'CORS', 'MVC', 'MVP', 'MVVM', 'SPA', 'PWA', 'SSR', 'CSR', 'SSG', 'ISR', 'CDN', 'DHCP', 'NAT', 'VPN', 'LAN', 'WAN', 'MAC', 'BIOS', 'UEFI', 'GPT', 'MBR', 'RAID', 'NAS', 'SAN', 'AWS', 'GCP', 'Azure', 'SaaS', 'PaaS', 'IaaS', 'FaaS', 'CaaS', 'K8s', 'Docker', 'Kubernetes'}
    
    import re as _re
    
    # Protect acronyms with placeholders (longest first to avoid partial matches)
    acronym_map = {}
    for i, acro in enumerate(sorted(acronyms, key=len, reverse=True)):
        placeholder = f'\x00{i:04d}\x00'
        pattern = _re.compile(r'(?<![a-zA-Z])' + _re.escape(acro) + r'(?![a-zA-Z])', _re.IGNORECASE)
        text = pattern.sub(placeholder, text)
        acronym_map[placeholder] = acro
    
    # Sentence case
    sentences = _re.split(r'(?<=[.!?])\s+', text)
    formatted_sentences = []
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        s = s.lower()
        if s and s[0].isalpha():
            s = s[0].upper() + s[1:]
        formatted_sentences.append(s)
    
    text = ' '.join(formatted_sentences)
    
    # Restore acronyms
    for placeholder, acro in acronym_map.items():
        text = text.replace(placeholder, acro)
    
    return text




# ---------------------------------------------------------------------------
# Office Plugin Registry — allows external plugins to register document processors
# ---------------------------------------------------------------------------
_office_plugins: dict = {}

def register_office_plugin(name: str):
    """Decorator to register an office document processor plugin."""
    def decorator(func):
        _office_plugins[name] = func
        return func
    return decorator

def _call_office_plugins(plugin_type: str, *args, **kwargs) -> dict:
    """Call all registered plugins of a given type and return results dict."""
    results = {}
    for pname, plugin in _office_plugins.items():
        try:
            if callable(plugin):
                results[pname] = plugin(plugin_type, *args, **kwargs)
        except Exception as e:
            print(f"[office-plugin] '{pname}' failed on '{plugin_type}': {e}", file=sys.stderr)
    return results

# ---------------------------------------------------------------------------
# Base64 Filename Encoding — safe storage of filenames in SQLite3
# ---------------------------------------------------------------------------
import base64 as _b64_mod

def _encode_filename(filename: str) -> str:
    """Encode filename to base64 for safe SQLite3 storage (handles international chars, special chars, path traversal)."""
    if not filename:
        return filename
    safe = filename.encode('utf-8')
    encoded = _b64_mod.urlsafe_b64encode(safe).decode('ascii').rstrip('=')
    _, ext = os.path.splitext(filename)
    return encoded + ext

def _decode_filename(encoded_name: str) -> str:
    """Decode a base64-encoded filename back to original. Returns as-is if decoding fails."""
    try:
        base = os.path.splitext(encoded_name)[0]
        padding = 4 - len(base) % 4
        if padding != 4:
            base += '=' * padding
        decoded = _b64_mod.urlsafe_b64decode(base).decode('utf-8')
        return decoded + os.path.splitext(encoded_name)[1]
    except Exception:
        return encoded_name



def _read_odf(file_bytes: bytes, filename: str) -> str:
    """Read ODF files (.odt, .ods, .odp) and return structured text."""
    ext = os.path.splitext(filename)[1].lower()
    
    try:
        if ext == ".ods":
            from odf.opendocument import load
            from odf.table import Table, TableRow, TableCell
            from odf.text import P
            
            doc = load(io.BytesIO(file_bytes))
            result = []
            for table in doc.getElementsByType(Table):
                for row in table.getElementsByType(TableRow):
                    cells = []
                    for cell in row.getElementsByType(TableCell):
                        text_parts = []
                        for p in cell.getElementsByType(P):
                            try:
                                text_parts.append(str(p))
                            except Exception:
                                pass
                        cells.append(" ".join(text_parts).strip())
                    result.append(" | ".join(cells))
            return "\n".join(result) if result else "(empty spreadsheet)"
        
        elif ext == ".odt":
            from odf.opendocument import load
            from odf.text import P, H
            
            doc = load(io.BytesIO(file_bytes))
            result = []
            for elem in doc.getElementsByType(H):
                try:
                    result.append(f"## {str(elem)}")
                except Exception:
                    pass
            for elem in doc.getElementsByType(P):
                try:
                    text = str(elem).strip()
                    if text:
                        result.append(text)
                except Exception:
                    pass
            return "\n\n".join(result) if result else "(empty document)"
        
        elif ext == ".odp":
            from odf.opendocument import load
            from odf.text import P
            
            doc = load(io.BytesIO(file_bytes))
            result = []
            slide_num = 0
            for elem in doc.getElementsByType(P):
                try:
                    text = str(elem).strip()
                    if text:
                        result.append(f"Slide {slide_num + 1}: {text}")
                        slide_num += 1
                except Exception:
                    pass
            return "\n".join(result) if result else "(empty presentation)"
        
        return f"Unsupported ODF format: {ext}"
    except ImportError:
        return "Error: odfpy library not installed. Install with: pip install odfpy"
    except Exception as e:
        return f"Error reading {ext} file: {str(e)}"



# ---------------------------------------------------------------------------
# Professional document helpers
# ---------------------------------------------------------------------------
def _add_callout_box(doc, lines, colors, hex_to_rgb):
    """Add a professional callout box (note/tip/warning)."""
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml

    first_line = lines[0].lower() if lines else ""
    if first_line.startswith("**warning") or first_line.startswith("**alert"):
        border_color = "E74C3C"
        bg_color = "FDEDEC"
        icon = "\u26a0\ufe0f"
    elif first_line.startswith("**tip") or first_line.startswith("**pro tip"):
        border_color = "27AE60"
        bg_color = "E8F8F5"
        icon = "\U0001f4a1"
    elif first_line.startswith("**note") or first_line.startswith("**info") or True:
        border_color = colors.get("accent", "2E75B6")
        bg_color = colors.get("light", "D6E4F0")
        icon = "\U0001f4cc"

    table = doc.add_table(rows=1, cols=1)
    table.alignment = 1
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.0)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}" w:val="clear"/>')
    tcPr.append(shading)

    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="24" w:space="8" w:color="{border_color}"/>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    for i, line in enumerate(lines):
        text = _format_text(line)
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        if i == 0:
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(f"{icon} {text}")
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.name = 'Calibri'
        else:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            run = p.add_run(text)
            run.font.size = Pt(10)
            run.font.name = 'Calibri'

    doc.add_paragraph()


def _add_professional_table(doc, rows, colors, hex_to_rgb):
    """Add a professionally styled table."""
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml

    if not rows:
        return

    num_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=num_cols)
    table.style = 'Colorful Grid Accent 1'
    table.alignment = 1

    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            if j < num_cols:
                cell = table.rows[i].cells[j]
                cell.text = _format_text(cell_text)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(2)
                    paragraph.paragraph_format.space_after = Pt(2)
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Calibri'
                        if i == 0:
                            run.font.bold = True

    doc.add_paragraph()


def _render_content_slide(prs, lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_num):
    """Render a content slide with professional layout."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, colors["bg"])

    y_pos = 0.5
    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('# ') or line.startswith('## '):
            text = _format_text(line.lstrip('#').strip())
            add_accent_bar(slide, 0.8, y_pos + 0.15, 0.06, 0.5, colors["accent"])
            add_text_box(slide, 1.1, y_pos, 11, 0.7, text, font_size=28, bold=True, color=colors["text"])
            y_pos += 0.9
        elif line.startswith('- ') or line.startswith('* '):
            text = _format_text(line[2:].strip())
            add_text_box(slide, 1.5, y_pos, 10.5, 0.5, "\u2022 " + text, font_size=16, color=colors["text"])
            y_pos += 0.5
        elif line.startswith('|'):
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if any(c.startswith('---') for c in cells):
                continue
            add_text_box(slide, 1.0, y_pos, 11, 0.5, "  |  ".join(cells), font_size=14, color=colors["text"])
            y_pos += 0.4
        else:
            text = _format_text(line)
            add_text_box(slide, 1.0, y_pos, 11.5, 0.5, text, font_size=16, color=colors["text"])
            y_pos += 0.5

        if y_pos > 6.5:
            break


# =========================================================================
class Tools:
    class Valves(BaseModel):
        base_url: Optional[str] = Field(
            default=None,
            description="Override the base URL for download links. Auto-detected from X-Original-Host header or WEBUI_URL env var if unset.",
        )
        pass

    def __init__(self):
        self.valves = self.Valves()

    # -----------------------------------------------------------------
    # Internal: save and return markdown link
    # -----------------------------------------------------------------
    async def _save_and_link(self, file_bytes: bytes, filename: str, __request__=None) -> tuple:
        """Save file to Open WebUI uploads dir, register in DB, return download URL."""
        import base64 as _b64
        import hashlib
        import time as _time
        import uuid as _uuid

        ext = os.path.splitext(filename)[1].lower()
        mt = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".xls": "application/vnd.ms-excel",
            ".doc": "application/msword",
            ".ppt": "application/vnd.ms-powerpoint",
            ".pdf": "application/pdf",
            ".odt": "application/vnd.oasis.opendocument.text",
            ".ods": "application/vnd.oasis.opendocument.spreadsheet",
            ".odp": "application/vnd.oasis.opendocument.presentation",
        }
        content_type = mt.get(ext, "application/octet-stream")

        try:
            file_id = str(_uuid.uuid4())
            with open(os.path.join(_UPLOAD_DIR, file_id), "wb") as f:
                f.write(file_bytes)

            file_hash = hashlib.sha256(file_bytes).hexdigest()[:16]
            now = int(_time.time())

            conn = sqlite3.connect(_DB_PATH)
            conn.execute(
                """INSERT OR REPLACE INTO file
                   (id, user_id, hash, filename, path, data, meta, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    file_id,
                    "",
                    file_hash,
                    _encode_filename(filename),
                    os.path.join(_UPLOAD_DIR, file_id),
                    "{}",
                    json.dumps({
                        "name": filename,
                        "content_type": content_type,
                        "size": len(file_bytes),
                        "source": "office-plugin",
                        "generated": True,
                    }),
                    now,
                    now,
                ),
            )
            conn.commit()
            conn.close()

            base_url = self.valves.base_url
            if not base_url and __request__:
                try:
                    host = __request__.headers.get("x-original-host")
                    if host:
                        base_url = f"https://{host}"
                except Exception:
                    pass
            if not base_url:
                base_url = os.environ.get("WEBUI_URL", "http://localhost:3000")
            base_url = base_url.rstrip("/")

            url = f"{base_url}/api/v1/files/{file_id}/content"
            return (url, filename)

        except Exception as e:
            print(f"[office] Save failed: {e}", file=sys.stderr)
            try:
                data = _b64.b64encode(file_bytes).decode("ascii")
                return (f"data:{content_type};base64,{data}", filename)
            except Exception:
                return (None, None)

    # -----------------------------------------------------------------
    # READ
    # -----------------------------------------------------------------
    async def read_file(
        self,
        file_id: str,
        max_rows: int = 500,
        __user__=None,
        __request__=None,
    ) -> str:
        """Read any Office file (.xlsx, .xls, .docx, .pptx) and return its contents as structured JSON.

        Auto-detects the file type from the file ID or filename.
        For xlsx/xls: returns sheets with headers and rows.
        For docx: returns paragraphs with styles and tables.
        For pptx: returns slides with shapes and text.
        Legacy .doc and .ppt formats return a helpful error message.

        Args:
            file_id: The Open WebUI file ID (UUID) or filename
            max_rows: Maximum rows to return for xlsx (default 500)
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": (
                        f"Could not read file {file_id}. "
                        "Make sure the file was uploaded via the chat."
                    )
                })

            # Detect type from filename in DB
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            result: Dict[str, Any] = {
                "file_id": file_id,
                "filename": filename,
                "type": file_type,
            }

            if file_type == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
                result["sheets"] = []
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    sheet: Dict[str, Any] = {
                        "name": sn,
                        "headers": [],
                        "rows": [],
                        "total_rows": ws.max_row or 0,
                        "total_cols": ws.max_column or 0,
                    }
                    max_r = min(ws.max_row or 0, max_rows)
                    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=max_r), 1):
                        rd = [_cell_value(c) for c in row]
                        if ri == 1:
                            sheet["headers"] = [str(v) if v is not None else "" for v in rd]
                        else:
                            sheet["rows"].append(rd)
                    result["sheets"].append(sheet)
                wb.close()

            elif file_type == "xls":
                import xlrd
                xls_book = xlrd.open_workbook(file_contents=file_data)
                result["sheets"] = []
                for sheet_idx in range(xls_book.nsheets):
                    xls_sheet = xls_book.sheet_by_index(sheet_idx)
                    sheet: Dict[str, Any] = {
                        "name": xls_sheet.name,
                        "headers": [],
                        "rows": [],
                        "total_rows": xls_sheet.nrows,
                        "total_cols": xls_sheet.ncols,
                    }
                    max_r = min(xls_sheet.nrows, max_rows)
                    for rx in range(max_r):
                        row_values = []
                        for cx in range(xls_sheet.ncols):
                            cell = xls_sheet.cell(rx, cx)
                            value = cell.value
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                try:
                                    import datetime as _dt
                                    value = _dt.datetime(*xlrd.xldate_as_tuple(value, xls_book.datemode))
                                except Exception:
                                    pass
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                value = bool(value)
                            row_values.append(value)
                        if rx == 0:
                            sheet["headers"] = [str(v) if v is not None else "" for v in row_values]
                        else:
                            sheet["rows"].append(row_values)
                    result["sheets"].append(sheet)
                xls_book.release_resources()

            elif file_type == "docx":
                from docx import Document
                from docx.enum.text import WD_COLOR_INDEX
                doc = Document(io.BytesIO(file_data))
                paragraphs = []
                for p in doc.paragraphs:
                    if p.text.strip():
                        style = p.style.name if p.style else "Normal"
                        runs_info = []
                        for run in p.runs:
                            run_data = {"text": run.text}
                            if run.font.highlight_color and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
                                run_data["highlighted"] = True
                                run_data["highlight_color"] = str(run.font.highlight_color)
                            if run.font.bold:
                                run_data["bold"] = True
                            if run.font.italic:
                                run_data["italic"] = True
                            runs_info.append(run_data)
                        paragraphs.append({"style": style, "text": p.text, "runs": runs_info})
                tables = []
                for t in doc.tables:
                    tbl = {"rows": []}
                    for row in t.rows:
                        tbl["rows"].append([cell.text for cell in row.cells])
                    tables.append(tbl)
                result["paragraphs"] = paragraphs
                result["tables"] = tables

            elif file_type == "pptx":
                from pptx import Presentation
                prs = Presentation(io.BytesIO(file_data))
                slides = []
                for si, slide in enumerate(prs.slides, 1):
                    sdata: Dict[str, Any] = {"number": si, "shapes": []}
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            sdata["shapes"].append({
                                "type": str(shape.shape_type),
                                "name": shape.name,
                                "text": shape.text[:500],
                            })
                        if shape.has_table:
                            tbl = {"rows": []}
                            for row in shape.table.rows:
                                tbl["rows"].append([cell.text for cell in row.cells])
                            sdata["tables"] = tbl
                    slides.append(sdata)
                result["slides"] = slides

            elif file_type == "doc":
                result["error"] = "Legacy .doc format is not supported. Please convert to .docx first."

            elif file_type == "ppt":
                result["error"] = "Legacy .ppt format is not supported. Please convert to .pptx first."

            else:
                result["error"] = f"Unsupported file type. Detected: {file_type}. Supported: xlsx, xls, docx, pptx"

            return json.dumps(result, indent=2, default=str, ensure_ascii=False)

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # ADD CONTENT
    # -----------------------------------------------------------------
    async def add_content(
        self,
        file_id: str,
        content: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Add new content to an Office file while preserving original formatting.

        For spreadsheets (xlsx): content is CSV text with rows to add.
        For documents (docx): content is text to append at the end.
        For presentations (pptx): each line defines a new slide. Use "---" as separator between slides.

        Args:
            file_id: File ID to edit
            content: Content to add (CSV for xlsx, text for docx/pptx)
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            # Detect type
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            out = io.BytesIO()
            out_name = output_filename

            if file_type == "xlsx":
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                ws = wb.active
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"

                # Parse CSV content
                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                parsed_rows = []
                for csv_row in reader:
                    converted = []
                    for v in csv_row:
                        v = v.strip()
                        if v == '':
                            converted.append(None)
                        else:
                            try:
                                converted.append(int(v))
                            except ValueError:
                                try:
                                    converted.append(float(v))
                                except ValueError:
                                    if v.lower() == 'true':
                                        converted.append(True)
                                    elif v.lower() == 'false':
                                        converted.append(False)
                                    else:
                                        converted.append(v)
                    parsed_rows.append(converted)

                if not parsed_rows:
                    return json.dumps({"error": "No rows provided in CSV content"})

                # Get reference styles from last row
                ref = {}
                if ws.max_row and ws.max_row >= 1:
                    for cell in ws[ws.max_row]:
                        if cell.has_style:
                            ref[cell.column] = {
                                "font": copy(cell.font),
                                "fill": copy(cell.fill),
                                "border": copy(cell.border),
                                "alignment": copy(cell.alignment),
                                "number_format": cell.number_format,
                            }

                start = (ws.max_row or 0) + 1
                for i, rd in enumerate(parsed_rows):
                    for j, v in enumerate(rd, 1):
                        cell = ws.cell(row=start + i, column=j)
                        if j in ref:
                            try:
                                cell.font = copy(ref[j]["font"])
                                cell.fill = copy(ref[j]["fill"])
                                cell.border = copy(ref[j]["border"])
                                cell.alignment = copy(ref[j]["alignment"])
                                cell.number_format = ref[j]["number_format"]
                            except Exception:
                                pass
                        cell.value = v

                wb.save(out)
                wb.close()

            elif file_type == "xls":
                # Convert .xls to .xlsx, then apply same add logic
                file_data = _xls_to_xlsx(file_data)
                file_type = "xlsx"
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                ws = wb.active

                # Parse CSV content
                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                parsed_rows = []
                for csv_row in reader:
                    converted = []
                    for v in csv_row:
                        v = v.strip()
                        if v == '':
                            converted.append(None)
                        else:
                            try:
                                converted.append(int(v))
                            except ValueError:
                                try:
                                    converted.append(float(v))
                                except ValueError:
                                    if v.lower() == 'true':
                                        converted.append(True)
                                    elif v.lower() == 'false':
                                        converted.append(False)
                                    else:
                                        converted.append(v)
                    parsed_rows.append(converted)

                if not parsed_rows:
                    return json.dumps({"error": "No rows provided in CSV content"})

                ref = {}
                if ws.max_row and ws.max_row >= 1:
                    for cell in ws[ws.max_row]:
                        if cell.has_style:
                            ref[cell.column] = {
                                "font": copy(cell.font),
                                "fill": copy(cell.fill),
                                "border": copy(cell.border),
                                "alignment": copy(cell.alignment),
                                "number_format": cell.number_format,
                            }

                start = (ws.max_row or 0) + 1
                for i, rd in enumerate(parsed_rows):
                    for j, v in enumerate(rd, 1):
                        cell = ws.cell(row=start + i, column=j)
                        if j in ref:
                            try:
                                cell.font = copy(ref[j]["font"])
                                cell.fill = copy(ref[j]["fill"])
                                cell.border = copy(ref[j]["border"])
                                cell.alignment = copy(ref[j]["alignment"])
                                cell.number_format = ref[j]["number_format"]
                            except Exception:
                                pass
                        cell.value = v

                wb.save(out)
                wb.close()

            elif file_type == "docx":
                from docx import Document
                doc = Document(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.docx"

                # Append paragraphs, preserving last paragraph's style
                last_style = "Normal"
                if doc.paragraphs:
                    last_style = doc.paragraphs[-1].style.name if doc.paragraphs[-1].style else "Normal"

                for line in content.split("\n"):
                    doc.add_paragraph(line, style=last_style)

                out = io.BytesIO()
                doc.save(out)

            elif file_type == "pptx":
                from pptx import Presentation
                from pptx.util import Inches
                prs = Presentation(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.pptx"

                # Split content by "---" into slides
                slide_specs = re.split(r'\n---\n|\r\n---\r\n|\n---\n', content)
                blank_layout = prs.slide_layouts[6]  # Blank layout

                for spec in slide_specs:
                    spec = spec.strip()
                    if not spec:
                        continue
                    lines = spec.split("\n")
                    title = lines[0].strip()
                    body = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""

                    slide = prs.slides.add_slide(blank_layout)
                    # Add title
                    txBox = slide.shapes.add_textbox(
                        Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                    )
                    tf = txBox.text_frame
                    tf.text = title
                    p = tf.paragraphs[0]
                    p.font.size = Inches(0.6)
                    p.font.bold = True

                    # Add body text
                    if body:
                        txBox2 = slide.shapes.add_textbox(
                            Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
                        )
                        tf2 = txBox2.text_frame
                        tf2.text = body
                        for para in tf2.paragraphs:
                            para.font.size = Inches(0.3)

                out = io.BytesIO()
                prs.save(out)

            else:
                return json.dumps({"error": f"Unsupported type: {file_type}"})

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\nAdded content to {file_type.upper()} file, preserving original formatting."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # REPLACE TEXT
    # -----------------------------------------------------------------
    async def replace_text(
        self,
        file_id: str,
        find_text: str,
        replace_with: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Find and replace text in any Office file while preserving original formatting.

        Works on cell values in xlsx, paragraph text in docx, and shape text in pptx.

        Args:
            file_id: File ID to edit
            find_text: Text to find
            replace_with: Text to replace with
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            # Detect type
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            out = io.BytesIO()
            out_name = output_filename
            count = 0

            if file_type == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"

                for sn in wb.sheetnames:
                    ws = wb[sn]
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value is None:
                                continue
                            if isinstance(cell.value, str) and find_text in cell.value:
                                cell.value = cell.value.replace(find_text, replace_with)
                                count += 1
                            elif not isinstance(cell.value, str):
                                sval = str(cell.value)
                                if find_text in sval:
                                    cell.value = replace_with
                                    count += 1

                wb.save(out)
                wb.close()

            elif file_type == "xls":
                # Convert .xls to .xlsx, then apply same replace logic
                file_data = _xls_to_xlsx(file_data)
                file_type = "xlsx"
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data))

                for sn in wb.sheetnames:
                    ws = wb[sn]
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value is None:
                                continue
                            if isinstance(cell.value, str) and find_text in cell.value:
                                cell.value = cell.value.replace(find_text, replace_with)
                                count += 1
                            elif not isinstance(cell.value, str):
                                sval = str(cell.value)
                                if find_text in sval:
                                    cell.value = replace_with
                                    count += 1

                wb.save(out)
                wb.close()

            elif file_type == "docx":
                from docx import Document
                doc = Document(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.docx"

                for para in doc.paragraphs:
                    if find_text in para.text:
                        # Preserve formatting: replace in runs
                        full_text = para.text
                        if find_text in full_text:
                            new_text = full_text.replace(find_text, replace_with)
                            # Clear all runs and set new text in first run
                            if para.runs:
                                para.runs[0].text = new_text
                                for run in para.runs[1:]:
                                    run.text = ""
                                count += 1
                            else:
                                para.text = new_text
                                count += 1

                # Also replace in tables
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if find_text in cell.text:
                                for para in cell.paragraphs:
                                    if find_text in para.text:
                                        if para.runs:
                                            para.runs[0].text = para.text.replace(find_text, replace_with)
                                            for run in para.runs[1:]:
                                                run.text = ""
                                        else:
                                            para.text = para.text.replace(find_text, replace_with)
                                        count += 1

                doc.save(out)

            elif file_type == "pptx":
                from pptx import Presentation
                prs = Presentation(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.pptx"

                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame"):
                            for para in shape.text_frame.paragraphs:
                                if find_text in para.text:
                                    if para.runs:
                                        para.runs[0].text = para.text.replace(find_text, replace_with)
                                        for run in para.runs[1:]:
                                            run.text = ""
                                    else:
                                        para.text = para.text.replace(find_text, replace_with)
                                    count += 1

                prs.save(out)

            else:
                return json.dumps({"error": f"Unsupported type: {file_type}"})

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\nReplaced '{find_text}' with '{replace_with}' in {count} place(s), preserving all formatting."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # CREATE NEW FILE
    # -----------------------------------------------------------------
    async def create_file(
        self,
        file_type: str,
        content: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Create a new Office file from scratch.

        For xlsx: content is CSV with headers on first line.
        For docx: content is plain text (one paragraph per line).
        For pptx: each line defines a slide. Use "---" as separator between slides.

        Args:
            file_type: 'xlsx', 'docx', or 'pptx'
            content: Content specification
            output_filename: Output filename
        """
        try:
            ftype = file_type.lower().replace(".", "")
            if ftype not in ("xlsx", "docx", "pptx"):
                return json.dumps({"error": f"Unsupported type: {file_type}. Use xlsx, docx, or pptx."})

            out_name = output_filename or f"document.{ftype}"
            out = io.BytesIO()

            if ftype == "xlsx":
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.Workbook()
                ws = wb.active

                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                rows = list(reader)

                if rows:
                    # First row = headers (styled)
                    for j, h in enumerate(rows[0], 1):
                        c = ws.cell(row=1, column=j, value=h.strip())
                        c.font = Font(bold=True, color="FFFFFF", size=11)
                        c.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                        c.alignment = Alignment(horizontal="center")

                    # Data rows
                    for i, rd in enumerate(rows[1:], 2):
                        for j, v in enumerate(rd, 1):
                            v = v.strip()
                            c = ws.cell(row=i, column=j)
                            if v == '':
                                c.value = None
                            else:
                                try:
                                    c.value = int(v)
                                except ValueError:
                                    try:
                                        c.value = float(v)
                                    except ValueError:
                                        if v.lower() == 'true':
                                            c.value = True
                                        elif v.lower() == 'false':
                                            c.value = False
                                        else:
                                            c.value = v
                            c.alignment = Alignment(
                                horizontal="center" if isinstance(c.value, (int, float)) else "left"
                            )

                    # Auto-fit columns
                    for col in ws.columns:
                        mx = max((len(str(c.value or "")) for c in col), default=5)
                        ws.column_dimensions[col[0].column_letter].width = min(mx + 3, 50)

                wb.save(out)
                wb.close()

            elif ftype == "docx":
                from docx import Document
                from docx.shared import Pt
                doc = Document()
                for line in content.split("\n"):
                    doc.add_paragraph(line)
                doc.save(out)

            elif ftype == "pptx":
                from pptx import Presentation
                from pptx.util import Inches
                prs = Presentation()
                slide_specs = re.split(r'\n---\n|\r\n---\r\n|\n---\n', content)
                blank_layout = prs.slide_layouts[6]

                for spec in slide_specs:
                    spec = spec.strip()
                    if not spec:
                        continue
                    lines = spec.split("\n")
                    title = lines[0].strip()
                    body = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""

                    slide = prs.slides.add_slide(blank_layout)
                    txBox = slide.shapes.add_textbox(
                        Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                    )
                    tf = txBox.text_frame
                    tf.text = title
                    p = tf.paragraphs[0]
                    p.font.size = Inches(0.6)
                    p.font.bold = True

                    if body:
                        txBox2 = slide.shapes.add_textbox(
                            Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
                        )
                        tf2 = txBox2.text_frame
                        tf2.text = body
                        for para in tf2.paragraphs:
                            para.font.size = Inches(0.3)

                prs.save(out)

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\nCreated new {ftype.upper()} file."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})


    async def generate_document(self, content: str, title: str = "Document", theme: str = "professional", __user__=None, __request__=None) -> str:
        """Generate a professional Word document with modern styling and themes.

        Args:
            content: Markdown-formatted content
            title: Document title (used for cover page and filename)
            theme: Visual theme - professional, modern, creative, corporate, minimal, elegant
        Returns:
            Markdown link to the generated file
        """
        try:
            from docx import Document
            from docx.shared import Inches, Pt, Cm, RGBColor, Emu
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.enum.table import WD_TABLE_ALIGNMENT
            from docx.oxml.ns import qn, nsdecls
            from docx.oxml import parse_xml
            import datetime

            doc = Document()

            # --- Page Setup ---
            section = doc.sections[0]
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

            # --- Color Themes ---
            themes = {
                "professional": {"primary": "1F4E79", "accent": "2E75B6", "light": "D6E4F0", "text": "333333"},
                "modern": {"primary": "2D3436", "accent": "6C5CE7", "light": "DFE6E9", "text": "2D3436"},
                "creative": {"primary": "E17055", "accent": "FDCB6E", "light": "FFF3E0", "text": "2D3436"},
                "corporate": {"primary": "003366", "accent": "CC0000", "light": "E8EEF4", "text": "1A1A1A"},
                "minimal": {"primary": "000000", "accent": "666666", "light": "F5F5F5", "text": "333333"},
                "elegant": {"primary": "4A235A", "accent": "8E44AD", "light": "F3E5F5", "text": "1A1A1A"},
            }
            colors = themes.get(theme, themes["professional"])

            def hex_to_rgb(hex_color):
                return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

            # --- Default Font ---
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Calibri'
            font.size = Pt(11)
            font.color.rgb = hex_to_rgb(colors["text"])
            style.paragraph_format.space_after = Pt(8)
            style.paragraph_format.line_spacing = 1.15

            # --- Heading Styles ---
            for i in range(1, 4):
                heading_style = doc.styles['Heading %d' % i]
                heading_style.font.name = 'Calibri'
                heading_style.font.color.rgb = hex_to_rgb(colors["primary"])
                if i == 1:
                    heading_style.font.size = Pt(24)
                    heading_style.paragraph_format.space_before = Pt(24)
                    heading_style.paragraph_format.space_after = Pt(12)
                elif i == 2:
                    heading_style.font.size = Pt(18)
                    heading_style.paragraph_format.space_before = Pt(18)
                    heading_style.paragraph_format.space_after = Pt(8)
                else:
                    heading_style.font.size = Pt(14)
                    heading_style.paragraph_format.space_before = Pt(12)
                    heading_style.paragraph_format.space_after = Pt(6)

            # --- Cover Page ---
            cover_table = doc.add_table(rows=1, cols=1)
            cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = cover_table.rows[0].cells[0]
            cell.width = Inches(6.5)
            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["primary"]}"/>')
            cell._tc.get_or_add_tcPr().append(shading_elm)

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(40)
            p.paragraph_format.space_after = Pt(8)
            run = p.add_run(_format_text(title))
            run.font.size = Pt(28)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.name = 'Calibri'

            p = cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(30)
            run = p.add_run(datetime.datetime.now().strftime("%B %d, %Y"))
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(200, 200, 200)
            run.font.name = 'Calibri'

            doc.add_paragraph()

            # --- Process Content ---
            lines = content.split('\n')
            table_rows = []
            callout_lines = []

            for line in lines:
                line = line.strip()
                if not line:
                    if callout_lines:
                        _add_callout_box(doc, callout_lines, colors, hex_to_rgb)
                        callout_lines = []
                    continue

                if line.startswith('# ') or line.startswith('## ') or line.startswith('### '):
                    level = line.count('#')
                    text = _format_text(line.lstrip('#').strip())
                    doc.add_heading(text, level=min(level, 3))
                    continue

                if line.startswith('> '):
                    callout_lines.append(line[2:])
                    continue
                elif callout_lines:
                    _add_callout_box(doc, callout_lines, colors, hex_to_rgb)
                    callout_lines = []

                if line.startswith('|') and line.endswith('|'):
                    cells = [c.strip() for c in line.split('|')[1:-1]]
                    if all(c.startswith('---') for c in cells):
                        continue
                    table_rows.append(cells)
                    continue
                elif table_rows:
                    _add_professional_table(doc, table_rows, colors, hex_to_rgb)
                    table_rows = []

                if line.startswith('- ') or line.startswith('* '):
                    text = _format_text(line[2:].strip())
                    doc.add_paragraph(text, style='List Bullet')
                    continue

                if re.match(r'^\d+\.\s', line):
                    text = _format_text(re.sub(r'^\d+\.\s', '', line))
                    doc.add_paragraph(text, style='List Number')
                    continue

                text = _format_text(line)
                doc.add_paragraph(text)

            if table_rows:
                _add_professional_table(doc, table_rows, colors, hex_to_rgb)

            # --- Footer with page numbers ---
            footer = section.footer
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run("Page ")
            run.font.size = Pt(8)
            run.font.color.rgb = hex_to_rgb(colors["accent"])
            fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
            run._r.append(fldChar1)
            instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
            run._r.append(instrText)
            fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
            run._r.append(fldChar2)

            out = io.BytesIO()
            doc.save(out)
            out.seek(0)
            url, fname = await self._save_and_link(out.getvalue(), "%s.docx" % title, __request__)
            if url:
                return "Document created: [%s](%s)" % (fname, url)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def generate_slides(self, content: str, title: str = "Presentation", theme: str = "modern", __user__=None, __request__=None) -> str:
        """Generate professional PowerPoint slides with modern design.

        Args:
            content: Markdown content (headings become slides, bullets become content)
            title: Presentation title for the first slide
            theme: Visual theme - modern, light, dark, corporate, creative, minimal
        Returns:
            Markdown link to the generated file
        """
        try:
            from pptx import Presentation
            from pptx.util import Inches, Pt, Emu, Cm
            from pptx.dml.color import RGBColor
            from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
            from pptx.enum.shapes import MSO_SHAPE
            import datetime

            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            # --- Color Themes ---
            themes = {
                "modern": {"bg": "1A1A2E", "accent": "E94560", "text": "FFFFFF", "subtitle": "B0B0B0", "card": "16213E"},
                "light": {"bg": "FFFFFF", "accent": "2563EB", "text": "1E293B", "subtitle": "64748B", "card": "F1F5F9"},
                "dark": {"bg": "0F172A", "accent": "38BDF8", "text": "F8FAFC", "subtitle": "94A3B8", "card": "1E293B"},
                "corporate": {"bg": "FFFFFF", "accent": "003366", "text": "1A1A1A", "subtitle": "666666", "card": "F0F4F8"},
                "creative": {"bg": "FFF8F0", "accent": "FF6B6B", "text": "2D3436", "subtitle": "636E72", "card": "FFEAA7"},
                "minimal": {"bg": "FAFAFA", "accent": "000000", "text": "333333", "subtitle": "999999", "card": "F0F0F0"},
            }
            colors = themes.get(theme, themes["modern"])

            def hex_to_rgb(h):
                return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

            def set_slide_bg(slide, color_hex):
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = hex_to_rgb(color_hex)

            def add_text_box(slide, left, top, width, height, text, font_size=18, bold=False, color=None, alignment=PP_ALIGN.LEFT):
                txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = _format_text(text)
                p.font.size = Pt(font_size)
                p.font.bold = bold
                p.font.color.rgb = hex_to_rgb(color or colors["text"])
                p.font.name = 'Calibri'
                p.alignment = alignment
                return txBox

            def add_accent_bar(slide, left, top, width, height, color_hex):
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
                shape.fill.solid()
                shape.fill.fore_color.rgb = hex_to_rgb(color_hex)
                shape.line.fill.background()
                return shape

            # --- Slide 1: Title Slide ---
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_slide_bg(slide, colors["bg"])

            add_accent_bar(slide, 0, 3.0, 13.333, 0.06, colors["accent"])
            add_text_box(slide, 1.5, 1.5, 10, 1.5, _format_text(title), font_size=44, bold=True, color=colors["text"])
            add_text_box(slide, 1.5, 3.3, 10, 0.8, "Generated %s" % datetime.datetime.now().strftime("%B %d, %Y"), font_size=18, color=colors["subtitle"])

            # --- Process content into slides ---
            lines = content.split('\n')
            current_slide_lines = []
            slide_count = 1

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                if line.startswith('# ') or line.startswith('## '):
                    if current_slide_lines:
                        _render_content_slide(prs, current_slide_lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_count)
                        slide_count += 1
                    current_slide_lines = [line]
                else:
                    current_slide_lines.append(line)

            if current_slide_lines:
                _render_content_slide(prs, current_slide_lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_count)

            out = io.BytesIO()
            prs.save(out)
            out.seek(0)
            url, fname = await self._save_and_link(out.getvalue(), "%s.pptx" % title, __request__)
            if url:
                return "Presentation created: [%s](%s)" % (fname, url)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def generate_spreadsheet(self, content: str, title: str = "Spreadsheet", theme: str = "professional", __user__=None, __request__=None) -> str:
        """Generate a professional Excel spreadsheet with modern styling.

        Args:
            content: CSV or tab-delimited data (first row = headers, rest = data)
            title: Spreadsheet title / filename
            theme: Visual theme - professional, modern, corporate, minimal, colorful, pastel
        Returns:
            Markdown link to the generated file
        """
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from openpyxl.worksheet.table import Table, TableStyleInfo
            import io

            # --- Color Themes ---
            themes = {
                "professional": {"header": "1F4E79", "header_font": "FFFFFF", "alt_row": "F0F5FA", "border": "CCCCCC", "accent": "2E75B6"},
                "modern": {"header": "2D3436", "header_font": "FFFFFF", "alt_row": "F5F5F5", "border": "DFE6E9", "accent": "6C5CE7"},
                "corporate": {"header": "003366", "header_font": "FFFFFF", "alt_row": "E8EEF4", "border": "B0C4DE", "accent": "CC0000"},
                "minimal": {"header": "333333", "header_font": "FFFFFF", "alt_row": "F8F8F8", "border": "DDDDDD", "accent": "666666"},
                "colorful": {"header": "E17055", "header_font": "FFFFFF", "alt_row": "FFF3E0", "border": "FABEAB", "accent": "FDCB6E"},
                "pastel": {"header": "6C5CE7", "header_font": "FFFFFF", "alt_row": "F3E5F5", "border": "CE93D8", "accent": "A29BFE"},
            }
            colors = themes.get(theme, themes["professional"])

            header_fill = PatternFill(start_color=colors["header"], end_color=colors["header"], fill_type="solid")
            header_font = Font(name='Calibri', size=11, bold=True, color=colors["header_font"])
            alt_fill = PatternFill(start_color=colors["alt_row"], end_color=colors["alt_row"], fill_type="solid")
            body_font = Font(name='Calibri', size=11)
            thin_border = Border(
                left=Side(style='thin', color=colors["border"]),
                right=Side(style='thin', color=colors["border"]),
                top=Side(style='thin', color=colors["border"]),
                bottom=Side(style='thin', color=colors["border"])
            )

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = title[:31]

            # Parse CSV or tab-delimited content
            lines = [l.strip() for l in content.split('\n') if l.strip()]
            data = []
            for line in lines:
                sep = '\t' if '\t' in line else ','
                cells = [c.strip().strip('"') for c in line.split(sep)]
                data.append(cells)

            if not data:
                return json.dumps({"error": "No data provided"})

            # Write data
            for row in data:
                ws.append([_format_text(c) if isinstance(c, str) else c for c in row])

            # Style header row
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border

            # Zebra striping for data rows
            for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
                if i % 2 == 0:
                    for cell in row:
                        cell.fill = alt_fill
                        cell.border = thin_border
                        cell.font = body_font
                else:
                    for cell in row:
                        cell.border = thin_border
                        cell.font = body_font

            # Auto-fit columns
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value or '')))
                    except:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 3, 40)

            # Freeze top row
            ws.freeze_panes = 'A2'

            # Add auto-filter
            if ws.max_row > 1:
                ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(ws.max_column), ws.max_row)

            # Add Excel Table if data exists
            if ws.max_row > 1 and ws.max_column > 0:
                try:
                    tab = Table(
                        displayName="Table_" + ws.title.replace(' ', '')[:20],
                        ref="A1:%s%d" % (get_column_letter(ws.max_column), ws.max_row)
                    )
                    style = TableStyleInfo(
                        name="TableStyleMedium2",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False
                    )
                    tab.tableStyleInfo = style
                    ws.add_table(tab)
                except Exception:
                    pass

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            url, fname = await self._save_and_link(out.getvalue(), "%s.xlsx" % title, __request__)
            if url:
                return "Spreadsheet created: [%s](%s)" % (fname, url)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def tracked_change(self, file_id: str, change_type: str, content: str, author: str = "Reviewer", paragraph_index: int = -1, output_filename: str = "", __user__=None, __request__=None) -> str:
        """Apply tracked changes (redlines) to a Word document with custom author name.
    
        change_type: replace (use old_text|||new_text), insert (append text with redline), delete (mark paragraph for deletion)
        author: Name shown in Word's Track Changes (e.g., "Sergio Pedro")
        """
        try:
            import sqlite3 as s3
            conn2 = s3.connect(_DB_PATH)
            row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
            if not row:
                row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
            if not row:
                conn2.close()
                return json.dumps({"error": "File not found"})
            filename = row[0]
            meta = json.loads(row[1]) if row[1] else {}
            fp = meta.get("path", file_id)
            if not os.path.exists(fp):
                fp = os.path.join(_get_owui_data_dir(), "uploads", os.path.basename(fp))
            if not os.path.exists(fp):
                conn2.close()
                return json.dumps({"error": "File not found on disk"})
            with open(fp, "rb") as f:
                data = f.read()
            conn2.close()
    
            from docx import Document
            from docx_revisions import RevisionParagraph
            doc = Document(io.BytesIO(data))
            out_name = output_filename or filename
            results = []
    
            if change_type == "replace":
                parts = content.split("|||", 1)
                if len(parts) != 2:
                    return json.dumps({"error": "Format: old_text|||new_text"})
                find_t, replace_t = parts
                for i, p in enumerate(doc.paragraphs):
                    if paragraph_index >= 0 and i != paragraph_index:
                        continue
                    if find_t in p.text:
                        rp = RevisionParagraph.from_paragraph(p)
                        cnt = rp.replace_tracked(find_t, replace_t, author=author)
                        results.append(f"Para {i}: {cnt} replacements")
            elif change_type == "insert":
                p = doc.add_paragraph()
                rp = RevisionParagraph.from_paragraph(p)
                rp.add_tracked_insertion(content, author=author)
                results.append("Inserted tracked text")
            elif change_type == "delete":
                idx = int(content) if content.isdigit() else paragraph_index
                if idx >= 0 and idx < len(doc.paragraphs):
                    p = doc.paragraphs[idx]
                    rp = RevisionParagraph.from_paragraph(p)
                    rp.add_tracked_deletion(0, len(p.text), author=author)
                    results.append(f"Marked para {idx} for deletion")
    
            out = io.BytesIO()
            doc.save(out)
            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\nTracked changes by '{author}':\n" + "\n".join(results)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})


    async def merge_sheets(self, file_ids: str, output_filename: str = "", __user__=None, __request__=None) -> str:
        try:
            import sqlite3 as s3, openpyxl, io, os
            from copy import copy
            conn2 = s3.connect(_DB_PATH)
            ids = [fid.strip() for fid in file_ids.split(",") if fid.strip()]
            wb_out = openpyxl.Workbook()
            wb_out.remove(wb_out.active)
            merged = 0
            for fid in ids:
                row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (fid,)).fetchone()
                if not row:
                    row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", ("%"+fid+"%",)).fetchone()
                if not row:
                    continue
                filename = row[0]
                meta = json.loads(row[1]) if row[1] else {}
                fp = meta.get("path", fid)
                if not os.path.exists(fp):
                    alt = os.path.join(_get_owui_data_dir(), "uploads", os.path.basename(fp))
                    fp = alt if os.path.exists(alt) else ""
                if not fp or not os.path.exists(fp):
                    continue
                wb_src = openpyxl.load_workbook(io.BytesIO(open(fp,"rb").read()))
                base_name = os.path.splitext(os.path.basename(filename))[0][:15]
                for sn in wb_src.sheetnames:
                    ws_src = wb_src[sn]
                    sheet_name = (base_name + "_" + sn)[:31]
                    ws_out = wb_out.create_sheet(title=sheet_name)
                    for ri, row_data in enumerate(ws_src.iter_rows(), 1):
                        for ci, cell in enumerate(row_data, 1):
                            out_cell = ws_out.cell(row=ri, column=ci, value=cell.value)
                            if cell.has_style:
                                out_cell.font = copy(cell.font)
                                out_cell.fill = copy(cell.fill)
                                out_cell.border = copy(cell.border)
                                out_cell.alignment = copy(cell.alignment)
                                out_cell.number_format = cell.number_format
                    merged += 1
                wb_src.close()
            conn2.close()
            if merged == 0:
                return json.dumps({"error": "No files could be merged"})
            out = io.BytesIO()
            wb_out.save(out)
            out.seek(0)
            fname = output_filename or "merged_workbook.xlsx"
            url, name = await self._save_and_link(out.read(), fname, __request__)
            if url:
                return f"[{name}]({url})\n\nMerged {merged} sheets from {len(ids)} files."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def batch_process(self, file_ids: str, operation: str, params: str = "", output_filename: str = "", __user__=None, __request__=None) -> str:
        try:
            ids = [fid.strip() for fid in file_ids.split(",") if fid.strip()]
            results = []
            for fid in ids:
                if operation == "replace":
                    parts = params.split("|||", 1)
                    if len(parts) == 2:
                        await self.replace_text(fid, parts[0], parts[1], "", __user__, __request__)
                        results.append(f"  {fid}: replaced")
                elif operation == "add_rows":
                    await self.add_content(fid, params, "", __user__, __request__)
                    results.append(f"  {fid}: rows added")
            if results:
                return "Batch processed " + str(len(ids)) + " files:\n" + "\n".join(results)
            return json.dumps({"error": "No files processed"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def auto_backup(self, __user__=None, __request__=None) -> str:
        try:
            import shutil, datetime
            db_path = _DB_PATH
            backup_dir = os.path.join(os.path.expanduser("~"), "open-webui", "backups")
            os.makedirs(backup_dir, exist_ok=True)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"webui_backup_{timestamp}.db"
            backup_path = os.path.join(backup_dir, backup_name)
            shutil.copy2(db_path, backup_path)
            size_kb = os.path.getsize(backup_path) / 1024
            return json.dumps({"success": True, "backup_path": backup_path, "size_kb": round(size_kb,1), "message": f"Backup: {backup_name} ({size_kb:.1f} KB)"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})



    async def merge_pdfs(self, file_ids: str, output_filename: str = "", __user__=None, __request__=None) -> str:
        try:
            import fitz, sqlite3 as s3, io, os
            conn2 = s3.connect(_DB_PATH)
            ids = [fid.strip() for fid in file_ids.split(",") if fid.strip()]
            merger = fitz.open()
            count = 0
            for fid in ids:
                row = conn2.execute("SELECT meta FROM file WHERE id=?", (fid,)).fetchone()
                if not row:
                    row = conn2.execute("SELECT meta FROM file WHERE filename LIKE ?", ("%"+fid+"%",)).fetchone()
                if not row:
                    continue
                meta = json.loads(row[0]) if row[0] else {}
                fp = meta.get("path", fid)
                if not os.path.exists(fp):
                    fp = os.path.join(_get_owui_data_dir(), "uploads", os.path.basename(fp))
                if not os.path.exists(fp):
                    continue
                src = fitz.open(fp)
                merger.insert_pdf(src)
                src.close()
                count += 1
            conn2.close()
            if count == 0:
                merger.close()
                return json.dumps({"error": "No PDFs could be merged"})
            out = io.BytesIO()
            merger.save(out)
            merger.close()
            out.seek(0)
            fname = output_filename or "merged.pdf"
            url, name = await self._save_and_link(out.read(), fname, __request__)
            if url:
                return f"[{name}]({url})\n\nMerged {count} PDFs into one file."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def split_pdf(self, file_id: str, pages_per_file: int = 1, output_filename: str = "", __user__=None, __request__=None) -> str:
        try:
            import fitz, sqlite3 as s3, io, os
            conn2 = s3.connect(_DB_PATH)
            row = conn2.execute("SELECT meta FROM file WHERE id=?", (file_id,)).fetchone()
            if not row:
                row = conn2.execute("SELECT meta FROM file WHERE filename LIKE ?", ("%"+file_id+"%",)).fetchone()
            if not row:
                conn2.close()
                return json.dumps({"error": "File not found"})
            meta = json.loads(row[0]) if row[0] else {}
            fp = meta.get("path", file_id)
            if not os.path.exists(fp):
                fp = os.path.join(_get_owui_data_dir(), "uploads", os.path.basename(fp))
            if not os.path.exists(fp):
                conn2.close()
                return json.dumps({"error": "File not found on disk"})
            conn2.close()
            src = fitz.open(fp)
            total_pages = src.page_count
            urls = []
            for start in range(0, total_pages, pages_per_file):
                end = min(start + pages_per_file, total_pages)
                sub = fitz.open()
                sub.insert_pdf(src, from_page=start, to_page=end-1)
                out = io.BytesIO()
                sub.save(out)
                sub.close()
                out.seek(0)
                part_name = f"part_{start+1}_{end}.pdf"
                url, name = await self._save_and_link(out.read(), part_name, __request__)
                if url:
                    urls.append(f"[{name}]({url})")
            src.close()
            if urls:
                return "Split into " + str(len(urls)) + " files:\n" + "\n".join(urls)
            return json.dumps({"error": "Could not split PDF"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def tool_stats(self, __user__=None, __request__=None) -> str:
        try:
            import sqlite3 as s3
            conn2 = s3.connect(_DB_PATH)
            tool_count = conn2.execute("SELECT COUNT(*) FROM tool WHERE is_active=1").fetchone()[0]
            func_count = conn2.execute("SELECT COUNT(*) FROM function WHERE is_active=1").fetchone()[0]
            model_count = conn2.execute("SELECT COUNT(*) FROM model WHERE is_active=1").fetchone()[0]
            exports_dir = os.path.join(os.path.expanduser("~"), "open-webui", "exports")
            export_count = len([f for f in os.listdir(exports_dir) if os.path.isfile(os.path.join(exports_dir, f))]) if os.path.exists(exports_dir) else 0
            db_size_kb = os.path.getsize(_DB_PATH) / 1024
            conn2.close()
            return json.dumps({
                "tools": tool_count,
                "functions": func_count,
                "models": model_count,
                "exported_files": export_count,
                "db_size_kb": round(db_size_kb, 1)
            }, indent=2)
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    

    async def cleanup_files(self, days_old: int = 30) -> str:
        """Remove generated Office files older than N days. Default 30 days."""
        import time as _time
        
        cutoff = int(_time.time()) - (days_old * 86400)
        
        conn = sqlite3.connect(_DB_PATH)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT id, filename, created_at FROM file WHERE meta LIKE '%office-plugin%' AND created_at < ?",
            (cutoff,)
        ).fetchall()
        
        if not rows:
            conn.close()
            return f"No office-generated files older than {days_old} days found."
        
        deleted = []
        errors = []
        for row in rows:
            try:
                fpath = os.path.join(_UPLOAD_DIR, row["id"])
                if os.path.exists(fpath):
                    os.remove(fpath)
                conn.execute("DELETE FROM file WHERE id = ?", (row["id"],))
                deleted.append(row["filename"])
            except Exception as e:
                errors.append(f"{row['filename']}: {e}")
        
        conn.commit()
        conn.close()
        
        result = f"Cleaned up {len(deleted)} file(s) older than {days_old} days:\n"
        for f in deleted:
            result += f"- {f}\n"
        if errors:
            result += f"\nErrors ({len(errors)}):\n"
            for e in errors:
                result += f"- {e}\n"
        return result

    async def manage_revisions(self, file_id: str, action: str, output_filename: str = "", __user__=None, __request__=None) -> str:
        """List, accept_all or reject_all tracked changes in a Word document."""
        try:
            import sqlite3 as s3
            conn2 = s3.connect(_DB_PATH)
            row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
            if not row:
                row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
            if not row:
                conn2.close()
                return json.dumps({"error": "File not found"})
            filename = row[0]
            meta = json.loads(row[1]) if row[1] else {}
            fp = meta.get("path", file_id)
            if not os.path.exists(fp):
                fp = os.path.join(_get_owui_data_dir(), "uploads", os.path.basename(fp))
            with open(fp, "rb") as f:
                data = f.read()
            conn2.close()
    
            from docx_revisions import RevisionDocument
    
            if action == "list":
                rdoc = RevisionDocument(io.BytesIO(data))
                revs = []
                for para in rdoc.paragraphs:
                    try:
                        rp = RevisionParagraph.from_paragraph(para)
                        if rp.has_track_changes:
                            for ins in rp.insertions:
                                revs.append({"type": "insertion", "author": ins.author, "text": ins.text[:100]})
                            for d in rp.deletions:
                                revs.append({"type": "deletion", "author": d.author, "text": d.text[:100]})
                    except Exception:
                        pass
                return json.dumps({"revisions": revs, "count": len(revs)}, indent=2)
    
            out_name = output_filename or filename
            rdoc = RevisionDocument(io.BytesIO(data))
            if action == "accept_all":
                rdoc.accept_all()
                msg = "All track changes accepted"
            elif action == "reject_all":
                rdoc.reject_all()
                msg = "All track changes rejected"
            else:
                return json.dumps({"error": f"Unknown action: {action}"})
    
            out = io.BytesIO()
            rdoc.save(out)
            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\n{msg}."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

