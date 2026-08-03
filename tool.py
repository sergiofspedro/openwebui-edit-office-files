"""
title: Edit Office Files
author: giofsp
author_url: https://github.com/sergiofspedro
description: Unified tool to read, edit, and create Office files (.xlsx, .xls, .docx, .pptx) preserving original formatting and styles. Supports markdown rendering in DOCX (headings, bold, italic, code, links). Detects highlights, bold, italic formatting. Detects legacy .doc and .ppt. Note: Track changes are not supported.
version: 3.11.3
requirements: openpyxl, python-docx, python-pptx, xlrd, odfpy, docx-revisions, lxml, PyMuPDF, Pillow, pytesseract, qrcode, google-api-python-client, google-auth
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

import base64 as _b64_mod

def _get_owui_data_dir() -> str:
    """Return the Open WebUI data directory for the current OS."""
    data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
    if data_dir:
        return data_dir
    home = os.path.expanduser("~")
    system = platform.system()
    if system == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "open-webui", "data")
    if system == "Darwin":
        return os.path.join(home, "Library", "Application Support", "open-webui", "data")
    # Linux
    xdg_data_home = os.environ.get("XDG_DATA_HOME", "")
    if xdg_data_home:
        return os.path.join(xdg_data_home, "open-webui", "data")
    return os.path.join(home, ".open-webui", "data")

def _get_owui_uploads_dir() -> str:
    """Return the Open WebUI uploads directory for the current OS."""
    data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
    if data_dir:
        return os.path.join(data_dir, "data", "uploads")
    home = os.path.expanduser("~")
    system = platform.system()
    if system == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "open-webui", "data", "uploads")
    if system == "Darwin":
        return os.path.join(home, "Library", "Application Support", "open-webui", "data", "uploads")
    # Linux
    xdg_data_home = os.environ.get("XDG_DATA_HOME", "")
    if xdg_data_home:
        return os.path.join(xdg_data_home, "open-webui", "data", "uploads")
    return os.path.join(home, ".open-webui", "data", "uploads")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
_data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
if _data_dir:
    _DB_PATH = os.path.join(_data_dir, "data", "webui.db")
else:
    _DB_PATH = os.path.join(_get_owui_data_dir(), "webui.db")

_data_dir = os.environ.get("OPEN_WEBUI_DATA_DIR", "")
if _data_dir:
    _UPLOAD_DIR = os.path.join(_data_dir, "data", "uploads")
else:
    _UPLOAD_DIR = _get_owui_uploads_dir()

_EXPORT_DIR = os.environ.get("OWUI_EXPORTS_DIR", os.path.join(os.path.expanduser("~"), "open-webui", "exports"))

# ---------------------------------------------------------------------------
# PPTX namespace constants (used by add_comment)
# ---------------------------------------------------------------------------
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_CM_REL_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsModern"
_CT_MODERN = "application/vnd.ms-office.presentation.commentsModern"
_CT_AUTHORS = "application/vnd.ms-office.presentation.commentsAuthors"

__all__ = [
    "_get_owui_data_dir", "_get_owui_uploads_dir",
    "_data_dir", "_DB_PATH", "_UPLOAD_DIR", "_EXPORT_DIR",
    "_P_NS", "_P14_NS", "_R_NS", "_PKG_REL_NS", "_CT_NS",
    "_CM_REL_TYPE", "_CT_MODERN", "_CT_AUTHORS",
]

def _resolve_file_path(file_id: str) -> Optional[str]:
    """Resolve an Open WebUI file UUID to an absolute disk path."""
    try:
        conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
        try:
            row = conn.execute(
                "SELECT path FROM file WHERE id = ?", (file_id,)
            ).fetchone()
        finally:
            conn.close()
    except Exception as exc:
        print(f"[office] DB lookup failed for {file_id}: {exc}", file=sys.stderr)
        return None

    if not row or not row[0]:
        # Fallback: try by filename
        try:
            conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
            try:
                row = conn.execute(
                    "SELECT path FROM file WHERE filename LIKE ?",
                    (f"%{file_id}%",),
                ).fetchone()
            finally:
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
            if _base and _abs.startswith(os.path.realpath(_base) + os.sep):
                _allowed = True
                break
        if not _allowed:
            print(f"[office] Path traversal blocked: {path}", file=sys.stderr)
            return None
    except Exception as exc:
        print(f"[office] Path traversal check failed for {path}: {exc}", file=sys.stderr)
        return None
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
    if ext in (".csv",):
        return "csv"
    if ext in (".pdf",):
        return "pdf"
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


def _format_text(text: str, mode: str = "format") -> str:
    """Apply consistent text formatting: normalize dashes, sentence case, preserve acronyms.

    Args:
        text: The text to format.
        mode: "format" (default) applies sentence case + acronym preservation + dash normalization.
              "preserve" returns text unchanged.
    """
    if mode == "preserve":
        return text
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
    
    # Smart sentence splitting: avoid splitting on abbreviations
    # Heuristic 1: A word containing an internal period (e.g., "U.S.A.", "e.g.", "i.e.", "Ph.D.")
    #              is an abbreviation — don't split after it (handles U.S.A., U.K., E.U., a.k.a., w.r.t., etc.)
    # Mini-list: short common abbreviations without internal periods (Dr., Mr., St., etc.)
    # Note: month abbreviations (Jan, Feb, etc.) are included since they're commonly followed by
    # a number (e.g., "Jan 15") which would pass the sentence-split check.
    _MINI_ABBREV = {'Dr', 'Mr', 'Mrs', 'Ms', 'Prof', 'Sr', 'Jr', 'St', 'vs', 'etc',
                    'Jan', 'Feb', 'Mar', 'Apr', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'}

    _abbrev_pattern = '|'.join(_re.escape(a) for a in sorted(_MINI_ABBREV, key=len, reverse=True))
    _sentence_re = _re.compile(
        r'(?<!\b(?:' + _abbrev_pattern + r'))'   # not after mini-abbrev
        r'(?<!\.[A-Za-z])'                         # not after word with internal period
        r'(?<=[.!?])\s+'                            # after .!? + space
    )
    sentences = _sentence_re.split(text)

    # Sentence case with intentional capitalization preservation
    # Rule: words with ANY uppercase letter are preserved as-is (proper nouns,
    # acronyms not in the list, product names, the pronoun "I", etc.).
    # Words that are entirely lowercase get sentence-cased.
    formatted_sentences = []
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        
        # Split into words while preserving spacing
        words = s.split(' ')
        formatted_words = []
        
        for word in words:
            if not word:
                formatted_words.append(word)
                continue
            
            # Skip placeholder-protected acronyms (they start with \x00)
            if '\x00' in word:
                formatted_words.append(word)
                continue
            
            # Check if word has intentional capitalization
            # (any uppercase letter means the user intentionally capitalized it)
            has_uppercase = any(c.isupper() for c in word)
            
            if has_uppercase:
                # Preserve original capitalization
                formatted_words.append(word)
            else:
                # Safe to lowercase
                formatted_words.append(word.lower())
        
        s = ' '.join(formatted_words)
        
        # Capitalize first letter of the first word in the sentence
        if s and s[0].isalpha():
            s = s[0].upper() + s[1:]
        
        formatted_sentences.append(s)
    
    text = ' '.join(formatted_sentences)
    
    # Restore acronyms
    for placeholder, acro in acronym_map.items():
        text = text.replace(placeholder, acro)
    
    return text


def _parse_inline_md(text: str) -> list[tuple[str, dict]]:
    """Parse inline markdown in *text* and return segments with formatting metadata.

    Returns a list of ``(plain_text, formatting_dict)`` tuples. The formatting dict
    may contain the keys ``bold``, ``italic``, ``code``, or ``link``.

    Parsing order (priority, first matched wins):
        1. Links — ``[text](url)``
        2. Bold — ``**text**``
        3. Italic — ``*text*``  (conflicts with ``* `` bullet lists are avoided)
        4. Inline code — `` `code` ``
    """
    import re as _re

    if not isinstance(text, str):
        text = str(text) if text is not None else ""

    if text == "":
        return [("", {})]

    # Regexes — link first to prevent its brackets from being consumed by bold/italic
    _link_re = _re.compile(r'\[(.+?)\]\(([^()]*(?:\([^()]*\)[^()]*)*)\)')
    _bold_re = _re.compile(r'\*\*(.+?)\*\*')
    _italic_re = _re.compile(r'(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)')
    _code_re = _re.compile(r'`(.+?)`')

    segments: list[tuple[str, dict]] = []

    # Work on a mutable list of (start, end, kind, data) markers.
    # Kind: 'link', 'bold', 'italic', 'code'
    markers: list[tuple[int, int, str, str]] = []

    # Links
    for m in _link_re.finditer(text):
        markers.append((m.start(), m.end(), 'link', (m.group(1), m.group(2))))  # (label, url)

    # Bold
    for m in _bold_re.finditer(text):
        # Only add if this range doesn't overlap with a higher-priority marker
        markers.append((m.start(), m.end(), 'bold', m.group(1)))

    # Italic
    for m in _italic_re.finditer(text):
        markers.append((m.start(), m.end(), 'italic', m.group(1)))

    # Code
    for m in _code_re.finditer(text):
        markers.append((m.start(), m.end(), 'code', m.group(1)))

    # Sort by start position, then by end position (longer span first wins ties)
    markers.sort(key=lambda x: (x[0], -x[1]))

    # Remove overlapping markers: keep the earliest-starting, and if tied,
    # the longer one; discard any later marker whose start < retained end.
    cleaned: list[tuple[int, int, str, str]] = []
    for mkr in markers:
        if not cleaned or mkr[0] >= cleaned[-1][1]:
            cleaned.append(mkr)
        else:
            # Overlap — keep only if this marker starts at the same position
            # as the last kept one but is longer (shouldn't happen with our
            # sort, but be safe) or is a link that overlaps with something else.
            pass

    markers = cleaned

    pos = 0
    for start, end, kind, data in markers:
        # Plain text before this marker
        if start > pos:
            segments.append((text[pos:start], {}))
        # The formatted segment
        fmt: dict = {}
        if kind == 'link':
            label, url = data
            fmt['link'] = url
            text_part = label
        else:
            if kind == 'bold':
                fmt['bold'] = True
            elif kind == 'italic':
                fmt['italic'] = True
            elif kind == 'code':
                fmt['code'] = True
            text_part = data
        segments.append((text_part, fmt))
        pos = end

    # Trailing plain text
    if pos < len(text):
        segments.append((text[pos:], {}))

    return segments


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
def _add_callout_box(doc, lines, colors, hex_to_rgb, fmt_mode="format"):
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
    elif first_line.startswith("**note") or first_line.startswith("**info"):
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
        text = _format_text(line, mode=fmt_mode)
        segments = _parse_inline_md(text)
        if i == 0:
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(f"{icon} ")
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.name = 'Calibri'
            for seg_text, fmt in segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.name = 'Calibri'
                if fmt.get('italic'):
                    run.font.italic = True
                if fmt.get('code'):
                    run.font.name = 'Consolas'
                    run.font.size = Pt(9)
        else:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            for seg_text, fmt in segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                run.font.size = Pt(10)
                run.font.name = 'Calibri'
                if fmt.get('bold'):
                    run.font.bold = True
                if fmt.get('italic'):
                    run.font.italic = True
                if fmt.get('code'):
                    run.font.name = 'Consolas'
                    run.font.size = Pt(9)

    doc.add_paragraph()


def _add_professional_table(doc, rows, colors, hex_to_rgb, fmt_mode="format"):
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
                cell.text = ''
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                segments = _parse_inline_md(_format_text(cell_text, mode=fmt_mode))
                for seg_text, fmt in segments:
                    if not seg_text:
                        continue
                    run = p.add_run(seg_text)
                    run.font.size = Pt(10)
                    run.font.name = 'Calibri'
                    if i == 0:
                        run.font.bold = True
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)

    doc.add_paragraph()


def _render_content_slide(prs, lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_num, fmt_mode="format"):
    """Render a content slide with professional layout."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, colors["bg"])

    y_pos = 0.5
    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('# ') or line.startswith('## '):
            text = line.lstrip('#').strip()
            add_accent_bar(slide, 0.8, y_pos + 0.15, 0.06, 0.5, colors["accent"])
            add_text_box(slide, 1.1, y_pos, 11, 0.7, text, font_size=28, bold=True, color=colors["text"])
            y_pos += 0.9
        elif line.startswith('- ') or line.startswith('* '):
            text = _format_text(line[2:].strip(), mode=fmt_mode)
            add_text_box(slide, 1.5, y_pos, 10.5, 0.5, "\u2022 " + text, font_size=16, color=colors["text"])
            y_pos += 0.5
        elif line.startswith('|'):
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if any(c.startswith('---') for c in cells):
                continue
            add_text_box(slide, 1.0, y_pos, 11, 0.5, "  |  ".join(cells), font_size=14, color=colors["text"])
            y_pos += 0.4
        else:
            text = _format_text(line, mode=fmt_mode)
            add_text_box(slide, 1.0, y_pos, 11.5, 0.5, text, font_size=16, color=colors["text"])
            y_pos += 0.5

        if y_pos > 6.5:
            break

__all__ = [
    "_office_plugins", "register_office_plugin", "_call_office_plugins",
    "_resolve_file_path", "_read_file_bytes", "_detect_type", "_cell_value",
    "_xls_to_xlsx", "_format_text", "_parse_inline_md", "_encode_filename",
    "_decode_filename", "_read_odf", "_add_callout_box", "_add_professional_table",
    "_render_content_slide",
]

# =========================================================================
# =========================================================================
class Tools:
    class Valves(BaseModel):
        base_url: Optional[str] = Field(
            default=None,
            description="Override the base URL for download links. Auto-detected from X-Original-Host header or WEBUI_URL env var if unset.",
        )
        file_url_pattern: Optional[str] = Field(
            default=None,
            description="Custom file download URL pattern. Use {file_id} placeholder. Examples: /api/v1/files/{file_id}/content (default), /api/files/{file_id}. If unset, uses standard Open WebUI URL.",
        )
        templates: Optional[str] = Field(default="{}", description="JSON map of template names to content strings.")
        cleanup_schedule: Optional[str] = Field(default="{}", description="JSON schedule for auto-cleanup.")
        language: Optional[str] = Field(default="en", description="Language for error messages: en, pt, es, fr, de.")
        pass

    def __init__(self):
        self.valves = self.Valves()

    def _resolve_file(self, file_id: str):
        """Resolve a file ID to (bytes, filename, file_type)."""
        path = _resolve_file_path(file_id)
        if not path:
            return None, None, None
        filename = os.path.basename(path)
        ftype = _detect_type(filename)
        if ftype == "unknown":
            # Files this tool itself creates via _save_and_link are stored on disk under
            # their bare file_id (no extension), so the path's basename never carries a
            # usable extension. Fall back to the DB's filename column, which always does.
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                try:
                    row = conn.execute("SELECT filename FROM file WHERE id = ?", (file_id,)).fetchone()
                finally:
                    conn.close()
                if row and row[0]:
                    filename = row[0]
                    ftype = _detect_type(filename)
            except Exception:
                pass
        file_bytes = _read_file_bytes(file_id)
        return file_bytes, filename, ftype

    def _read_xlsx(self, file_bytes: bytes, filename: str) -> str:
        """Extract text from an XLSX file."""
        from openpyxl import load_workbook
        import io
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        result = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) if c is not None else "" for c in row]
                result.append("\t".join(cells))
        wb.close()
        return "\n\n".join(result)

    def _read_xls(self, file_bytes: bytes, filename: str) -> str:
        """Extract text from an XLS file."""
        import xlrd
        import io
        wb = xlrd.open_workbook(file_contents=file_bytes)
        result = []
        for si in range(wb.nsheets):
            ws = wb.sheet_by_index(si)
            for r in range(ws.nrows):
                cells = [str(ws.cell_value(r, c)) for c in range(ws.ncols)]
                result.append("\t".join(cells))
        return "\n\n".join(result)

    def _read_docx(self, file_bytes: bytes, filename: str) -> str:
        """Extract text from a DOCX file."""
        from docx import Document
        import io
        doc = Document(io.BytesIO(file_bytes))
        result = []
        for p in doc.paragraphs:
            if p.style.name.startswith("Heading"):
                level = p.style.name.replace("Heading ", "")
                result.append(f"{'#' * int(level)} {p.text}")
            else:
                result.append(p.text)
        for t in doc.tables:
            for row in t.rows:
                result.append(" | ".join(cell.text for cell in row.cells))
        return "\n\n".join(result)

    def _read_pptx(self, file_bytes: bytes, filename: str) -> str:
        """Extract text from a PPTX file."""
        from pptx import Presentation
        import io
        prs = Presentation(io.BytesIO(file_bytes))
        result = []
        for si, slide in enumerate(prs.slides, 1):
            result.append(f"# Slide {si}")
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    result.append(shape.text)
        return "\n\n".join(result)

    def _parse_csv_rows(self, content: str) -> list:
        """Parse CSV content into a list of rows with type conversion."""
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
        return parsed_rows

    # -----------------------------------------------------------------
    # Internal: save and return markdown link
    # -----------------------------------------------------------------
    async def _save_and_link(self, file_bytes: bytes, filename: str, __request__=None, __user__=None) -> tuple:
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

            # Generate a preview for the data column
            try:
                if content_type.startswith("text/") or ext in (".csv", ".md", ".txt"):
                    preview = file_bytes.decode('utf-8', errors='replace')[:500]
                elif ext in (".xlsx", ".xls"):
                    preview = f"[Excel spreadsheet: {len(file_bytes)} bytes]"
                elif ext in (".docx",):
                    preview = f"[Word document: {len(file_bytes)} bytes]"
                elif ext in (".pptx",):
                    preview = f"[PowerPoint presentation: {len(file_bytes)} bytes]"
                elif ext in (".pdf",):
                    preview = f"[PDF document: {len(file_bytes)} bytes]"
                else:
                    preview = f"[File: {len(file_bytes)} bytes]"
            except Exception:
                preview = "{}"

            conn = sqlite3.connect(_DB_PATH)
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO file
                       (id, user_id, hash, filename, path, data, meta, created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        file_id,
                        __user__.get("id", "") if __user__ and isinstance(__user__, dict) else "",
                        file_hash,
                        _encode_filename(filename),
                        os.path.join(_UPLOAD_DIR, file_id),
                        preview,
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
            finally:
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

            # Standard Open WebUI file download URL (works for most versions)
            # Customize via file_url_pattern valve if needed
            if self.valves.file_url_pattern:
                url = f"{base_url}{self.valves.file_url_pattern.replace('{file_id}', file_id)}"
            else:
                url = f"{base_url}/api/v1/files/{file_id}/content"
            return (url, filename)

        except Exception as e:
            print(f"[office] Save failed: {e}", file=sys.stderr)
            try:
                if len(file_bytes) > 1_000_000:  # 1MB limit
                    return (None, None)
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
        sheet_name: str = "",
        row_start: int = 1,
        row_end: int = 0,
        __user__=None,
        __request__=None,
    ) -> str:
        """Read any Office file (.xlsx, .xls, .csv, .docx, .pptx) and return its contents as structured JSON.

        Auto-detects the file type from the file ID or filename.
        For xlsx/xls/csv: returns sheets with headers and rows.
        For docx: returns paragraphs with styles and tables.
        For pptx: returns slides with shapes and text.
        Legacy .doc and .ppt formats return a helpful error message.

        Args:
            file_id: The Open WebUI file ID (UUID) or filename
            max_rows: Maximum rows to return (default 500)
            sheet_name: Optional - read only this sheet (xlsx/xls only)
            row_start: Starting row (1-indexed, default 1)
            row_end: Ending row (0 = use max_rows limit)
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
                sheetnames = [sn for sn in wb.sheetnames] if not sheet_name else [sn for sn in wb.sheetnames if sn == sheet_name]
                if sheet_name and sheet_name not in wb.sheetnames:
                    result["warning"] = f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}"
                for sn in sheetnames:
                    ws = wb[sn]
                    sheet: Dict[str, Any] = {
                        "name": sn,
                        "headers": [],
                        "rows": [],
                        "total_rows": ws.max_row or 0,
                        "total_cols": ws.max_column or 0,
                    }
                    r_start = max(1, row_start)
                    r_end = row_end if row_end > 0 else (ws.max_row or 0)
                    r_end = min(r_end, r_start + max_rows - 1) if max_rows else r_end
                    for ri, row in enumerate(ws.iter_rows(min_row=r_start, max_row=r_end), r_start):
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
                sheet_indices = range(xls_book.nsheets)
                if sheet_name:
                    sheet_indices = [i for i in range(xls_book.nsheets) if xls_book.sheet_by_index(i).name == sheet_name]
                    if not sheet_indices:
                        all_names = [xls_book.sheet_by_index(i).name for i in range(xls_book.nsheets)]
                        result["warning"] = f"Sheet '{sheet_name}' not found. Available sheets: {all_names}"
                for sheet_idx in sheet_indices:
                    xls_sheet = xls_book.sheet_by_index(sheet_idx)
                    sheet: Dict[str, Any] = {
                        "name": xls_sheet.name,
                        "headers": [],
                        "rows": [],
                        "total_rows": xls_sheet.nrows,
                        "total_cols": xls_sheet.ncols,
                    }
                    r_start = max(0, row_start - 1)
                    r_end = row_end if row_end > 0 else xls_sheet.nrows
                    r_end = min(r_end, r_start + max_rows) if max_rows else r_end
                    r_end = min(r_end, xls_sheet.nrows)
                    for rx in range(r_start, r_end):
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

            elif file_type == "csv":
                import csv as _csv_mod
                result["sheets"] = []
                d_bytes = file_data
                try:
                    text = d_bytes.decode("utf-8-sig")
                except Exception:
                    text = d_bytes.decode("utf-8", errors="replace")
                reader = _csv_mod.DictReader(io.StringIO(text))
                headers = reader.fieldnames or []
                rows = []
                for ri, row in enumerate(reader):
                    if max_rows and ri >= max_rows:
                        break
                    rows.append(dict(row))
                base_name = os.path.splitext(os.path.basename(filename))[0] if filename else "CSV"
                sheet = {
                    "name": base_name,
                    "headers": [str(h) for h in headers],
                    "rows": [[r.get(h, "") for h in headers] for r in rows],
                    "total_rows": len(rows),
                    "total_cols": len(headers),
                }
                result["sheets"].append(sheet)

            elif file_type == "docx":
                from docx import Document
                from docx.shared import Pt
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
        For documents (docx): content is text to append at the end. Preserves original
            formatting and capitalization of the existing document.
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
                parsed_rows = self._parse_csv_rows(content)

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
                parsed_rows = self._parse_csv_rows(content)

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
                    line = line.strip()
                    if not line:
                        continue
                    if line.startswith('# ') or line.startswith('## ') or line.startswith('### '):
                        level = line.count('#')
                        text = line.lstrip('#').strip()
                        doc.add_heading(text, level=min(level, 3))
                    elif line.startswith('- ') or line.startswith('* '):
                        text = line[2:].strip()
                        doc.add_paragraph(text, style='List Bullet')
                    else:
                        segments = _parse_inline_md(line)
                        p = doc.add_paragraph(style=last_style)
                        for seg_text, fmt in segments:
                            if not seg_text:
                                continue
                            run = p.add_run(seg_text)
                            if fmt.get('code'):
                                run.font.name = "Consolas"
                                run.font.size = Pt(9)
                            if fmt.get('bold'):
                                run.font.bold = True
                            if fmt.get('italic'):
                                run.font.italic = True

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
                    tf.text = _format_text(title, mode="format")
                    p = tf.paragraphs[0]
                    p.font.size = Inches(0.6)
                    p.font.bold = True

                    # Add body text
                    if body:
                        txBox2 = slide.shapes.add_textbox(
                            Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
                        )
                        tf2 = txBox2.text_frame
                        tf2.text = _format_text(body, mode="format")
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
        """Find and replace text in any Office file while preserving original formatting
        and capitalization of the existing document.

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
                        if para.runs:
                            # Preserve formatting: replace within each run individually
                            for run in para.runs:
                                if find_text in run.text:
                                    run.text = run.text.replace(find_text, replace_with)
                            count += 1
                        else:
                            para.text = para.text.replace(find_text, replace_with)
                            count += 1

                # Also replace in tables
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if find_text in cell.text:
                                for para in cell.paragraphs:
                                    if find_text in para.text:
                                        if para.runs:
                                            for run in para.runs:
                                                if find_text in run.text:
                                                    run.text = run.text.replace(find_text, replace_with)
                                            count += 1
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
                                        for run in para.runs:
                                            if find_text in run.text:
                                                run.text = run.text.replace(find_text, replace_with)
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
    # UPDATE CELLS
    # -----------------------------------------------------------------
    async def update_cells(
        self,
        file_id: str,
        cells: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Update individual cells in an XLSX file. Pass a JSON array of cell updates.

        Each update: {"cell": "A1", "value": "new value"} or {"cell": "B2", "value": 42, "sheet": "Sheet1"}
        If sheet is omitted, uses the active sheet.

        Args:
            file_id: The file ID of the XLSX file to edit
            cells: JSON string - array of {cell, value[, sheet]} objects
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            try:
                conn2 = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row2 = conn2.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn2.close()
                filename = row2[0] if row2 else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            if file_type not in ("xlsx", "xls"):
                return json.dumps({"error": f"Unsupported type: {file_type}. Only xlsx/xls supported."})

            import openpyxl
            from openpyxl.utils import column_index_from_string
            import re as _re_cell

            try:
                updates = json.loads(cells)
                if not isinstance(updates, list):
                    updates = [updates]
            except json.JSONDecodeError:
                return json.dumps({"error": "cells must be valid JSON (array of {cell, value} objects)"})

            out_name = output_filename or os.path.splitext(filename)[0] + "_updated.xlsx"
            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            count = 0
            for upd in updates:
                cell_ref = upd.get("cell", "")
                value = upd.get("value", "")
                sname = upd.get("sheet", "")
                ws = wb[sname] if sname else wb.active
                m = _re_cell.match(r"([A-Za-z]+)(\d+)", cell_ref)
                if not m:
                    continue
                col_str, row_str = m.group(1), m.group(2)
                col_idx = column_index_from_string(col_str.upper())
                row_idx = int(row_str)
                ws.cell(row=row_idx, column=col_idx).value = value
                count += 1

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\nUpdated {count} cell(s) in {file_type.upper()} file."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # MODIFY ROWS (insert/delete)
    # -----------------------------------------------------------------
    async def modify_rows(
        self,
        file_id: str,
        action: str,
        row_number: int = 1,
        count: int = 1,
        sheet_name: str = "",
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Insert or delete rows in an XLSX file.

        Args:
            file_id: The file ID of the XLSX file to edit
            action: 'insert' or 'delete'
            row_number: Row number to start at (1-indexed)
            count: Number of rows to insert/delete (default 1)
            sheet_name: Optional sheet name (default: active sheet)
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            try:
                conn2 = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row2 = conn2.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn2.close()
                filename = row2[0] if row2 else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            if file_type not in ("xlsx", "xls"):
                return json.dumps({"error": f"Unsupported type: {file_type}. Only xlsx/xls supported."})

            import openpyxl

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            ws = wb[sheet_name] if sheet_name else wb.active
            out_name = output_filename or os.path.splitext(filename)[0] + "_modified.xlsx"

            action_lower = action.strip().lower()
            if action_lower == "insert":
                ws.insert_rows(idx=row_number, amount=count)
                msg = f"Inserted {count} row(s) at row {row_number}"
            elif action_lower == "delete":
                ws.delete_rows(idx=row_number, amount=count)
                msg = f"Deleted {count} row(s) starting at row {row_number}"
            else:
                return json.dumps({"error": f"Unknown action '{action}'. Use 'insert' or 'delete'."})

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            url, name = await self._save_and_link(out.read(), out_name, __request__)
            if url:
                return f"[{name}]({url})\n\n{msg} in {file_type.upper()} file."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # PASSWORD PROTECT FILE
    # -----------------------------------------------------------------
    async def protect_file(
        self,
        file_id: str,
        password: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Add password protection to an XLSX or DOCX file.

        For XLSX: uses openpyxl workbook protection. For DOCX: uses write protection.
        Install msoffcrypto-tool (pip install msoffcrypto-tool) for encryption support.

        Args:
            file_id: The file ID of the Office file to protect
            password: Password to set
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            try:
                conn2 = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row2 = conn2.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn2.close()
                filename = row2[0] if row2 else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            if file_type not in ("xlsx", "docx"):
                return json.dumps({"error": f"Unsupported type: {file_type}. Only xlsx and docx supported."})

            import hashlib

            if file_type == "xlsx":
                import openpyxl
                from openpyxl.workbook.protection import WorkbookProtection

                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                out_name = output_filename or os.path.splitext(filename)[0] + "_protected.xlsx"

                # Set workbook protection
                wb.security = WorkbookProtection(workbookPassword=password, lockStructure=True)

                # Protect all sheets
                for ws in wb.worksheets:
                    ws.protection.set_password(password)

                out = io.BytesIO()
                wb.save(out)
                wb.close()
                out.seek(0)
                protected_bytes = out.read()

            elif file_type == "docx":
                from docx import Document
                from docx.oxml.ns import qn
                from lxml import etree

                out_name = output_filename or os.path.splitext(filename)[0] + "_protected.docx"
                doc = Document(io.BytesIO(file_data))

                # Add write protection XML
                settings_element = doc.settings.element
                protection = etree.SubElement(
                    settings_element,
                    qn("w:documentProtection")
                )
                protection.set(qn("w:edit"), "readOnly")
                protection.set(qn("w:enforcement"), "1")
                pw_hash = hashlib.sha256(password.encode("utf-8")).hexdigest().upper()
                protection.set(qn("w:cryptProviderType"), "rsaAES")
                protection.set(qn("w:cryptAlgorithmClass"), "hash")
                protection.set(qn("w:cryptAlgorithmType"), "typeAny")
                protection.set(qn("w:cryptAlgorithmSid"), "4")
                protection.set(qn("w:cryptSpinCount"), "100000")
                protection.set(qn("w:hash"), pw_hash)
                protection.set(qn("w:salt"), "AAAAAAAAAAAAAAAAAAAAAA==")

                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                protected_bytes = out.read()

            url, fname = await self._save_and_link(protected_bytes, out_name, __request__)
            if url:
                return f"[{fname}]({url})\n\nPassword-protected {file_type.upper()} file created."
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
        raw_text: bool = False,
        __user__=None,
        __request__=None,
    ) -> str:
        """Create a new Office file from scratch.

        For xlsx: content is CSV with headers on first line.
        For docx: supports markdown formatting — headings (#, ##, ###), bullets (-, *),
            and inline markdown (**bold**, *italic*, `code`).
        For pptx: each line defines a slide. Use "---" as separator between slides.

        Args:
            file_type: 'xlsx', 'docx', or 'pptx'
            content: Content specification
            output_filename: Output filename
            raw_text: If True, skip text formatting
        """
        fmt_mode = "preserve" if raw_text else "format"
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
                from docx.shared import Pt, RGBColor
                doc = Document()
                for raw_line in content.split("\n"):
                    line = raw_line
                    stripped = line.strip()

                    if stripped == "":
                        # Empty line → paragraph break
                        doc.add_paragraph("")
                        continue

                    # Heading detection (must start with # markers)
                    if stripped.startswith("### "):
                        text = stripped[4:]
                        doc.add_heading(text, level=3)
                        continue
                    if stripped.startswith("## "):
                        text = stripped[3:]
                        doc.add_heading(text, level=2)
                        continue
                    if stripped.startswith("# "):
                        text = stripped[2:]
                        doc.add_heading(text, level=1)
                        continue

                    # Bullet detection
                    if stripped.startswith("- ") or stripped.startswith("* "):
                        text = _format_text(stripped[2:], mode=fmt_mode)
                        doc.add_paragraph(text, style='List Bullet')
                        continue

                    # Body text — rich formatting via inline markdown
                    text = _format_text(line, mode=fmt_mode)
                    segments = _parse_inline_md(text)
                    p = doc.add_paragraph()
                    for seg_text, fmt in segments:
                        if not seg_text:
                            continue
                        run = p.add_run(seg_text)
                        if fmt.get('code'):
                            run.font.name = "Consolas"
                            run.font.size = Pt(9)
                        else:
                            run.font.size = Pt(11)
                        if fmt.get('bold'):
                            run.font.bold = True
                        if fmt.get('italic'):
                            run.font.italic = True
                        if fmt.get('link'):
                            run.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
                            run.font.underline = True
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



    async def generate_document(self, content: str, title: str = "Document", theme: str = "professional", typography: str = "modern", raw_text: bool = False, __user__=None, __request__=None) -> str:
        """Generate a professional Word document with modern styling, emojis, cards, and visual elements.

        Headings preserve their original capitalization (no forced sentence case).
        Body text uses sentence case (first letter of each sentence capitalized, acronyms preserved).
        Supports inline markdown: **bold**, *italic*, `code`, [links](url), and ```code blocks```.

        Args:
            content: Markdown content to convert to a document
            title: Document title / filename
            theme: Visual theme - professional, modern, creative, corporate, minimal, elegant, ocean, sunset, forest, midnight
            typography: Font preset - modern, classic, serif, sans
            raw_text: If True, skip all text formatting — text is used exactly as provided.
        """
        from docx import Document
        from docx.shared import Inches, Pt, Cm, RGBColor, Emu
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import qn, nsdecls
        from docx.oxml import parse_xml
        from lxml import etree
        import datetime, re as _re
        fmt_mode = "preserve" if raw_text else "format"
        
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        
        # --- Color Palettes (10+) ---
        palettes = {
            "professional": {"primary": "1F4E79", "accent": "2E75B6", "light": "D6E4F0", "text": "333333", "bg": "FFFFFF", "success": "27AE60", "warning": "F39C12", "danger": "E74C3C", "info": "3498DB"},
            "modern": {"primary": "2D3436", "accent": "6C5CE7", "light": "DFE6E9", "text": "2D3436", "bg": "FFFFFF", "success": "00B894", "warning": "FDCB6E", "danger": "FF7675", "info": "74B9FF"},
            "creative": {"primary": "E17055", "accent": "FDCB6E", "light": "FFF3E0", "text": "2D3436", "bg": "FFFAF5", "success": "00B894", "warning": "FDCB6E", "danger": "D63031", "info": "6C5CE7"},
            "corporate": {"primary": "003366", "accent": "CC0000", "light": "E8EEF4", "text": "1A1A1A", "bg": "FFFFFF", "success": "006633", "warning": "FF6600", "danger": "CC0000", "info": "0066CC"},
            "minimal": {"primary": "000000", "accent": "666666", "light": "F5F5F5", "text": "333333", "bg": "FAFAFA", "success": "333333", "warning": "666666", "danger": "000000", "info": "999999"},
            "elegant": {"primary": "4A235A", "accent": "8E44AD", "light": "F3E5F5", "text": "1A1A1A", "bg": "FDFBF7", "success": "27AE60", "warning": "D4AC0D", "danger": "C0392B", "info": "2980B9"},
            "ocean": {"primary": "0A3D62", "accent": "38ADA9", "light": "D1F2EB", "text": "1E272E", "bg": "F8FFFE", "success": "079992", "warning": "F6B93B", "danger": "E55039", "info": "3C6382"},
            "sunset": {"primary": "B33771", "accent": "FD7272", "light": "FFE4E4", "text": "2C3A47", "bg": "FFFBF5", "success": "58B19F", "warning": "F8B500", "danger": "E66767", "info": "786FA6"},
            "forest": {"primary": "1B4332", "accent": "40916C", "light": "D8F3DC", "text": "1A1A1A", "bg": "F7FFF7", "success": "2D6A4F", "warning": "B7B73F", "danger": "D00000", "info": "52B788"},
            "midnight": {"primary": "0F172A", "accent": "38BDF8", "light": "1E293B", "text": "F8FAFC", "bg": "0F172A", "success": "34D399", "warning": "FBBF24", "danger": "F87171", "info": "60A5FA"},
        }
        colors = palettes.get(theme, palettes["professional"])
        
        # --- Typography Presets ---
        fonts = {
            "modern": {"heading": "Calibri", "body": "Calibri"},
            "classic": {"heading": "Georgia", "body": "Calibri"},
            "serif": {"heading": "Cambria", "body": "Calibri"},
            "sans": {"heading": "Arial", "body": "Arial"},
        }
        font_pair = fonts.get(typography, fonts["modern"])
        
        def hex_to_rgb(h):
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        
        def add_colored_bar(doc, color_hex, height_pt=4):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(" ")
            run.font.size = Pt(height_pt)
            pPr = p._p.get_or_add_pPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
            pPr.append(shd)
        
        def add_card_box(doc, lines, border_color, bg_color, icon="", title=""):
            table = doc.add_table(rows=1, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = table.rows[0].cells[0]
            cell.width = Inches(6.0)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}" w:val="clear"/>')
            tcPr.append(shd)
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:left w:val="single" w:sz="24" w:space="8" w:color="{border_color}"/>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(borders)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            if icon or title:
                header_text = f"{icon} {title}" if icon and title else (icon or title)
                run = p.add_run(_format_text(header_text, mode=fmt_mode))
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.name = font_pair["heading"]
                run.font.color.rgb = hex_to_rgb(border_color)
            for line in lines:
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                text = _format_text(line, mode=fmt_mode)
                segments = _parse_inline_md(text)
                for seg_text, fmt in segments:
                    if not seg_text:
                        continue
                    run = p.add_run(seg_text)
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
                    else:
                        run.font.size = Pt(10)
                        run.font.name = font_pair["body"]
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    run.font.color.rgb = hex_to_rgb(colors["text"])
                    if fmt.get('link'):
                        run.font.color.rgb = hex_to_rgb("2563EB")
                        run.font.underline = True
            doc.add_paragraph()
        
        def add_kpi_card(doc, value, label, color_hex, icon=""):
            table = doc.add_table(rows=1, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = table.rows[0].cells[0]
            cell.width = Inches(2.8)
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["light"]}" w:val="clear"/>')
            tcPr.append(shd)
            borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="12" w:space="0" w:color="{color_hex}"/><w:left w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/><w:right w:val="single" w:sz="4" w:space="0" w:color="DDDDDD"/></w:tcBorders>')
            tcPr.append(borders)
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(f"{icon} {value}" if icon else value)
            run.font.size = Pt(24); run.font.bold = True
            run.font.color.rgb = hex_to_rgb(color_hex); run.font.name = font_pair["heading"]
            p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_after = Pt(12)
            label_text = _format_text(label, mode=fmt_mode)
            label_segments = _parse_inline_md(label_text)
            for seg_text, fmt in label_segments:
                if not seg_text:
                    continue
                run2 = p2.add_run(seg_text)
                run2.font.size = Pt(9)
                run2.font.color.rgb = hex_to_rgb(colors["text"])
                run2.font.name = font_pair["body"]
                if fmt.get('bold'):
                    run2.font.bold = True
                if fmt.get('italic'):
                    run2.font.italic = True
                if fmt.get('code'):
                    run2.font.name = "Consolas"
                    run2.font.size = Pt(9)
        
        def add_progress_bar(doc, label, percentage, color_hex):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            label_text = _format_text(label, mode=fmt_mode)
            label_segments = _parse_inline_md(label_text)
            for seg_text, fmt in label_segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                run.font.size = Pt(10); run.font.bold = True
                run.font.name = font_pair["body"]; run.font.color.rgb = hex_to_rgb(colors["text"])
                if fmt.get('italic'):
                    run.font.italic = True
                if fmt.get('link'):
                    run.font.color.rgb = hex_to_rgb("2563EB")
                    run.font.underline = True
                if fmt.get('code'):
                    run.font.name = "Consolas"
                    run.font.size = Pt(9)
            run = p.add_run(f": {percentage}%")
            run.font.size = Pt(10); run.font.bold = True
            run.font.name = font_pair["body"]; run.font.color.rgb = hex_to_rgb(colors["text"])
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            filled = table.rows[0].cells[0]; empty = table.rows[0].cells[1]
            filled.width = Inches(5.0 * percentage / 100)
            empty.width = Inches(5.0 * (100 - percentage) / 100)
            tcF = filled._tc; tcPrF = tcF.get_or_add_tcPr()
            shdF = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
            tcPrF.append(shdF)
            tcE = empty._tc; tcPrE = tcE.get_or_add_tcPr()
            shdE = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E0E0E0" w:val="clear"/>')
            tcPrE.append(shdE)
            pF = filled.paragraphs[0]; runF = pF.add_run(" "); runF.font.size = Pt(6)
            pE = empty.paragraphs[0]; runE = pE.add_run(" "); runE.font.size = Pt(6)
        
        def add_step_guide(doc, steps):
            for i, step in enumerate(steps, 1):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(f"  {i}  ")
                run.font.size = Pt(14); run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255); run.font.name = font_pair["heading"]
                rPr = run._r.get_or_add_rPr()
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["primary"]}" w:val="clear"/>')
                rPr.append(shd)
                text = _format_text(step, mode=fmt_mode)
                segments = _parse_inline_md(text)
                first = True
                for seg_text, fmt in segments:
                    if not seg_text:
                        continue
                    prefix = "  " if first else ""
                    first = False
                    run = p.add_run(prefix + seg_text)
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
                    else:
                        run.font.size = Pt(11)
                        run.font.name = font_pair["body"]
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    run.font.color.rgb = hex_to_rgb(colors["text"])
                    if fmt.get('link'):
                        run.font.color.rgb = hex_to_rgb("2563EB")
                        run.font.underline = True
        
        def add_pull_quote(doc, text, author=""):
            table = doc.add_table(rows=1, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = table.rows[0].cells[0]; cell.width = Inches(5.5)
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["light"]}" w:val="clear"/>')
            tcPr.append(shd)
            borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="36" w:space="12" w:color="{colors["accent"]}"/><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tcBorders>')
            tcPr.append(borders)
            p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            text_formatted = _format_text(text, mode=fmt_mode)
            segments = _parse_inline_md(text_formatted)
            # Opening quote
            run = p.add_run('"')
            run.font.size = Pt(14); run.font.italic = True
            run.font.name = font_pair["heading"]; run.font.color.rgb = hex_to_rgb(colors["text"])
            for seg_text, fmt in segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                run.font.size = Pt(14); run.font.italic = True
                run.font.name = font_pair["heading"]
                run.font.color.rgb = hex_to_rgb(colors["text"])
                if fmt.get('bold'):
                    run.font.bold = True
                if fmt.get('link'):
                    run.font.color.rgb = hex_to_rgb("2563EB")
                    run.font.underline = True
                if fmt.get('code'):
                    run.font.name = "Consolas"
                    run.font.size = Pt(9)
            # Closing quote
            run = p.add_run('"')
            run.font.size = Pt(14); run.font.italic = True
            run.font.name = font_pair["heading"]; run.font.color.rgb = hex_to_rgb(colors["text"])
            if author:
                p2 = cell.add_paragraph(); p2.paragraph_format.space_after = Pt(12)
                p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                author_text = _format_text(author, mode=fmt_mode)
                author_segments = _parse_inline_md(author_text)
                first = True
                for seg_text, fmt in author_segments:
                    if not seg_text:
                        continue
                    prefix = "\u2014 " if first else ""
                    first = False
                    run = p2.add_run(prefix + seg_text)
                    run.font.size = Pt(10); run.font.name = font_pair["body"]
                    run.font.color.rgb = hex_to_rgb(colors["accent"])
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    if fmt.get('link'):
                        run.font.color.rgb = hex_to_rgb("2563EB")
                        run.font.underline = True
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
            doc.add_paragraph()
        
        def add_comparison_table(doc, headers, rows):
            table = doc.add_table(rows=len(rows)+1, cols=len(headers))
            table.style = 'Colorful Grid Accent 1'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for j, h in enumerate(headers):
                cell = table.rows[0].cells[j]
                cell.text = ''
                p = cell.paragraphs[0]
                segments = _parse_inline_md(_format_text(h, mode=fmt_mode))
                for seg_text, fmt in segments:
                    if not seg_text:
                        continue
                    r = p.add_run(seg_text)
                    r.font.size = Pt(10); r.font.bold = True; r.font.name = font_pair["heading"]
                    if fmt.get('italic'):
                        r.font.italic = True
                    if fmt.get('code'):
                        r.font.name = "Consolas"
                        r.font.size = Pt(9)
            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    if j < len(headers):
                        cell = table.rows[i+1].cells[j]
                        display = val
                        if val.lower() in ("yes","true","sim","si","oui","ja"): display = "\u2705 " + val
                        elif val.lower() in ("no","false","nao","non","nein"): display = "\u274c " + val
                        elif val.lower() in ("partial","maybe","talvez"): display = "\u26a0\ufe0f " + val
                        cell.text = ''
                        p = cell.paragraphs[0]
                        segments = _parse_inline_md(_format_text(display, mode=fmt_mode))
                        for seg_text, fmt in segments:
                            if not seg_text:
                                continue
                            r = p.add_run(seg_text)
                            r.font.size = Pt(10); r.font.name = font_pair["body"]
                            if fmt.get('bold'):
                                r.font.bold = True
                            if fmt.get('italic'):
                                r.font.italic = True
                            if fmt.get('code'):
                                r.font.name = "Consolas"
                                r.font.size = Pt(9)
            doc.add_paragraph()
        
        def add_timeline(doc, events):
            for i, (date, desc) in enumerate(events):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                icon = "\u25cf" if i == 0 else "\u25cb"
                run = p.add_run(f"  {icon}  ")
                run.font.size = Pt(12); run.font.color.rgb = hex_to_rgb(colors["accent"])
                run.font.name = font_pair["heading"]
                date_text = _format_text(date, mode=fmt_mode)
                date_segments = _parse_inline_md(date_text)
                for seg_text, fmt in date_segments:
                    if not seg_text:
                        continue
                    run = p.add_run(seg_text)
                    run.font.size = Pt(10); run.font.bold = True
                    run.font.name = font_pair["body"]; run.font.color.rgb = hex_to_rgb(colors["primary"])
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
                    if fmt.get('link'):
                        run.font.color.rgb = hex_to_rgb("2563EB")
                        run.font.underline = True
                run_sep = p.add_run("  ")
                run_sep.font.size = Pt(10)
                run_sep.font.name = font_pair["body"]
                desc_text = _format_text(desc, mode=fmt_mode)
                desc_segments = _parse_inline_md(desc_text)
                for seg_text, fmt in desc_segments:
                    if not seg_text:
                        continue
                    run = p.add_run(seg_text)
                    run.font.size = Pt(10); run.font.name = font_pair["body"]
                    run.font.color.rgb = hex_to_rgb(colors["text"])
                    if fmt.get('bold'):
                        run.font.bold = True
                    if fmt.get('italic'):
                        run.font.italic = True
                    if fmt.get('code'):
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
                    if fmt.get('link'):
                        run.font.color.rgb = hex_to_rgb("2563EB")
                        run.font.underline = True
        
        def add_status_badge(doc, text, status="info"):
            colors_map = {"success": colors["success"], "warning": colors["warning"], "danger": colors["danger"], "info": colors["info"]}
            icons = {"success": "\U0001f7e2", "warning": "\U0001f7e1", "danger": "\U0001f534", "info": "\U0001f535"}
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            badge_color = hex_to_rgb(colors_map.get(status, colors["info"]))
            icon_text = icons.get(status, icons['info'])
            run = p.add_run(f" {icon_text} ")
            run.font.size = Pt(10); run.font.bold = True
            run.font.name = font_pair["body"]; run.font.color.rgb = badge_color
            badge_text = _format_text(text, mode=fmt_mode)
            badge_segments = _parse_inline_md(badge_text)
            for seg_text, fmt in badge_segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                run.font.size = Pt(10); run.font.bold = True
                run.font.name = font_pair["body"]; run.font.color.rgb = badge_color
                if fmt.get('italic'):
                    run.font.italic = True
                if fmt.get('link'):
                    run.font.color.rgb = hex_to_rgb("2563EB")
                    run.font.underline = True
                if fmt.get('code'):
                    run.font.name = "Consolas"
                    run.font.size = Pt(9)
        
        def add_visual_separator(doc, style="dots"):
            separators = {"dots": "\u25cf \u25cf \u25cf", "line": "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501", "dash": "\u2500 \u2500 \u2500 \u2500 \u2500 \u2500 \u2500 \u2500 \u2500 \u2500"}
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)
            run = p.add_run(separators.get(style, separators["dots"]))
            run.font.size = Pt(8); run.font.color.rgb = hex_to_rgb(colors["accent"])
        
        def _add_rich_paragraph(doc, text, style=None, font_name="Calibri", font_size=11, color=None):
            """Add a paragraph with rich formatting from inline markdown parsing."""
            segments = _parse_inline_md(text)
            if style and style.startswith('Heading'):
                try:
                    level = int(style.split()[-1])
                except (ValueError, IndexError):
                    level = 1
                p = doc.add_heading('', level=level)
                for run in p.runs:
                    run.text = ''
            else:
                p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
            for seg_text, fmt in segments:
                if not seg_text:
                    continue
                run = p.add_run(seg_text)
                if fmt.get('code'):
                    run.font.name = "Consolas"
                    run.font.size = Pt(9)
                else:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                if fmt.get('bold'):
                    run.font.bold = True
                if fmt.get('italic'):
                    run.font.italic = True
                if color:
                    if isinstance(color, str):
                        run.font.color.rgb = hex_to_rgb(color)
                    else:
                        run.font.color.rgb = color
                if fmt.get('link'):
                    run.font.color.rgb = hex_to_rgb("2563EB")
                    run.font.underline = True
            return p
        
        # --- Default Style ---
        style = doc.styles['Normal']
        font = style.font; font.name = font_pair["body"]; font.size = Pt(11)
        font.color.rgb = hex_to_rgb(colors["text"])
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.line_spacing = 1.15
        
        for i in range(1, 4):
            hs = doc.styles[f'Heading {i}']
            hs.font.name = font_pair["heading"]
            hs.font.color.rgb = hex_to_rgb(colors["primary"])
            sizes = {1: 24, 2: 18, 3: 14}
            hs.font.size = Pt(sizes.get(i, 14))
            spaces = {1: (24, 12), 2: (18, 8), 3: (12, 6)}
            hs.paragraph_format.space_before = Pt(spaces[i][0])
            hs.paragraph_format.space_after = Pt(spaces[i][1])
        
        # --- Cover Page ---
        add_colored_bar(doc, colors["primary"], 4)
        cover_table = doc.add_table(rows=1, cols=1)
        cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = cover_table.rows[0].cells[0]; cell.width = Inches(6.5)
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["primary"]}" w:val="clear"/>')
        tcPr.append(shd)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(40); p.paragraph_format.space_after = Pt(8)
        run = p.add_run(title)
        run.font.size = Pt(28); run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255); run.font.name = font_pair["heading"]
        p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(30)
        run2 = p2.add_run(datetime.datetime.now().strftime("%B %d, %Y"))
        run2.font.size = Pt(12); run2.font.color.rgb = RGBColor(200, 200, 200)
        run2.font.name = font_pair["body"]
        doc.add_paragraph()
        
        # --- Auto TOC ---
        toc_added = False
        
        # --- Process Content ---
        lines = content.split('\n')
        in_card = False; card_lines = []; card_icon = ""; card_title = ""
        in_kpi = False; kpi_data = []
        in_timeline = False; timeline_data = []
        in_steps = False; step_lines = []
        in_quote = False; quote_text = ""; quote_author = ""
        in_comparison = False; comp_headers = []; comp_rows = []
        in_progress = False; progress_data = []
        in_code_block = False; code_lines = []
        _timeline_re = _re.compile(r'^(\d{4}|\d{1,2}[/-]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*)\s*[-:]\s*(.+)$', _re.IGNORECASE)
        
        for line in lines:
            line = line.strip()
            if not line:
                if in_code_block:
                    code_lines.append("")
                    continue
                if in_card:
                    add_card_box(doc, card_lines, colors["accent"], colors["light"], card_icon, card_title)
                    card_lines = []; in_card = False
                if in_steps:
                    if len(step_lines) >= 3:
                        add_step_guide(doc, step_lines)
                    else:
                        for i, step_text in enumerate(step_lines, 1):
                            _add_rich_paragraph(doc, f"{i}. {_format_text(step_text, mode=fmt_mode)}", font_name=font_pair["body"], font_size=11, color=colors["text"])
                    step_lines = []; in_steps = False
                if in_timeline:
                    add_timeline(doc, timeline_data)
                    timeline_data = []; in_timeline = False
                if in_comparison:
                    add_comparison_table(doc, comp_headers, comp_rows)
                    comp_headers = []; comp_rows = []; in_comparison = False
                if in_progress:
                    for lbl, pct, clr in progress_data:
                        add_progress_bar(doc, lbl, pct, clr)
                    progress_data = []; in_progress = False
                continue
            
            # Code block: ``` ... ```
            if line.startswith('```'):
                if not in_code_block:
                    in_code_block = True
                    code_lines = []
                else:
                    for code_line in code_lines:
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        run = p.add_run(code_line)
                        run.font.name = "Consolas"
                        run.font.size = Pt(9)
                        run.font.color.rgb = hex_to_rgb(colors["text"])
                    in_code_block = False
                    code_lines = []
                continue
            if in_code_block:
                code_lines.append(line)
                continue
            
            # Card: > 📊 **Title** or > content
            if line.startswith('> ') and not in_card and not in_kpi and not in_timeline and not in_steps and not in_quote and not in_comparison and not in_progress:
                in_card = True
                text = line[2:]
                m = _re.match(r'^([\U0001F300-\U0001F9FF]\s*)?\*\*(.+?)\*\*', text)
                if m:
                    card_icon = m.group(1).strip() if m.group(1) else ""
                    card_title = m.group(2)
                    remaining = text[m.end():].strip()
                    if remaining: card_lines.append(remaining)
                else:
                    card_lines.append(text)
                continue
            elif in_card and line.startswith('> '):
                card_lines.append(line[2:])
                continue
            
            # KPI: 📊 85% | Customer Satisfaction
            if '|' in line and '%' in line and not in_card and not in_timeline and not in_steps and not in_quote and not in_comparison and not in_progress:
                parts = [p.strip() for p in line.split('|')]
                if len(parts) >= 2 and any('%' in p for p in parts):
                    in_kpi = True
                    kpi_data.append(parts)
                    continue
            elif in_kpi and '|' in line and '%' in line:
                parts = [p.strip() for p in line.split('|')]
                kpi_data.append(parts)
                continue
            elif in_kpi:
                # Render KPI cards
                kpi_table = doc.add_table(rows=1, cols=len(kpi_data))
                kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                for j, kpi in enumerate(kpi_data):
                    if j < len(kpi_data):
                        val = kpi[0] if len(kpi) > 0 else ""
                        lbl = kpi[1] if len(kpi) > 1 else ""
                        icon_match = _re.match(r'^([\U0001F300-\U0001F9FF]\s*)', val)
                        icon = icon_match.group(1).strip() if icon_match else ""
                        if icon: val = val[len(icon):].strip()
                        cell = kpi_table.rows[0].cells[j]
                        cell.width = Inches(2.8)
                        tcC = cell._tc; tcPrC = tcC.get_or_add_tcPr()
                        shdC = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["light"]}" w:val="clear"/>')
                        tcPrC.append(shdC)
                        pC = cell.paragraphs[0]; pC.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        pC.paragraph_format.space_before = Pt(12)
                        runC = pC.add_run(f"{icon} {val}" if icon else val)
                        runC.font.size = Pt(24); runC.font.bold = True
                        runC.font.color.rgb = hex_to_rgb(colors["primary"])
                        runC.font.name = font_pair["heading"]
                        pC2 = cell.add_paragraph(); pC2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        pC2.paragraph_format.space_after = Pt(12)
                        runC2 = pC2.add_run(_format_text(lbl, mode=fmt_mode))
                        runC2.font.size = Pt(9); runC2.font.color.rgb = hex_to_rgb(colors["text"])
                        runC2.font.name = font_pair["body"]
                doc.add_paragraph()
                kpi_data = []; in_kpi = False
            
            # Timeline: 📅 2024 | Event description (pipe format)
            if '|' in line and _re.match(r'^[\U0001F300-\U0001F9FF\s]*\d{4}', line) and not in_card and not in_kpi and not in_steps and not in_quote and not in_comparison and not in_progress:
                in_timeline = True
                parts = [p.strip() for p in line.split('|', 1)]
                timeline_data.append((parts[0], parts[1] if len(parts) > 1 else ""))
                continue
            elif in_timeline and '|' in line and _re.match(r'^[\U0001F300-\U0001F9FF\s]*\d{4}', line):
                parts = [p.strip() for p in line.split('|', 1)]
                timeline_data.append((parts[0], parts[1] if len(parts) > 1 else ""))
                continue
            elif in_timeline:
                add_timeline(doc, timeline_data)
                timeline_data = []; in_timeline = False
            
            # Timeline (colon/dash format): 2024 - Event or Jan 15: Event
            if not in_timeline and not in_card and not in_kpi and not in_steps and not in_quote and not in_comparison and not in_progress:
                tl_m = _timeline_re.match(line)
                if tl_m:
                    in_timeline = True
                    timeline_data.append((tl_m.group(1), tl_m.group(2)))
                    continue
            
            # Steps: 1. Step one / 2. Step two
            if _re.match(r'^\d+\.\s', line) and not in_card and not in_kpi and not in_timeline and not in_quote and not in_comparison and not in_progress:
                in_steps = True
                step_lines.append(_re.sub(r'^\d+\.\s', '', line))
                continue
            elif in_steps and _re.match(r'^\d+\.\s', line):
                step_lines.append(_re.sub(r'^\d+\.\s', '', line))
                continue
            elif in_steps:
                if len(step_lines) >= 3:
                    add_step_guide(doc, step_lines)
                else:
                    for i, step_text in enumerate(step_lines, 1):
                        _add_rich_paragraph(doc, f"{i}. {_format_text(step_text, mode=fmt_mode)}", font_name=font_pair["body"], font_size=11, color=colors["text"])
                step_lines = []; in_steps = False

            # Pull quote: "Quote text" — Author
            if line.startswith('"') and line.endswith('"') and not in_card and not in_kpi and not in_timeline and not in_steps and not in_comparison and not in_progress:
                in_quote = True
                quote_text = line.strip('"')
                continue
            elif in_quote and (line.startswith('\u2014') or (line.strip() and not line.startswith('"') and not line.startswith('#') and not line.startswith('-') and not line.startswith('*') and not line.startswith('>') and not line.startswith('|') and not line.startswith('@') and not line.startswith('```'))):
                quote_author = line.lstrip('\u2014 ').strip()
                add_pull_quote(doc, quote_text, quote_author)
                in_quote = False; quote_text = ""; quote_author = ""
                continue
            
            # Comparison table: | Feature | A | B |
            if line.startswith('|') and line.endswith('|') and not in_card and not in_kpi and not in_timeline and not in_steps and not in_quote and not in_progress:
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if all(c.startswith('---') for c in cells):
                    continue
                if not in_comparison:
                    in_comparison = True
                    comp_headers = cells
                else:
                    comp_rows.append(cells)
                continue
            elif in_comparison and line.startswith('|') and line.endswith('|'):
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if not all(c.startswith('---') for c in cells):
                    comp_rows.append(cells)
                continue
            elif in_comparison:
                add_comparison_table(doc, comp_headers, comp_rows)
                comp_headers = []; comp_rows = []; in_comparison = False
            
            # Progress: Label: 75%
            if ':' in line and _re.search(r'(\d+)%', line) and not in_card and not in_kpi and not in_timeline and not in_steps and not in_quote and not in_comparison:
                in_progress = True
                parts = line.split(':', 1)
                lbl = parts[0].strip()
                pct_match = _re.search(r'(\d+)%', parts[1])
                pct = int(pct_match.group(1)) if pct_match else 0
                progress_data.append((lbl, pct, colors["accent"]))
                continue
            elif in_progress and ':' in line and _re.search(r'(\d+)%', line):
                parts = line.split(':', 1)
                lbl = parts[0].strip()
                pct_match = _re.search(r'(\d+)%', parts[1])
                pct = int(pct_match.group(1)) if pct_match else 0
                progress_data.append((lbl, pct, colors["accent"]))
                continue
            elif in_progress:
                for lbl, pct, clr in progress_data:
                    add_progress_bar(doc, lbl, pct, clr)
                progress_data = []; in_progress = False
            
            # Status badge: @success Task complete
            if line.startswith('@') and not in_card and not in_kpi and not in_timeline and not in_steps and not in_quote and not in_comparison and not in_progress:
                parts = line[1:].split(None, 1)
                status = parts[0].lower() if parts else "info"
                text = parts[1] if len(parts) > 1 else ""
                add_status_badge(doc, text, status)
                continue
            
            # Visual separator: --- or ***
            if line in ('---', '***', '...'):
                style_map = {'---': 'line', '***': 'dots', '...': 'dash'}
                add_visual_separator(doc, style_map.get(line, 'dots'))
                continue
            
            # Headings with emoji
            if line.startswith('# ') or line.startswith('## ') or line.startswith('### '):
                level = line.count('#')
                text = line.lstrip('#').strip()
                heading_level = min(level, 3)
                _add_rich_paragraph(doc, text, style=f'Heading {heading_level}', font_name=font_pair["heading"], font_size=sizes.get(heading_level, 14), color=colors["primary"])
                continue
            
            # Icon bullets
            if line.startswith('- ') or line.startswith('* '):
                bullet_text = line[2:].strip()
                text = _format_text(bullet_text, mode=fmt_mode)
                icon = "\u2022"
                if text.lower().startswith(('done','complete','yes','ok','success','conclu')):
                    icon = "\u2705"
                elif text.lower().startswith(('no','not','fail','error','wrong','nao')):
                    icon = "\u274c"
                elif text.lower().startswith(('warn','caution','careful','cuidado')):
                    icon = "\u26a0\ufe0f"
                elif text.lower().startswith(('note','info','note','nota')):
                    icon = "\U0001f4cc"
                elif text.lower().startswith(('idea','tip','suggestion','sugest')):
                    icon = "\U0001f4a1"
                _add_rich_paragraph(doc, f"{icon} {text}", style='List Bullet', font_name=font_pair["body"], font_size=11, color=colors["text"])
                continue
            
            # Regular paragraph
            text = _format_text(line, mode=fmt_mode)
            _add_rich_paragraph(doc, text, font_name=font_pair["body"], font_size=11, color=colors["text"])
        
        # Render remaining
        if in_card:
            add_card_box(doc, card_lines, colors["accent"], colors["light"], card_icon, card_title)
        if in_steps:
            if len(step_lines) >= 3:
                add_step_guide(doc, step_lines)
            else:
                for i, step_text in enumerate(step_lines, 1):
                    _add_rich_paragraph(doc, f"{i}. {_format_text(step_text, mode=fmt_mode)}", font_name=font_pair["body"], font_size=11, color=colors["text"])
        if in_timeline:
            add_timeline(doc, timeline_data)
        if in_comparison:
            add_comparison_table(doc, comp_headers, comp_rows)
        if in_progress:
            for lbl, pct, clr in progress_data:
                add_progress_bar(doc, lbl, pct, clr)
        if in_code_block:
            for code_line in code_lines:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                run = p.add_run(code_line)
                run.font.name = "Consolas"
                run.font.size = Pt(9)
                run.font.color.rgb = hex_to_rgb(colors["text"])
        
        # --- Branded Footer ---
        footer = section.footer
        footer.is_linked_to_previous = False
        ft = footer.add_table(rows=1, cols=3, width=Inches(6.5))
        ft.alignment = WD_TABLE_ALIGNMENT.CENTER
        c1 = ft.rows[0].cells[0]; c1.width = Inches(2)
        r1 = c1.paragraphs[0]; r1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run1 = r1.add_run("Edit Office Files")
        run1.font.size = Pt(8); run1.font.color.rgb = hex_to_rgb(colors["accent"])
        run1.font.name = font_pair["body"]
        c2 = ft.rows[0].cells[1]; c2.width = Inches(2.5)
        r2 = c2.paragraphs[0]; r2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = r2.add_run("Page ")
        run2.font.size = Pt(8); run2.font.color.rgb = hex_to_rgb(colors["accent"])
        run2.font.name = font_pair["body"]
        fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        run2._r.append(fldChar1)
        instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
        run2._r.append(instrText)
        fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        run2._r.append(fldChar2)
        c3 = ft.rows[0].cells[2]; c3.width = Inches(2)
        r3 = c3.paragraphs[0]; r3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run3 = r3.add_run(datetime.datetime.now().strftime("%Y-%m-%d"))
        run3.font.size = Pt(8); run3.font.color.rgb = hex_to_rgb(colors["accent"])
        run3.font.name = font_pair["body"]
        
        # --- Save ---
        file_bytes = io.BytesIO()
        doc.save(file_bytes)
        file_bytes.seek(0)
        url, fname = await self._save_and_link(file_bytes.getvalue(), f"{title}.docx", __request__)
        if url:
            return f"Document created: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})


    async def generate_slides(self, content: str, title: str = "Presentation", theme: str = "modern", raw_text: bool = False, __user__=None, __request__=None) -> str:
        """Generate professional PowerPoint slides with modern design.

        Args:
            content: Markdown content (headings become slides, bullets become content)
            title: Presentation title for the first slide
            theme: Visual theme - modern, light, dark, corporate, creative, minimal
            raw_text: If True, skip text formatting
        Returns:
            Markdown link to the generated file
        """
        fmt_mode = "preserve" if raw_text else "format"
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
                p.text = text
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
            add_text_box(slide, 1.5, 1.5, 10, 1.5, title, font_size=44, bold=True, color=colors["text"])
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
                        _render_content_slide(prs, current_slide_lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_count, fmt_mode=fmt_mode)
                        slide_count += 1
                    current_slide_lines = [line]
                else:
                    current_slide_lines.append(line)

            if current_slide_lines:
                _render_content_slide(prs, current_slide_lines, colors, hex_to_rgb, add_text_box, add_accent_bar, set_slide_bg, slide_count, fmt_mode=fmt_mode)

            out = io.BytesIO()
            prs.save(out)
            out.seek(0)
            url, fname = await self._save_and_link(out.getvalue(), "%s.pptx" % title, __request__)
            if url:
                return "Presentation created: [%s](%s)" % (fname, url)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def generate_spreadsheet(self, content: str, title: str = "Spreadsheet", theme: str = "professional", raw_text: bool = False, __user__=None, __request__=None) -> str:
        """Generate a professional Excel spreadsheet with modern styling.

        Args:
            content: CSV or tab-delimited data (first row = headers, rest = data)
            title: Spreadsheet title / filename
            theme: Visual theme - professional, modern, corporate, minimal, colorful, pastel
            raw_text: If True, skip text formatting
        Returns:
            Markdown link to the generated file
        """
        fmt_mode = "preserve" if raw_text else "format"
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
            for i, row in enumerate(data):
                if i == 0:
                    ws.append(row)  # Headers: preserve original case
                else:
                    ws.append([_format_text(c, mode=fmt_mode) if isinstance(c, str) else c for c in row])

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
        Preserves original formatting and capitalization of the existing document.
    
        change_type: replace (use old_text|||new_text), insert (append text with redline), delete (mark paragraph for deletion)
        author: Name shown in Word's Track Changes (e.g., "Sergio Pedro")
        """
        try:
            import sqlite3 as s3
            conn2 = s3.connect(_DB_PATH)
            try:
                row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
                if not row:
                    row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
                if not row:
                    return json.dumps({"error": "File not found"})
                filename = row[0]
                meta = json.loads(row[1]) if row[1] else {}
                data = _read_file_bytes(meta.get("path", file_id))
                if not data:
                    return json.dumps({"error": "File not found on disk"})
            finally:
                conn2.close()

            # Check file type — track changes only supported for DOCX
            ftype = _detect_type(filename)
            if ftype != "docx":
                return json.dumps({"error": f"Track changes are only supported for DOCX files. {ftype.upper()} format does not support revision marks."})

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
            try:
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
                    data = _read_file_bytes(meta.get("path", fid))
                    if not data:
                        continue
                    wb_src = openpyxl.load_workbook(io.BytesIO(data))
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
            finally:
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
                else:
                    results.append(f"  {fid}: unsupported operation '{operation}'")
            if results:
                return "Batch processed " + str(len(ids)) + " files:\n" + "\n".join(results)
            return json.dumps({"error": f"No files processed. Unsupported operation: {operation}"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    async def auto_backup(self, __user__=None, __request__=None) -> str:
        try:
            import shutil, datetime
            db_path = _DB_PATH
            backup_dir = os.path.join(_get_owui_data_dir(), "backups")
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
            try:
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
                    data = _read_file_bytes(meta.get("path", fid))
                    if not data:
                        continue
                    src = fitz.open(stream=data, filetype="pdf")
                    merger.insert_pdf(src)
                    src.close()
                    count += 1
            finally:
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
            try:
                row = conn2.execute("SELECT meta FROM file WHERE id=?", (file_id,)).fetchone()
                if not row:
                    row = conn2.execute("SELECT meta FROM file WHERE filename LIKE ?", ("%"+file_id+"%",)).fetchone()
                if not row:
                    return json.dumps({"error": "File not found"})
                meta = json.loads(row[0]) if row[0] else {}
                data = _read_file_bytes(meta.get("path", file_id))
                if not data:
                    return json.dumps({"error": "File not found on disk"})
            finally:
                conn2.close()
            src = fitz.open(stream=data, filetype="pdf")
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
            try:
                try:
                    tool_count = conn2.execute("SELECT COUNT(*) FROM tool WHERE is_active=1").fetchone()[0]
                except Exception:
                    tool_count = 0
                try:
                    func_count = conn2.execute("SELECT COUNT(*) FROM function WHERE is_active=1").fetchone()[0]
                except Exception:
                    func_count = 0
                try:
                    model_count = conn2.execute("SELECT COUNT(*) FROM model WHERE is_active=1").fetchone()[0]
                except Exception:
                    model_count = 0
                exports_dir = os.path.join(_get_owui_data_dir(), "exports")
                export_count = len([f for f in os.listdir(exports_dir) if os.path.isfile(os.path.join(exports_dir, f))]) if os.path.exists(exports_dir) else 0
                db_size_kb = os.path.getsize(_DB_PATH) / 1024
            finally:
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
        try:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT id, filename, created_at FROM file WHERE meta LIKE '%office-plugin%' AND created_at < ?",
                (cutoff,)
            ).fetchall()
            
            if not rows:
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
        finally:
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
            try:
                row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
                if not row:
                    row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
                if not row:
                    return json.dumps({"error": "File not found"})
                filename = row[0]
                meta = json.loads(row[1]) if row[1] else {}
                data = _read_file_bytes(meta.get("path", file_id))
                if not data:
                    return json.dumps({"error": "File not found on disk"})
            finally:
                conn2.close()

            # Check file type — revision management only supported for DOCX
            ftype = _detect_type(filename)
            if ftype != "docx":
                return json.dumps({"error": f"Revision management is only supported for DOCX files. {ftype.upper()} format does not support revision marks."})

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



    # --- v3.2.0: ODF Write ---
    async def create_odf(self, content: str, filename: str = "document", format: str = "odt", raw_text: bool = False, __user__=None, __request__=None) -> str:
        """Create a new ODF file (.odt, .ods, .odp).

        Args:
            content: Content for the file (CSV for .ods, text for .odt, markdown for .odp)
            filename: Output filename
            format: 'odt', 'ods', or 'odp'
            raw_text: If True, skip text formatting
        """
        from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet, OpenDocumentPresentation
        from odf.text import P, H
        from odf.table import Table, TableRow, TableCell
        import io
        fmt_mode = "preserve" if raw_text else "format"
        
        try:
            if format == "ods":
                doc = OpenDocumentSpreadsheet()
                lines = content.split('\n')
                table = Table(name="Sheet1")
                for line in lines:
                    line = line.strip()
                    if not line: continue
                    cells = [c.strip() for c in line.split(',')]
                    row = TableRow()
                    for c in cells:
                        cell = TableCell()
                        cell.addElement(P(text=_format_text(c, mode=fmt_mode)))
                        row.addElement(cell)
                    table.addElement(row)
                doc.spreadsheet.addElement(table)
            elif format == "odp":
                doc = OpenDocumentPresentation()
                for line in content.split('\n'):
                    line = line.strip()
                    if not line: continue
                    if line.startswith('# '):
                        doc.presentation.addElement(H(outlinelevel=1, text=line[2:]))
                    else:
                        doc.presentation.addElement(P(text=_format_text(line, mode=fmt_mode)))
            else:
                doc = OpenDocumentText()
                for line in content.split('\n'):
                    line = line.strip()
                    if not line: continue
                    if line.startswith('# '):
                        doc.text.addElement(H(outlinelevel=1, text=line[2:]))
                    elif line.startswith('## '):
                        doc.text.addElement(H(outlinelevel=2, text=line[3:]))
                    else:
                        doc.text.addElement(P(text=_format_text(line, mode=fmt_mode)))
            
            buf = io.BytesIO()
            doc.write(buf)
            buf.seek(0)
            url, fname = await self._save_and_link(buf.getvalue(), f"{filename}.{format}", __request__, __user__=__user__)
            if url:
                return f"ODF file created: [{fname}]({url})"
            return json.dumps({"error": "Could not save file"})
        except ImportError:
            return json.dumps({"error": "odfpy not installed. Install with: pip install odfpy"})
        except Exception as e:
            return json.dumps({"error": str(e)})

    # --- v3.2.0: Format Conversion ---
    async def convert_format(self, file_id: str, target_format: str, __user__=None, __request__=None) -> str:
        """Convert between Office formats (docx<->odt, xlsx<->ods, pptx<->odp)."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        base_name = os.path.splitext(filename)[0]
        
        if ftype in ("xlsx", "xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx":
            content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx":
            content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt", "ods", "odp"):
            content = _read_odf(file_bytes, filename)
        else:
            return json.dumps({"error": f"Unsupported source format: {ftype}"})
        
        if target_format in ("odt", "ods", "odp"):
            return await self.create_odf(content, base_name, target_format, __user__=__user__, __request__=__request__)
        elif target_format == "docx":
            return await self.generate_document(content, base_name, __user__=__user__, __request__=__request__)
        elif target_format == "pptx":
            return await self.generate_slides(content, base_name, __user__=__user__, __request__=__request__)
        elif target_format == "xlsx":
            return await self.generate_spreadsheet(content, base_name, __user__=__user__, __request__=__request__)
        return json.dumps({"error": f"Unsupported target format: {target_format}"})

    # --- v3.2.0: Template System ---
    async def save_template(self, name: str, content: str) -> str:
        """Save a document template for reuse."""
        templates = json.loads(self.valves.templates or "{}")
        templates[name] = content
        self.valves.templates = json.dumps(templates)
        return f"Template '{name}' saved."

    async def use_template(self, name: str, __user__=None, __request__=None, **kwargs) -> str:
        """Generate a document from a saved template, replacing {placeholders}."""
        templates = json.loads(self.valves.templates or "{}")
        if name not in templates:
            return f"Template '{name}' not found. Available: {', '.join(templates.keys())}"
        content = templates[name]
        for key, value in kwargs.items():
            content = content.replace(f"{{{key}}}", str(value))
        return await self.generate_document(content, name, __user__=__user__, __request__=__request__)

    async def list_templates(self) -> str:
        """List all saved templates."""
        templates = json.loads(self.valves.templates or "{}")
        if not templates:
            return "No templates saved."
        result = "Available templates:\n"
        for name in templates:
            preview = templates[name][:50].replace('\n', ' ')
            result += f"- {name}: {preview}...\n"
        return result

    # --- v3.2.0: Scheduled Cleanup ---
    async def schedule_cleanup(self, days_old: int = 30, interval_hours: int = 24) -> str:
        """Schedule automatic cleanup every N hours. Set interval_hours=0 to disable."""
        schedule = {"days_old": days_old, "interval_hours": interval_hours, "enabled": interval_hours > 0}
        self.valves.cleanup_schedule = json.dumps(schedule)
        if interval_hours > 0:
            return f"Cleanup scheduled: remove files older than {days_old} days, every {interval_hours} hours."
        return "Scheduled cleanup disabled."

    # --- v3.2.0: Mail Merge ---
    async def mail_merge(self, template_file_id: str, data_file_id: str, output_prefix: str = "merged", __user__=None, __request__=None) -> str:
        """Generate personalized documents by merging CSV/Excel data into a DOCX template."""
        from docx import Document
        import io, csv
        
        t_bytes, t_name, t_type = self._resolve_file(template_file_id)
        if not t_bytes:
            return json.dumps({"error": f"Template not found: {template_file_id}"})
        
        d_bytes, d_name, d_type = self._resolve_file(data_file_id)
        if not d_bytes:
            return json.dumps({"error": f"Data file not found: {data_file_id}"})
        
        rows = []
        if d_type in ("xlsx", "xls"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(d_bytes))
            ws = wb.active
            headers = [str(c.value or '') for c in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, [str(v or '') for v in row])))
        else:
            reader = csv.DictReader(io.StringIO(d_bytes.decode('utf-8')))
            rows = list(reader)
        
        if not rows:
            return json.dumps({"error": "No data rows found"})
        
        results = []
        for i, row in enumerate(rows):
            doc = Document(io.BytesIO(t_bytes))
            for para in doc.paragraphs:
                for key, value in row.items():
                    if f"{{{{{key}}}}}" in para.text:
                        for run in para.runs:
                            run.text = run.text.replace(f"{{{{{key}}}}}", value)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            fname = f"{output_prefix}_{i+1}.docx"
            url, saved = await self._save_and_link(buf.getvalue(), fname, __request__, __user__=__user__)
            if url:
                results.append(f"[{saved}]({url})")
        
        return f"Merged {len(results)} documents:\n" + "\n".join(results)

    # --- v3.2.0: Chart Generation ---
    async def add_chart(self, file_id: str, chart_type: str = "bar", title: str = "Chart", __user__=None, __request__=None) -> str:
        """Add a chart to an Excel spreadsheet. chart_type: bar, line, pie, scatter."""
        from openpyxl import load_workbook
        from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
        from openpyxl.utils import get_column_letter
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        
        chart_types = {"bar": BarChart, "line": LineChart, "pie": PieChart, "scatter": ScatterChart}
        chart_class = chart_types.get(chart_type, BarChart)
        chart = chart_class()
        chart.title = title
        chart.style = 10
        
        if ws.max_row > 1:
            data = Reference(ws, min_col=1, min_row=1, max_row=ws.max_row, max_col=ws.max_column)
            chart.add_data(data, titles_from_data=True)
        
        chart_col = get_column_letter(ws.max_column + 2)
        ws.add_chart(chart, f"{chart_col}1")
        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url:
            return f"Chart added: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # --- v3.2.0: Watermark ---
    async def add_watermark(self, file_id: str, text: str = "DRAFT", __user__=None, __request__=None) -> str:
        """Add a diagonal watermark to a DOCX or PDF file."""
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        if ftype == "docx":
            from docx import Document
            from docx.shared import Pt, RGBColor
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            doc = Document(io.BytesIO(file_bytes))
            for section in doc.sections:
                header = section.header
                header.is_linked_to_previous = False
                p = header.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.text = text
                run.font.size = Pt(72)
                run.font.color.rgb = RGBColor(128, 128, 128)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
        elif ftype == "pdf":
            try:
                import fitz
                pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
                for page in pdf_doc:
                    rect = page.rect
                    page.insert_text((rect.width/2-100, rect.height/2), text, fontsize=72, color=(0.5,0.5,0.5), alpha=0.1, rotate=45)
                buf = io.BytesIO()
                pdf_doc.save(buf)
                pdf_doc.close()
                buf.seek(0)
            except ImportError:
                return json.dumps({"error": "PyMuPDF not installed. Install with: pip install PyMuPDF"})
        else:
            return json.dumps({"error": f"Watermark not supported for {ftype}. Use DOCX or PDF."})
        
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url:
            return f"Watermark added: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # --- v3.2.0: File Preview ---
    async def preview_file(self, file_id: str, max_lines: int = 20) -> str:
        """Show a text preview of any Office file before downloading."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        if ftype in ("xlsx", "xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx":
            content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx":
            content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt", "ods", "odp"):
            content = _read_odf(file_bytes, filename)
        else:
            return json.dumps({"error": f"Preview not supported for {ftype}"})
        
        lines = content.split('\n')
        preview = '\n'.join(lines[:max_lines])
        total = len(lines)
        result = f"**{filename}** ({ftype.upper()}, {total} lines)\n\n```\n{preview}\n```"
        if total > max_lines:
            result += f"\n... ({total - max_lines} more lines)"
        return result

    # --- v3.2.0: Metadata Editing ---
    async def edit_metadata(self, file_id: str, author: str = "", title: str = "", subject: str = "", keywords: str = "", __user__=None, __request__=None) -> str:
        """Edit document metadata (author, title, subject, keywords)."""
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        changes = []
        if ftype == "docx":
            from docx import Document
            doc = Document(io.BytesIO(file_bytes))
            cp = doc.core_properties
            if author: cp.author = author; changes.append("author")
            if title: cp.title = title; changes.append("title")
            if subject: cp.subject = subject; changes.append("subject")
            if keywords: cp.keywords = keywords; changes.append("keywords")
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        elif ftype == "xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(file_bytes))
            if author: wb.properties.creator = author; changes.append("author")
            if title: wb.properties.title = title; changes.append("title")
            if subject: wb.properties.subject = subject; changes.append("subject")
            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        elif ftype == "pptx":
            from pptx import Presentation
            prs = Presentation(io.BytesIO(file_bytes))
            cp = prs.core_properties
            if author: cp.author = author; changes.append("author")
            if title: cp.title = title; changes.append("title")
            if subject: cp.subject = subject; changes.append("subject")
            if keywords: cp.keywords = keywords; changes.append("keywords")
            buf = io.BytesIO(); prs.save(buf); buf.seek(0)
        else:
            return json.dumps({"error": f"Metadata editing not supported for {ftype}"})
        
        if not changes:
            return json.dumps({"error": "No metadata fields specified"})
        
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url:
            return f"Metadata updated ({', '.join(changes)}): [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # --- v3.2.0: Accessibility Check ---
    async def check_accessibility(self, file_id: str) -> str:
        """Check a document for accessibility issues."""
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        issues = []
        if ftype == "docx":
            from docx import Document
            doc = Document(io.BytesIO(file_bytes))
            headings = []
            for p in doc.paragraphs:
                if p.style.name.startswith('Heading'):
                    level = int(p.style.name.split()[-1]) if p.style.name.split()[-1].isdigit() else 1
                    headings.append((level, p.text[:50]))
            prev = 0
            for level, text in headings:
                if level > prev + 1:
                    issues.append(f"Heading skip: H{prev} to H{level} ('{text}')")
                prev = level
            if not headings:
                issues.append("No headings found - document may lack structure")
        elif ftype == "pptx":
            from pptx import Presentation
            prs = Presentation(io.BytesIO(file_bytes))
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if shape.shape_type == 13 and not shape.alt_text:
                        issues.append(f"Slide {i+1}: Image missing alt text")
        
        if not issues:
            return f"Accessibility check passed for {filename}. No issues found."
        result = f"**Accessibility Report: {filename}**\n\n"
        for issue in issues:
            result += f"- {issue}\n"
        result += f"\n{len(issues)} issue(s) found."
        return result

    # --- v3.2.0: Add Alt Text ---
    async def add_alt_text(self, file_id: str, slide_num: int = 1, shape_index: int = 0, alt_text: str = "", __user__=None, __request__=None) -> str:
        """Add alt text to an image in a PowerPoint slide."""
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        if ftype != "pptx":
            return json.dumps({"error": "Alt text only supported for PPTX files"})
        
        from pptx import Presentation
        prs = Presentation(io.BytesIO(file_bytes))
        if slide_num > len(prs.slides):
            return json.dumps({"error": f"Slide {slide_num} not found. Has {len(prs.slides)} slides."})
        
        slide = prs.slides[slide_num - 1]
        pic_count = 0
        for shape in slide.shapes:
            if shape.shape_type == 13:
                if pic_count == shape_index:
                    shape.alt_text = alt_text
                    buf = io.BytesIO()
                    prs.save(buf)
                    buf.seek(0)
                    url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
                    if url:
                        return f"Alt text added: [{fname}]({url})"
                    return json.dumps({"error": "Could not save file"})
                pic_count += 1
        return json.dumps({"error": f"Image {shape_index} not found. Found {pic_count} images."})

    # --- v3.2.0: Progress Indicators ---
    async def _progress(self, current: int, total: int, operation: str = "Processing", __event_emitter__=None) -> None:
        """Emit progress via __event_emitter__ if available."""
        if __event_emitter__ is None:
            return
        try:
            await __event_emitter__({"type": "status", "data": {"description": f"{operation}: {current}/{total}", "done": current >= total}})
        except Exception:
            pass


    # --- v3.3.0: Document Comparison ---
    async def compare_documents(self, file_id_a: str, file_id_b: str) -> str:
        """Compare two documents and show differences."""
        a_bytes, a_name, a_type = self._resolve_file(file_id_a)
        b_bytes, b_name, b_type = self._resolve_file(file_id_b)
        if not a_bytes or not b_bytes:
            return json.dumps({"error": "One or both files not found"})
        
        def get_text(ftype, fb, fn):
            if ftype in ("xlsx","xls"):
                return self._read_xlsx(fb, fn) if ftype == "xlsx" else self._read_xls(fb, fn)
            elif ftype == "docx": return self._read_docx(fb, fn)
            elif ftype == "pptx": return self._read_pptx(fb, fn)
            elif ftype in ("odt","ods","odp"): return _read_odf(fb, fn)
            return ""
        
        text_a = get_text(a_type, a_bytes, a_name)
        text_b = get_text(b_type, b_bytes, b_name)
        
        lines_a = text_a.split('\n')
        lines_b = text_b.split('\n')
        
        result = f"**Comparison: {a_name} vs {b_name}**\n\n"
        added = 0; removed = 0; changed = 0
        max_len = max(len(lines_a), len(lines_b))
        
        for i in range(max_len):
            la = lines_a[i] if i < len(lines_a) else None
            lb = lines_b[i] if i < len(lines_b) else None
            if la is None:
                result += f"+ L{i+1}: {lb}\n"; added += 1
            elif lb is None:
                result += f"- L{i+1}: {la}\n"; removed += 1
            elif la != lb:
                result += f"~ L{i+1}:\n  - {la}\n  + {lb}\n"; changed += 1
        
        result += f"\n**Summary:** {added} added, {removed} removed, {changed} changed"
        return result

    # --- v3.3.0: Export to Markdown ---
    async def export_to_markdown(self, file_id: str, __user__=None, __request__=None) -> str:
        """Export any Office file to Markdown format."""
        import io
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        if ftype in ("xlsx","xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx":
            content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx":
            content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"):
            content = _read_odf(file_bytes, filename)
        else:
            return json.dumps({"error": f"Export not supported for {ftype}"})
        
        base = os.path.splitext(filename)[0]
        md_content = f"# {base}\n\n{content}"
        md_bytes = md_content.encode('utf-8')
        
        url, fname = await self._save_and_link(md_bytes, f"{base}.md", __request__, __user__=__user__)
        if url:
            return f"Exported to Markdown: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # --- v3.3.0: Import from URL ---
    async def import_from_url(self, url: str, title: str = "Web Document", __user__=None, __request__=None) -> str:
        """Fetch a web page and convert it to a Word document."""
        try:
            import urllib.request as _urllib
            req = _urllib.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            resp = _urllib.urlopen(req, timeout=15)
            html = resp.read().decode('utf-8', errors='replace')
        except Exception as e:
            return json.dumps({"error": f"Failed to fetch URL: {str(e)}"})
        
        # Simple HTML to text
        import re as _re
        text = _re.sub(r'<script[^>]*>.*?</script>', '', html, flags=_re.DOTALL|_re.IGNORECASE)
        text = _re.sub(r'<style[^>]*>.*?</style>', '', text, flags=_re.DOTALL|_re.IGNORECASE)
        text = _re.sub(r'<[^>]+>', '\n', text)
        text = _re.sub(r'\n\s*\n', '\n\n', text)
        text = _re.sub(r'[ \t]+', ' ', text)
        text = '\n'.join(line.strip() for line in text.split('\n') if line.strip())
        
        if not text:
            return json.dumps({"error": "No text content extracted from URL"})
        
        return await self.generate_document(text[:50000], title, __user__=__user__, __request__=__request__)

    # --- v3.3.0: File Versioning ---
    async def version_file(self, file_id: str, label: str = "", __user__=None, __request__=None) -> str:
        """Save a versioned copy of a file before editing."""
        import time as _time
        
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": f"File not found: {file_id}"})
        
        base, ext = os.path.splitext(filename)
        ts = _time.strftime("%Y%m%d_%H%M%S")
        label_str = f"_{label}" if label else ""
        version_name = f"{base}_v{ts}{label_str}{ext}"
        
        url, fname = await self._save_and_link(file_bytes, version_name, __request__, __user__=__user__)
        if url:
            return f"Version saved: [{fname}]({url})"
        return json.dumps({"error": "Could not save version"})

    # --- v3.3.0: Cloud Storage (Google Drive) ---
    async def upload_to_drive(self, file_id: str, folder_id: str = "root") -> str:
        """Upload a file to Google Drive. Requires google-api-python-client and credentials."""
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaIoBaseUpload
            import io as _io
            
            file_bytes, filename, ftype = self._resolve_file(file_id)
            if not file_bytes:
                return json.dumps({"error": f"File not found: {file_id}"})
            
            creds_path = os.environ.get("GOOGLE_CREDENTIALS", "")
            if not creds_path or not os.path.exists(creds_path):
                return json.dumps({"error": "GOOGLE_CREDENTIALS env var not set or file not found"})
            
            mime_map = {"xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       "pdf": "application/pdf"}
            mime = mime_map.get(ftype, "application/octet-stream")
            
            creds = service_account.Credentials.from_service_account_file(creds_path, scopes=["https://www.googleapis.com/auth/drive.file"])
            service = build("drive", "v3", credentials=creds)
            
            media = MediaIoBaseUpload(_io.BytesIO(file_bytes), mimetype=mime, resumable=True)
            file_metadata = {"name": filename, "parents": [folder_id]}
            drive_file = service.files().create(body=file_metadata, media_body=media, fields="id,webViewLink").execute()
            
            return f"Uploaded to Google Drive: {drive_file.get('webViewLink', drive_file.get('id'))}"
        except ImportError:
            return json.dumps({"error": "google-api-python-client not installed. pip install google-api-python-client google-auth"})
        except Exception as e:
            return json.dumps({"error": f"Drive upload failed: {str(e)}"})

    # --- v3.3.0: OCR ---
    async def ocr_extract(self, file_id: str, language: str = "eng") -> str:
        """Extract text from images in a document using OCR. Requires pytesseract and Pillow."""
        try:
            import pytesseract
            from PIL import Image
            import io as _io
            
            file_bytes, filename, ftype = self._resolve_file(file_id)
            if not file_bytes:
                return json.dumps({"error": f"File not found: {file_id}"})
            
            if ftype == "pdf":
                try:
                    import fitz
                    pdf = fitz.open(stream=file_bytes, filetype="pdf")
                    results = []
                    for i, page in enumerate(pdf):
                        pix = page.get_pixmap(dpi=200)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        text = pytesseract.image_to_string(img, lang=language)
                        if text.strip():
                            results.append(f"--- Page {i+1} ---\n{text.strip()}")
                    pdf.close()
                    return "\n\n".join(results) if results else "No text found in PDF images"
                except ImportError:
                    return json.dumps({"error": "PyMuPDF not installed"})
            else:
                return json.dumps({"error": f"OCR only supported for PDF files. Got: {ftype}"})
        except ImportError:
            return json.dumps({"error": "pytesseract or Pillow not installed. pip install pytesseract Pillow"})
        except Exception as e:
            return json.dumps({"error": f"OCR failed: {str(e)}"})

    # --- v3.3.0: i18n Error Messages ---
    async def translate_errors(self, language: str = "en") -> str:
        """Set the language for error messages. Supported: en, pt, es, fr, de."""
        translations = {
            "en": {"file_not_found": "File not found", "could_not_save": "Could not save file", "unsupported": "Unsupported format"},
            "pt": {"file_not_found": "Ficheiro nao encontrado", "could_not_save": "Nao foi possivel guardar", "unsupported": "Formato nao suportado"},
            "es": {"file_not_found": "Archivo no encontrado", "could_not_save": "No se pudo guardar", "unsupported": "Formato no soportado"},
            "fr": {"file_not_found": "Fichier introuvable", "could_not_save": "Impossible d'enregistrer", "unsupported": "Format non pris en charge"},
            "de": {"file_not_found": "Datei nicht gefunden", "could_not_save": "Konnte nicht gespeichert werden", "unsupported": "Nicht unterstutztes Format"},
        }
        if language not in translations:
            return json.dumps({"error": f"Language '{language}' not supported. Available: {', '.join(translations.keys())}"})
        self.valves.language = language
        return f"Language set to {language}. Error messages will now appear in {language}."


    # --- v3.4.0: AI Summarize ---
    async def ai_summarize(self, file_id: str) -> str:
        """Extract document text for LLM summarization. The LLM will then summarize it."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes:
            return json.dumps({"error": "File not found: " + str(file_id)})
        if ftype in ("xlsx","xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format: " + str(ftype)})
        words = len(content.split())
        preview = content[:3000]
        return "**" + str(filename) + "** (" + str(words) + " words)\n\n" + preview + ("\n\n... (truncated, " + str(words) + " total words)" if len(content) > 3000 else "")

    # --- v3.4.0: Speaker Notes ---
    async def add_speaker_notes(self, file_id: str, slide_num: int, notes: str, __user__=None, __request__=None) -> str:
        """Add speaker notes to a PowerPoint slide."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype != "pptx": return json.dumps({"error": "Speaker notes only supported for PPTX"})
        from pptx import Presentation
        prs = Presentation(io.BytesIO(file_bytes))
        if slide_num < 1 or slide_num > len(prs.slides):
            return json.dumps({"error": "Slide " + str(slide_num) + " not found. Has " + str(len(prs.slides)) + " slides."})
        slide = prs.slides[slide_num - 1]
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes
        buf = io.BytesIO(); prs.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url: return "Speaker notes added to slide " + str(slide_num) + ": [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})

    # --- v3.4.0: Document Stats ---
    async def document_stats(self, file_id: str) -> str:
        """Show document statistics: word count, reading time, complexity."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        words = len(content.split())
        chars = len(content)
        lines = content.count('\n') + 1
        sentences = content.count('.') + content.count('!') + content.count('?')
        reading_time = max(1, words // 200)
        avg_word_len = sum(len(w) for w in content.split()) / max(1, words)
        complexity = "Easy" if avg_word_len < 4 else ("Medium" if avg_word_len < 5.5 else "Complex")
        return "**Stats: " + str(filename) + "**\n- Words: " + str(words) + "\n- Characters: " + str(chars) + "\n- Lines: " + str(lines) + "\n- Sentences: " + str(sentences) + "\n- Reading time: ~" + str(reading_time) + " min\n- Avg word length: " + str(round(avg_word_len, 1)) + "\n- Complexity: " + str(complexity)

    # --- v3.4.0: QR Codes ---
    async def add_qr_code(self, file_id: str, data: str, __user__=None, __request__=None) -> str:
        """Add a QR code to a DOCX or PPTX file."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        try:
            import qrcode
            from PIL import Image
        except ImportError:
            return json.dumps({"error": "qrcode or Pillow not installed. pip install qrcode Pillow"})
        qr = qrcode.QRCode(version=1, box_size=10, border=2)
        qr.add_data(data); qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        qr_buf = io.BytesIO(); img.save(qr_buf, format='PNG'); qr_buf.seek(0)
        if ftype == "docx":
            from docx import Document
            from docx.shared import Inches
            doc = Document(io.BytesIO(file_bytes))
            doc.add_picture(qr_buf, width=Inches(2))
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        elif ftype == "pptx":
            from pptx import Presentation
            from pptx.util import Inches
            prs = Presentation(io.BytesIO(file_bytes))
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(qr_buf, Inches(4), Inches(2), Inches(2))
            buf = io.BytesIO(); prs.save(buf); buf.seek(0)
        else: return json.dumps({"error": "QR codes only supported for DOCX and PPTX"})
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url: return "QR code added: [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})

    # --- v3.4.0: Bulk Folder Ops ---
    async def bulk_folder_ops(self, operation: str = "list", pattern: str = "*", __user__=None, __request__=None) -> str:
        """Apply an operation to all files in the uploads folder. Operations: list, delete_old, stats."""
        import glob as _glob, time as _time
        uploads = _UPLOAD_DIR
        if operation == "list":
            files = sorted(_glob.glob(os.path.join(uploads, pattern)))
            if not files: return "No files found matching: " + str(pattern)
            result = "**Files in uploads (" + str(len(files)) + "):**\n"
            for f in files[:50]:
                size = os.path.getsize(f)
                mtime = _time.strftime("%Y-%m-%d %H:%M", _time.localtime(os.path.getmtime(f)))
                result += "- " + os.path.basename(f) + " (" + str(round(size/1024,1)) + " KB, " + str(mtime) + ")\n"
            if len(files) > 50: result += "... and " + str(len(files)-50) + " more"
            return result
        elif operation == "delete_old":
            cutoff = _time.time() - (30 * 86400)
            deleted = 0
            for f in _glob.glob(os.path.join(uploads, pattern)):
                if os.path.getmtime(f) < cutoff:
                    os.remove(f); deleted += 1
            return "Deleted " + str(deleted) + " files older than 30 days."
        elif operation == "stats":
            files = _glob.glob(os.path.join(uploads, pattern))
            total_size = sum(os.path.getsize(f) for f in files)
            return "**Uploads Stats:**\n- Files: " + str(len(files)) + "\n- Total size: " + str(round(total_size/1024/1024,1)) + " MB"
        return json.dumps({"error": "Unknown operation: " + str(operation) + ". Use: list, delete_old, stats"})

    # --- v3.4.0: File Search ---
    async def file_search(self, query: str, file_type: str = "all") -> str:
        """Search for text across all generated files. file_type: all, xlsx, docx, pptx, pdf."""
        import glob as _glob
        results = []
        patterns = {"all": "*.*", "xlsx": "*.xlsx", "docx": "*.docx", "pptx": "*.pptx", "pdf": "*.pdf"}
        pattern = patterns.get(file_type, "*.*")
        for fpath in _glob.glob(os.path.join(_UPLOAD_DIR, pattern)):
            try:
                fname = os.path.basename(fpath)
                ext = os.path.splitext(fname)[1].lower()
                with open(fpath, 'rb') as f: fb = f.read()
                if ext == ".xlsx":
                    import openpyxl, io
                    wb = openpyxl.load_workbook(io.BytesIO(fb))
                    for ws in wb.worksheets:
                        for row in ws.iter_rows(values_only=True):
                            for cell in row:
                                if cell and query.lower() in str(cell).lower():
                                    results.append(fname + ": " + str(cell)[:100])
                                    break
                elif ext == ".docx":
                    from docx import Document
                    import io
                    doc = Document(io.BytesIO(fb))
                    for p in doc.paragraphs:
                        if query.lower() in p.text.lower():
                            results.append(fname + ": " + p.text[:100])
                            break
                elif ext == ".pptx":
                    from pptx import Presentation
                    import io
                    prs = Presentation(io.BytesIO(fb))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame and query.lower() in shape.text_frame.text.lower():
                                results.append(fname + ": " + shape.text_frame.text[:100])
                                break
            except Exception:
                pass
        if not results: return "No matches found for '" + str(query) + "'"
        return "**Search: " + str(query) + "** (" + str(len(results)) + " matches)\n" + "\n".join("- " + r for r in results[:20])

    # --- v3.4.0: Data Validation ---
    async def add_data_validation(self, file_id: str, sheet: str = "", col: str = "A", validation_type: str = "list", values: str = "", __user__=None, __request__=None) -> str:
        """Add data validation to an Excel column. Types: list, whole, decimal, date. values: comma-separated for list, or 'min,max' for numeric."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype not in ("xlsx","xls"): return json.dumps({"error": "Data validation only for Excel files"})
        from openpyxl import load_workbook
        from openpyxl.worksheet.datavalidation import DataValidation
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
        col_letter = col.upper().replace("COL","").replace("COLUMN","").strip()
        if not col_letter.isalpha(): col_letter = "A"
        max_row = ws.max_row or 100
        cell_range = col_letter + "2:" + col_letter + str(max_row)
        if validation_type == "list":
            vals = [v.strip() for v in values.split(",") if v.strip()]
            formula = '"' + ",".join(vals) + '"'
            dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        elif validation_type == "whole":
            parts = values.split(",")
            mn = int(parts[0]) if parts else 0
            mx = int(parts[1]) if len(parts) > 1 else 100
            dv = DataValidation(type="whole", operator="between", formula1=str(mn), formula2=str(mx))
        elif validation_type == "decimal":
            parts = values.split(",")
            mn = float(parts[0]) if parts else 0
            mx = float(parts[1]) if len(parts) > 1 else 100
            dv = DataValidation(type="decimal", operator="between", formula1=str(mn), formula2=str(mx))
        elif validation_type == "date":
            dv = DataValidation(type="date", operator="greaterThan", formula1="2000-01-01")
        else: return json.dumps({"error": "Unknown type: " + str(validation_type)})
        dv.error = "Invalid value"
        dv.errorTitle = "Validation Error"
        ws.add_data_validation(dv)
        dv.add(cell_range)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url: return "Data validation added to column " + str(col_letter) + ": [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})

    # --- v3.4.0: Named Ranges ---
    async def add_named_range(self, file_id: str, name: str, range_str: str = "", __user__=None, __request__=None) -> str:
        """Define a named range in Excel. range_str: 'A1:B10' or auto-detected from active sheet."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype not in ("xlsx","xls"): return json.dumps({"error": "Named ranges only for Excel"})
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        from openpyxl.workbook.defined_name import DefinedName
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        if not range_str:
            range_str = "A1:" + get_column_letter(ws.max_column) + str(ws.max_row)
        dn = DefinedName(name, attr_text=ws.title + "!$" + range_str.replace(":", ":$"))
        wb.defined_names.add(dn)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url: return "Named range '" + str(name) + "' = " + str(range_str) + ": [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})

    # --- v3.4.0: Slide Transitions ---
    async def add_slide_transitions(self, file_id: str, transition_type: str = "fade", duration: float = 0.5, __user__=None, __request__=None) -> str:
        """Add transitions to all slides in a PPTX. Types: fade, push, wipe, split, random."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype != "pptx": return json.dumps({"error": "Transitions only for PPTX"})
        from pptx import Presentation
        from pptx.util import Pt
        prs = Presentation(io.BytesIO(file_bytes))
        transitions = {"fade": 0, "push": 1, "wipe": 2, "split": 3, "random": 4}
        ttype = transitions.get(transition_type, 0)
        for slide in prs.slides:
            try:
                from pptx.oxml.ns import qn
                trans_elem = slide._element.find(qn('p:transition'))
                if trans_elem is None:
                    from lxml import etree
                    trans_elem = etree.SubElement(slide._element, qn('p:transition'))
                if transition_type == "fade":
                    etree.SubElement(trans_elem, qn('p:fade'))
                elif transition_type == "push":
                    etree.SubElement(trans_elem, qn('p:push'))
                elif transition_type == "wipe":
                    etree.SubElement(trans_elem, qn('p:wipe'))
                elif transition_type == "split":
                    etree.SubElement(trans_elem, qn('p:split'))
                trans_elem.set('advTm', str(int(duration * 1000)))
            except Exception:
                pass
        buf = io.BytesIO(); prs.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__, __user__=__user__)
        if url: return "Transitions added (" + str(transition_type) + "): [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})

    # --- v3.4.0: Export to HTML ---
    async def export_to_html(self, file_id: str, __user__=None, __request__=None) -> str:
        """Export any Office file to a styled HTML page."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"):
            content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Export not supported for " + str(ftype)})
        base = os.path.splitext(filename)[0]
        html_c = content.replace('\n', '<br>\n')
        html_c = re.sub(r'^### (.+)$', r'<h3>\1</h3>', html_c, flags=re.MULTILINE)
        html_c = re.sub(r'^## (.+)$', r'<h2>\1</h2>', html_c, flags=re.MULTILINE)
        html_c = re.sub(r'^# (.+)$', r'<h1>\1</h1>', html_c, flags=re.MULTILINE)
        html_c = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', html_c)
        html_c = re.sub(r'\*(.+?)\*', r'<em>\1</em>', html_c)
        html_c = re.sub(r'`(.+?)`', r'<code>\1</code>', html_c)
        html_c = re.sub(r'^- (.+)$', r'<li>\1</li>', html_c, flags=re.MULTILINE)
        html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>""" + str(base) + """</title>
<style>
body { font-family: 'Segoe UI', system-ui, sans-serif; max-width: 800px; margin: 40px auto; padding: 0 20px; line-height: 1.6; color: #333; background: #fafafa; }
h1 { color: #1a1a2e; border-bottom: 3px solid #e94560; padding-bottom: 10px; }
h2 { color: #16213e; margin-top: 30px; }
h3 { color: #0f3460; }
table { border-collapse: collapse; width: 100%; margin: 20px 0; }
th { background: #1a1a2e; color: white; padding: 10px; text-align: left; }
td { padding: 8px 10px; border-bottom: 1px solid #ddd; }
tr:nth-child(even) { background: #f5f5f5; }
code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-size: 0.9em; }
blockquote { border-left: 4px solid #e94560; margin: 20px 0; padding: 10px 20px; background: #fff5f5; }
</style>
</head>
<body>
""" + html_c + """
</body>
</html>"""
        html_bytes = html.encode('utf-8')
        url, fname = await self._save_and_link(html_bytes, base + ".html", __request__, __user__=__user__)
        if url: return "Exported to HTML: [" + str(fname) + "](" + str(url) + ")"
        return json.dumps({"error": "Could not save file"})


    # === v3.6.0: AI-Powered Features ===

    async def ai_analyze(self, file_id: str) -> str:
        """Extract document text for AI analysis. The LLM will analyze topics, sentiment, entities, and provide a summary."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"): content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        words = len(content.split())
        preview = content[:5000]
        return f"**Document: {filename}** ({words} words)\n\nAnalyze this document and provide:\n1. Main topics (3-5 bullet points)\n2. Sentiment (positive/negative/neutral)\n3. Key entities (people, companies, dates)\n4. Executive summary (2-3 sentences)\n\n```\n{preview}\n```" + ("\n\n... (truncated)" if len(content) > 5000 else "")

    async def smart_fill(self, file_id: str, section: str, instruction: str, __user__=None, __request__=None) -> str:
        """Fill a document section using AI based on instructions. The LLM will generate content for the specified section."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"): content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        return f"**Smart Fill: {filename}**\n\nSection to fill: **{section}**\nInstructions: {instruction}\n\nCurrent document content:\n```\n{content[:3000]}\n```\n\nPlease generate the content for the '{section}' section based on the instructions and existing document context."

    async def grammar_check(self, file_id: str) -> str:
        """Check document for grammar and style issues. The LLM will provide corrections."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"): content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        return f"**Grammar Check: {filename}**\n\nReview this document for:\n1. Grammar errors\n2. Spelling mistakes\n3. Style inconsistencies\n4. Passive voice overuse\n5. Readability issues\n\nProvide corrections with line references:\n\n```\n{content[:4000]}\n```"

    async def translate_document(self, file_id: str, target_language: str, __user__=None, __request__=None) -> str:
        """Translate a document to another language. The LLM will translate while preserving structure."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"): content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        return f"**Translate to {target_language}: {filename}**\n\nTranslate the following document to {target_language}. Preserve all formatting markers (# for headings, | for tables, - for bullets). Keep numbers, dates, and proper names unchanged.\n\n```\n{content[:4000]}\n```"

    async def classify_document(self, file_id: str) -> str:
        """Auto-classify a document by type, theme, and department."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype in ("xlsx","xls"): content = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
        elif ftype == "docx": content = self._read_docx(file_bytes, filename)
        elif ftype == "pptx": content = self._read_pptx(file_bytes, filename)
        elif ftype in ("odt","ods","odp"): content = _read_odf(file_bytes, filename)
        else: return json.dumps({"error": "Unsupported format"})
        return f"**Classify: {filename}**\n\nAnalyze this document and provide:\n1. Document type (report, proposal, invoice, contract, presentation, spreadsheet, letter, memo, manual, other)\n2. Primary theme/topic\n3. Department (finance, HR, marketing, engineering, sales, legal, operations, other)\n4. Confidentiality level (public, internal, confidential, restricted)\n5. Suggested tags (3-5 keywords)\n\n```\n{content[:2000]}\n```"

    async def smart_template(self, name: str, description: str, __user__=None, __request__=None) -> str:
        """Generate a document from a smart template that adapts to the conversation context."""
        templates = json.loads(self.valves.templates or "{}")
        if name in templates:
            content = templates[name]
            return await self.generate_document(content, name, __user__=__user__, __request__=__request__)
        return f"**Smart Template: {name}**\n\nDescription: {description}\n\nGenerate a professional document template for '{name}' with the following sections and {placeholders} for customization. Use markdown format with # headings, - bullets, and | tables."

    # === v3.6.0: Data Manipulation ===

    async def add_pivot_table(self, file_id: str, rows_field: str = "", cols_field: str = "", data_field: str = "", aggregate: str = "sum", __user__=None, __request__=None) -> str:
        """Create a pivot table in Excel. aggregate: sum, count, average, min, max."""
        import io
        from collections import defaultdict
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype not in ("xlsx","xls"): return json.dumps({"error": "Pivot tables only for Excel"})
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        headers = [str(c.value or '') for c in ws[1]]
        if not rows_field:
            return f"**Available fields for pivot:** {', '.join(headers[:10])}\n\nUse: add_pivot_table(file_id, rows_field='FieldName', data_field='FieldName')"
        # Find column indices
        try:
            row_idx = headers.index(rows_field)
        except ValueError:
            return json.dumps({"error": f"Field '{rows_field}' not found in headers: {headers[:20]}"})
        data_idx = None
        if data_field:
            try:
                data_idx = headers.index(data_field)
            except ValueError:
                return json.dumps({"error": f"Field '{data_field}' not found in headers: {headers[:20]}"})
        # Read and aggregate data
        agg_map = defaultdict(list)
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_key = str(row[row_idx]) if row[row_idx] is not None else "(blank)"
            if data_idx is not None and row[data_idx] is not None:
                try:
                    agg_map[row_key].append(float(row[data_idx]))
                except (ValueError, TypeError):
                    agg_map[row_key].append(0.0)
            elif data_idx is None:
                agg_map[row_key].append(1)  # count mode
        # Compute aggregate
        pivot_ws = wb.create_sheet("Pivot")
        pivot_ws.title = "Pivot"
        agg_name = aggregate.lower()
        if agg_name == "count":
            pivot_ws.cell(row=1, column=1, value=rows_field)
            pivot_ws.cell(row=1, column=2, value="Count")
            for r, (key, vals) in enumerate(sorted(agg_map.items()), 2):
                pivot_ws.cell(row=r, column=1, value=key)
                pivot_ws.cell(row=r, column=2, value=len(vals))
        else:
            pivot_ws.cell(row=1, column=1, value=rows_field)
            pivot_ws.cell(row=1, column=2, value=f"{agg_name} of {data_field}")
            for r, (key, vals) in enumerate(sorted(agg_map.items()), 2):
                pivot_ws.cell(row=r, column=1, value=key)
                if agg_name == "sum":
                    pivot_ws.cell(row=r, column=2, value=sum(vals))
                elif agg_name == "average" or agg_name == "avg":
                    pivot_ws.cell(row=r, column=2, value=sum(vals) / len(vals) if vals else 0)
                elif agg_name == "min":
                    pivot_ws.cell(row=r, column=2, value=min(vals) if vals else 0)
                elif agg_name == "max":
                    pivot_ws.cell(row=r, column=2, value=max(vals) if vals else 0)
                else:
                    pivot_ws.cell(row=r, column=2, value=sum(vals))
        result = f"Pivot table created in sheet 'Pivot' ({len(agg_map)} rows). Fields: rows={rows_field}, data={data_field}, aggregate={aggregate}"
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
        if url: return f"{result}: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    async def sql_to_spreadsheet(self, query: str, output_filename: str = "query_results", __user__=None, __request__=None) -> str:
        """Execute a SQL query on the local SQLite database and export results to Excel."""
        import io
        conn2 = sqlite3.connect(_DB_PATH)
        try:
            conn2.row_factory = sqlite3.Row
            rows = conn2.execute(query).fetchall()
        except Exception as e:
            return json.dumps({"error": f"SQL error: {str(e)}"})
        finally:
            conn2.close()
        if not rows: return "Query returned no results."
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook(); ws = wb.active
        headers = list(rows[0].keys())
        for j, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=j, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        for i, row in enumerate(rows, 2):
            for j, key in enumerate(headers, 1):
                ws.cell(row=i, column=j, value=row[key])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), f"{output_filename}.xlsx", __request__)
        if url: return f"Query results ({len(rows)} rows): [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    async def fill_pdf_form(self, file_id: str, field_values: str, __user__=None, __request__=None) -> str:
        """Fill a PDF form with values. field_values: 'field1=value1,field2=value2'."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype != "pdf": return json.dumps({"error": "PDF form filling only for PDF files"})
        try:
            import fitz
            pdf = fitz.open(stream=file_bytes, filetype="pdf")
            pairs = {}
            for pair in field_values.split(","):
                if "=" in pair:
                    k, v = pair.split("=", 1)
                    pairs[k.strip()] = v.strip()
            filled = 0
            for page in pdf:
                for widget in page.widgets():
                    if widget.field_name in pairs:
                        widget.field_value = pairs[widget.field_name]
                        widget.update()
                        filled += 1
            buf = io.BytesIO(); pdf.save(buf); pdf.close(); buf.seek(0)
            url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
            if url: return f"Filled {filled} field(s): [{fname}]({url})"
            return json.dumps({"error": "Could not save file"})
        except ImportError:
            return json.dumps({"error": "PyMuPDF not installed"})
        except Exception as e:
            return json.dumps({"error": str(e)})

    async def convert_data(self, file_id: str, target_format: str, __user__=None, __request__=None) -> str:
        """Convert between CSV, JSON, and XML formats. target_format: csv, json, xml."""
        import io, csv as _csv
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        base = os.path.splitext(filename)[0]
        text = file_bytes.decode('utf-8', errors='replace')
        if target_format == "json":
            try:
                if ftype == "csv" or filename.endswith('.csv'):
                    reader = _csv.DictReader(io.StringIO(text))
                    data = list(reader)
                elif filename.endswith('.xml'):
                    import xml.etree.ElementTree as ET
                    root = ET.fromstring(text)
                    data = [{child.tag: child.text for child in elem} for elem in root]
                else:
                    data = [{"content": text}]
                result = json.dumps(data, indent=2, ensure_ascii=False)
                ext = ".json"
            except Exception as e:
                return json.dumps({"error": f"Conversion failed: {str(e)}"})
        elif target_format == "csv":
            try:
                if ftype == "json" or filename.endswith('.json'):
                    data = json.loads(text)
                    if isinstance(data, list) and data:
                        out = io.StringIO()
                        writer = _csv.DictWriter(out, fieldnames=data[0].keys())
                        writer.writeheader(); writer.writerows(data)
                        result = out.getvalue()
                    else:
                        result = text
                else:
                    result = text
                ext = ".csv"
            except Exception as e:
                return json.dumps({"error": f"Conversion failed: {str(e)}"})
        elif target_format == "xml":
            try:
                if ftype == "json" or filename.endswith('.json'):
                    data = json.loads(text)
                    import xml.etree.ElementTree as ET
                    root = ET.Element("root")
                    for item in (data if isinstance(data, list) else [data]):
                        elem = ET.SubElement(root, "item")
                        for k, v in (item.items() if isinstance(item, dict) else {"value": str(item)}.items()):
                            child = ET.SubElement(elem, k)
                            child.text = str(v)
                    result = ET.tostring(root, encoding='unicode')
                else:
                    result = text
                ext = ".xml"
            except Exception as e:
                return json.dumps({"error": f"Conversion failed: {str(e)}"})
        else:
            return json.dumps({"error": f"Unsupported target format: {target_format}. Use: csv, json, xml"})
        result_bytes = result.encode('utf-8')
        url, fname = await self._save_and_link(result_bytes, f"{base}{ext}", __request__)
        if url: return f"Converted to {target_format.upper()}: [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # === v3.6.0: Enterprise Features ===

    async def compliance_check(self, file_id: str, standard: str = "gdpr") -> str:
        """Check document for compliance issues. standards: gdpr, accessibility, branding, all."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        issues = []
        if standard in ("gdpr", "all"):
            text = ""
            if ftype == "docx":
                from docx import Document; import io
                doc = Document(io.BytesIO(file_bytes))
                text = " ".join(p.text for p in doc.paragraphs)
            gdpr_keywords = ["email", "phone", "address", "name", "birth", "passport", "ssn", "tax id", "iban", "credit card", "ip address", "cookie"]
            found = [k for k in gdpr_keywords if k in text.lower()]
            if found: issues.append(f"GDPR: Personal data detected: {', '.join(found)}. Ensure consent and data processing agreement.")
        if standard in ("accessibility", "all"):
            if ftype == "docx":
                from docx import Document; import io
                doc = Document(io.BytesIO(file_bytes))
                headings = [p for p in doc.paragraphs if p.style.name.startswith('Heading')]
                if not headings: issues.append("Accessibility: No headings found. Add heading structure.")
                images = len([r for r in doc.part.rels.values() if "image" in r.reltype])
                if images > 0: issues.append(f"Accessibility: {images} image(s) found. Ensure alt text is provided.")
        if standard in ("branding", "all"):
            if ftype == "docx":
                from docx import Document; import io
                doc = Document(io.BytesIO(file_bytes))
                text = " ".join(p.text for p in doc.paragraphs)
                if "confidential" not in text.lower() and "draft" not in text.lower():
                    issues.append("Branding: No confidentiality marking found. Consider adding DRAFT or CONFIDENTIAL watermark.")
        if not issues: return f"Compliance check passed for {filename}. No issues found."
        return f"**Compliance Report: {filename}**\n\n" + "\n".join(f"- {i}" for i in issues)

    async def audit_log(self, action: str = "list", limit: int = 20) -> str:
        """View or manage the audit trail of document operations."""
        import time as _time
        conn2 = sqlite3.connect(_DB_PATH)
        try:
            conn2.row_factory = sqlite3.Row
            if action == "list":
                rows = conn2.execute("SELECT filename, created_at FROM file WHERE meta LIKE '%office-plugin%' ORDER BY created_at DESC LIMIT ?", (limit,)).fetchall()
                if not rows: return "No audit records found."
                result = f"**Audit Trail (last {len(rows)}):**\n"
                for r in rows:
                    ts = _time.strftime("%Y-%m-%d %H:%M", _time.localtime(r["created_at"])) if r["created_at"] else "unknown"
                    result += f"- {ts}: {r['filename']}\n"
                return result
            elif action == "stats":
                total = conn2.execute("SELECT COUNT(*) FROM file WHERE meta LIKE '%office-plugin%'").fetchone()[0]
                today = conn2.execute("SELECT COUNT(*) FROM file WHERE meta LIKE '%office-plugin%' AND date(created_at, 'unixepoch') = date('now')").fetchone()[0]
                return f"**Audit Stats:**\n- Total files: {total}\n- Today: {today}"
            return json.dumps({"error": f"Unknown action: {action}. Use: list, stats"})
        finally:
            conn2.close()

    async def retention_policy(self, policy: str = "view", days: int = 90, file_type: str = "all") -> str:
        """Manage document retention policies. policy: view, set, apply."""
        if policy == "view":
            sched = json.loads(self.valves.cleanup_schedule or "{}")
            if sched.get("enabled"):
                return f"Retention policy active: delete files older than {sched.get('days_old', 30)} days, every {sched.get('interval_hours', 24)} hours."
            return "No retention policy active. Use: retention_policy(policy='set', days=90)"
        elif policy == "set":
            self.valves.cleanup_schedule = json.dumps({"days_old": days, "interval_hours": 24, "enabled": True, "file_type": file_type})
            return f"Retention policy set: delete {file_type} files older than {days} days."
        elif policy == "apply":
            return await self.cleanup_files(days_old=days)
        return json.dumps({"error": f"Unknown policy: {policy}. Use: view, set, apply"})

    async def scheduled_report(self, action: str = "list", name: str = "", schedule: str = "", template: str = "", __user__=None, __request__=None) -> str:
        """Manage scheduled reports. action: list, create, delete, run."""
        reports = json.loads(self.valves.templates or "{}")
        scheduled = json.loads(self.valves.cleanup_schedule or "{}")
        if action == "list":
            sched_reports = {k: v for k, v in reports.items() if k.startswith("_scheduled_")}
            if not sched_reports: return "No scheduled reports."
            result = "**Scheduled Reports:**\n"
            for k, v in sched_reports.items():
                result += f"- {k.replace('_scheduled_', '')}: {v[:50]}...\n"
            return result
        elif action == "create":
            reports[f"_scheduled_{name}"] = json.dumps({"template": template, "schedule": schedule})
            self.valves.templates = json.dumps(reports)
            return f"Scheduled report '{name}' created."
        elif action == "delete":
            key = f"_scheduled_{name}"
            if key in reports:
                del reports[key]
                self.valves.templates = json.dumps(reports)
                return f"Scheduled report '{name}' deleted."
            return f"Report '{name}' not found."
        elif action == "run":
            key = f"_scheduled_{name}"
            if key not in reports: return f"Report '{name}' not found."
            cfg = json.loads(reports[key])
            return await self.generate_document(cfg.get("template", ""), name, __user__=__user__, __request__=__request__)
        return json.dumps({"error": f"Unknown action: {action}"})

    async def document_assembly(self, template_name: str, data_file_id: str, output_prefix: str = "assembled", __user__=None, __request__=None) -> str:
        """Assemble multiple documents from a template and data source."""
        # Load template content
        templates = json.loads(self.valves.templates or "{}")
        if template_name not in templates:
            return f"Template '{template_name}' not found. Available: {', '.join(templates.keys())}"
        template_content = templates[template_name]
        
        # Load data
        d_bytes, d_name, d_type = self._resolve_file(data_file_id)
        if not d_bytes:
            return json.dumps({"error": f"Data file not found: {data_file_id}"})
        
        # Parse data rows
        import csv as _csv, io
        rows = []
        if d_type in ("xlsx", "xls"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(d_bytes))
            ws = wb.active
            headers = [str(c.value or '') for c in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, [str(v or '') for v in row])))
        else:
            reader = _csv.DictReader(io.StringIO(d_bytes.decode('utf-8')))
            rows = list(reader)
        
        if not rows:
            return json.dumps({"error": "No data rows found"})
        
        # Generate one document per data row
        results = []
        for i, row in enumerate(rows):
            content = template_content
            for key, value in row.items():
                content = content.replace(f"{{{key}}}", str(value))
            result = await self.generate_document(content, f"{output_prefix}_{i+1}", __user__=__user__, __request__=__request__)
            results.append(result)
        
        return f"Assembled {len(results)} documents from template '{template_name}'."

    async def conditional_format(self, file_id: str, rules: str, __user__=None, __request__=None) -> str:
        """Apply conditional formatting rules to Excel. rules: 'col:A,op:>,val:100,color:27AE60;col:B,op:<,val:0,color:E74C3C'."""
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        if ftype not in ("xlsx","xls"): return json.dumps({"error": "Conditional formatting only for Excel"})
        from openpyxl import load_workbook
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import PatternFill
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        applied = 0
        for rule_str in rules.split(";"):
            parts = {}
            for p in rule_str.split(","):
                if ":" in p:
                    k, v = p.split(":", 1)
                    parts[k.strip()] = v.strip()
            if "col" not in parts: continue
            col = parts["col"].upper()
            op_map = {">": "greaterThan", "<": "lessThan", ">=": "greaterThanOrEqual", "<=": "lessThanOrEqual", "=": "equal", "!=": "notEqual"}
            op = op_map.get(parts.get("op", ">"), "greaterThan")
            val = parts.get("val", "0")
            color = parts.get("color", "27AE60")
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            max_row = ws.max_row or 100
            cell_range = f"{col}2:{col}{max_row}"
            ws.conditional_formatting.add(cell_range, CellIsRule(operator=op, formula=[val], fill=fill))
            applied += 1
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
        if url: return f"Applied {applied} conditional formatting rule(s): [{fname}]({url})"
        return json.dumps({"error": "Could not save file"})

    # === v3.6.0: Collaboration Features ===

    async def add_comment(self, file_id: str, text: str, author: str = "Reviewer", paragraph_index: int = 0, cell_ref: str = "A1", slide_num: int = 1, __user__=None, __request__=None) -> str:
        """Add a review comment to a Word, Excel, or PowerPoint file.

        Args:
            file_id: File ID to comment on
            text: Comment text
            author: Name shown in the comment (e.g., "Sergio Pedro")
            paragraph_index: For DOCX: paragraph index to attach comment to (default 0)
            cell_ref: For XLSX: cell reference (default "A1")
            slide_num: For PPTX: slide number (default 1)
        """
        import io
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})

        if ftype == "docx":
            from docx import Document
            doc = Document(io.BytesIO(file_bytes))

            # Ensure we have a paragraph (with at least one run) to anchor the comment to
            if not doc.paragraphs:
                doc.add_paragraph("")
            if paragraph_index >= len(doc.paragraphs):
                paragraph_index = 0
            para = doc.paragraphs[paragraph_index]
            if not para.runs:
                para.add_run("")

            # python-docx's Document.add_comment(runs, text=..., author=...) anchors the
            # comment to a Run or sequence of Runs -- it does NOT take the comment text as
            # its first argument (that's the `text=` kwarg).
            doc.add_comment(para.runs, text=text, author=author)
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
            if url: return f"Comment added by {author} on paragraph {paragraph_index}: [{fname}]({url})"
            return json.dumps({"error": "Could not save file"})

        elif ftype == "xlsx":
            from openpyxl import load_workbook
            from openpyxl.comments import Comment
            wb = load_workbook(io.BytesIO(file_bytes))
            ws = wb.active
            comment = Comment(text, author)
            ws[cell_ref].comment = comment
            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
            url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
            if url: return f"Comment added by {author} on cell {cell_ref}: [{fname}]({url})"
            return json.dumps({"error": "Could not save file"})

        elif ftype == "pptx":
            try:
                import zipfile
                import uuid as _uuid
                from datetime import datetime, timezone
                from lxml import etree


                slide_name = "ppt/slides/slide%d.xml" % slide_num
                rels_name = "ppt/slides/_rels/slide%d.xml.rels" % slide_num
                modern_name = "ppt/comments/commentModern%d.xml" % slide_num
                authors_name = "ppt/commentsAuthors.xml"

                with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                    entries = {name: z.read(name) for name in z.namelist()}

                if slide_name not in entries:
                    return json.dumps({"error": "Slide %d not found in presentation" % slide_num})

                etree.register_namespace("p", _P_NS)
                etree.register_namespace("p14", _P14_NS)
                etree.register_namespace("r", _R_NS)

                utc_now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000")
                para_guid = "{" + str(_uuid.uuid4()).upper() + "}"

                # --- modern comment part: compute next comment idx ---
                if modern_name in entries:
                    mroot = etree.fromstring(entries[modern_name])
                    max_idx = 0
                    for el in mroot:
                        if el.tag == "{%s}cm" % _P_NS:
                            try:
                                max_idx = max(max_idx, int(el.get("idx") or 0))
                            except (TypeError, ValueError):
                                pass
                    next_idx = max_idx + 1
                else:
                    mroot = etree.Element("{%s}cmLst" % _P_NS)
                    next_idx = 1

                # --- commentsAuthors part: resolve or create author id ---
                if authors_name in entries:
                    aroot = etree.fromstring(entries[authors_name])
                else:
                    aroot = etree.Element("{%s}cmAuthorLst" % _P14_NS)

                author_id = None
                existing_ids = []
                author_els = []
                for el in aroot:
                    if el.tag == "{%s}cmAuthor" % _P14_NS:
                        author_els.append(el)
                        try:
                            existing_ids.append(int(el.get("id") or 0))
                        except (TypeError, ValueError):
                            existing_ids.append(0)
                        if el.get("name") == author:
                            author_id = el.get("id")

                if author_id is not None:
                    for el in author_els:
                        if el.get("name") == author:
                            el.set("lastIdx", str(next_idx))
                            break
                else:
                    new_id = (max(existing_ids) + 1) if existing_ids else 0
                    initials = "".join(w[0] for w in re.split(r"[\s._\-]+", author.strip()) if w)[:4].upper() or "AU"
                    new_author = etree.SubElement(aroot, "{%s}cmAuthor" % _P14_NS)
                    new_author.set("id", str(new_id))
                    new_author.set("name", author)
                    new_author.set("initials", initials)
                    new_author.set("lastIdx", "1")
                    new_author.set("clrIndex", str(len(author_els) % 14))
                    author_id = str(new_id)

                # --- append modern comment ---
                cm = etree.SubElement(mroot, "{%s}cm" % _P_NS)
                cm.set("authorId", author_id)
                cm.set("dt", utc_now)
                cm.set("idx", str(next_idx))
                pos = etree.SubElement(cm, "{%s}pos" % _P_NS)
                pos.set("x", "100")
                pos.set("y", "100")
                cm_txt = etree.SubElement(cm, "{%s}text" % _P_NS)
                cm_txt.text = text
                cm_extLst = etree.SubElement(cm, "{%s}extLst" % _P_NS)
                cm_ext = etree.SubElement(cm_extLst, "{%s}ext" % _P_NS)
                cm_ext.set("uri", "{D6B160D4-4F5E-48AF-9E34-8E18A86A1F0A}")
                comment_ex = etree.SubElement(cm_ext, "{%s}commentEx" % _P14_NS)
                comment_ex.set("paraId", para_guid)
                comment_ex.set("dt", utc_now)
                comment_ex.set("parentIdx", "0")
                ce_txt = etree.SubElement(comment_ex, "{%s}text" % _P14_NS)
                ce_txt.text = text
                ce_author = etree.SubElement(comment_ex, "{%s}authorId" % _P14_NS)
                ce_author.text = author_id

                # --- [Content_Types].xml: add overrides if missing ---
                ct_root = etree.fromstring(entries["[Content_Types].xml"])
                has_authors_ct = any(
                    el.tag == "{%s}Override" % _CT_NS and el.get("PartName") == "/ppt/commentsAuthors.xml"
                    for el in ct_root
                )
                if not has_authors_ct:
                    o = etree.SubElement(ct_root, "{%s}Override" % _CT_NS)
                    o.set("PartName", "/ppt/commentsAuthors.xml")
                    o.set("ContentType", _CT_AUTHORS)
                has_modern_ct = any(
                    el.tag == "{%s}Override" % _CT_NS and el.get("PartName") == "/ppt/comments/commentModern%d.xml" % slide_num
                    for el in ct_root
                )
                if not has_modern_ct:
                    o = etree.SubElement(ct_root, "{%s}Override" % _CT_NS)
                    o.set("PartName", "/ppt/comments/commentModern%d.xml" % slide_num)
                    o.set("ContentType", _CT_MODERN)

                # --- slide rels: add commentsModern relationship ---
                if rels_name in entries:
                    rroot = etree.fromstring(entries[rels_name])
                else:
                    rroot = etree.Element("{%s}Relationships" % _PKG_REL_NS, nsmap={None: _PKG_REL_NS})
                max_rid = 0
                for el in rroot:
                    rid = el.get("Id") or ""
                    m = re.match(r"rId(\d+)$", rid)
                    if m:
                        max_rid = max(max_rid, int(m.group(1)))
                new_rid = max_rid + 1
                target = "../comments/commentModern%d.xml" % slide_num
                # Deduplicate: reuse existing Relationship to same target (avoid duplicate on repeat calls)
                existing_rel = None
                for child in rroot:
                    if child.tag == "{%s}Relationship" % _PKG_REL_NS and child.get("Type") == _CM_REL_TYPE and child.get("Target") == target:
                        existing_rel = child
                        break
                if existing_rel is not None:
                    existing_rel.set("Id", "rId%d" % new_rid)
                else:
                    rel_el = etree.SubElement(rroot, "{%s}Relationship" % _PKG_REL_NS)
                    rel_el.set("Id", "rId%d" % new_rid)
                    rel_el.set("Type", _CM_REL_TYPE)
                    rel_el.set("Target", target)

                # --- slide XML: attach commentRel extension ---
                sroot = etree.fromstring(entries[slide_name])
                extLst = None
                for child in sroot:
                    if child.tag == "{%s}extLst" % _P_NS:
                        extLst = child
                        break
                if extLst is None:
                    extLst = etree.SubElement(sroot, "{%s}extLst" % _P_NS)
                # Deduplicate: reuse existing commentRel ext if present (avoid duplicate URI on repeat calls)
                slide_ext = None
                for child in extLst:
                    if child.tag == "{%s}ext" % _P_NS and child.get("uri") == "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}":
                        slide_ext = child
                        break
                if slide_ext is None:
                    slide_ext = etree.SubElement(extLst, "{%s}ext" % _P_NS)
                    slide_ext.set("uri", "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}")
                # Update or create commentRel child
                comment_rel = None
                for child in slide_ext:
                    if child.tag == "{%s}commentRel" % _P14_NS:
                        comment_rel = child
                        break
                if comment_rel is None:
                    comment_rel = etree.SubElement(slide_ext, "{%s}commentRel" % _P14_NS)
                comment_rel.set("{%s}id" % _R_NS, "rId%d" % new_rid)

                # --- serialize modified parts ---
                def _ser(root):
                    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

                entries[modern_name] = _ser(mroot)
                entries[authors_name] = _ser(aroot)
                entries[rels_name] = _ser(rroot)
                entries[slide_name] = _ser(sroot)
                entries["[Content_Types].xml"] = _ser(ct_root)

                # --- rebuild zip (OPC: [Content_Types].xml first) ---
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                    zout.writestr("[Content_Types].xml", entries["[Content_Types].xml"])
                    for name in entries:
                        if name != "[Content_Types].xml":
                            zout.writestr(name, entries[name])
                buf.seek(0)

                url, fname = await self._save_and_link(buf.getvalue(), filename, __request__)
                if url:
                    return f"Comment added by {author} on slide {slide_num}: [{fname}]({url})"
                return json.dumps({"error": "Could not save file"})
            except Exception as e:
                return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

        else:
            return json.dumps({"error": f"Comments not supported for {ftype}. Supported: DOCX, XLSX, PPTX."})

    async def version_diff(self, file_id: str, version_label: str = "") -> str:
        """Show differences between current file and a previous version."""
        file_bytes, filename, ftype = self._resolve_file(file_id)
        if not file_bytes: return json.dumps({"error": "File not found"})
        base = os.path.splitext(filename)[0]
        ext = os.path.splitext(filename)[1]
        import glob as _glob, time as _time
        pattern = os.path.join(_UPLOAD_DIR, f"{base}_v*{ext}")
        versions = sorted(_glob.glob(pattern), reverse=True)
        if not versions: return f"No previous versions found for {filename}. Use version_file() to create versions."
        if version_label:
            versions = [v for v in versions if version_label in v]
            if not versions: return f"No version matching '{version_label}' found."
        latest = versions[0]
        vname = os.path.basename(latest)
        vtime = _time.strftime("%Y-%m-%d %H:%M", _time.localtime(os.path.getmtime(latest)))
        with open(latest, 'rb') as f: vbytes = f.read()
        if ftype in ("xlsx","xls"):
            curr = self._read_xlsx(file_bytes, filename) if ftype == "xlsx" else self._read_xls(file_bytes, filename)
            prev = self._read_xlsx(vbytes, vname) if ftype == "xlsx" else self._read_xls(vbytes, vname)
        elif ftype == "docx":
            curr = self._read_docx(file_bytes, filename)
            prev = self._read_docx(vbytes, vname)
        elif ftype == "pptx":
            curr = self._read_pptx(file_bytes, filename)
            prev = self._read_pptx(vbytes, vname)
        else:
            return json.dumps({"error": f"Diff not supported for {ftype}"})
        cl = curr.split('\n'); pl = prev.split('\n')
        added = sum(1 for l in cl if l not in pl)
        removed = sum(1 for l in pl if l not in cl)
        return f"**Version Diff: {filename} vs {vname}** ({vtime})\n- Lines added: {added}\n- Lines removed: {removed}\n- Current: {len(cl)} lines\n- Previous: {len(pl)} lines"

    async def webhook_trigger(self, event: str = "test", url: str = "", file_id: str = "") -> str:
        """Trigger a webhook on document events. event: test, created, edited, deleted."""
        if not url:
            return json.dumps({"error": "url parameter required"})
        import urllib.request as _urllib
        payload = json.dumps({"event": event, "file_id": file_id, "timestamp": __import__('time').time()}).encode()
        try:
            req = _urllib.Request(url, data=payload, headers={"Content-Type": "application/json"})
            resp = _urllib.urlopen(req, timeout=10)
            return f"Webhook sent to {url}: HTTP {resp.getcode()}"
        except Exception as e:
            return json.dumps({"error": f"Webhook failed: {str(e)}"})

    async def import_from_api(self, url: str, data_path: str = "", output_filename: str = "api_data", __user__=None, __request__=None) -> str:
        """Import data from a REST API and export to Excel."""
        import urllib.request as _urllib, io
        try:
            req = _urllib.Request(url, headers={"Accept": "application/json", "User-Agent": "OpenWebUI/1.0"})
            resp = _urllib.urlopen(req, timeout=15)
            raw = json.loads(resp.read())
        except Exception as e:
            return json.dumps({"error": f"API request failed: {str(e)}"})
        data = raw
        if data_path:
            for key in data_path.split("."):
                if isinstance(data, dict):
                    data = data.get(key, [])
                elif isinstance(data, list) and key.isdigit():
                    data = data[int(key)]
        if not isinstance(data, list):
            data = [data] if isinstance(data, dict) else [{"value": str(data)}]
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        wb = Workbook(); ws = wb.active
        if data and isinstance(data[0], dict):
            headers = list(data[0].keys())
            for j, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=j, value=h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            for i, item in enumerate(data, 2):
                for j, key in enumerate(headers, 1):
                    ws.cell(row=i, column=j, value=str(item.get(key, "")))
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        url_out, fname = await self._save_and_link(buf.getvalue(), f"{output_filename}.xlsx", __request__)
        if url_out: return f"Imported {len(data)} records from API: [{fname}]({url_out})"
        return json.dumps({"error": "Could not save file"})

