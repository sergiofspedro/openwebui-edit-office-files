"""Edit Office Files — path and namespace constants."""
import os
import platform
import sys

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
