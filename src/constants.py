"""Edit Office Files — path and namespace constants."""
import os
import platform
import sys

# ---------------------------------------------------------------------------
# Resolution chain for the Open WebUI data directory
# ---------------------------------------------------------------------------
# Open WebUI's canonical env var is `DATA_DIR` (defined in
# backend/open_webui/env.py:222). The legacy plugin var `OPEN_WEBUI_DATA_DIR`
# is NOT recognized by upstream OWUI and is kept here only for backward
# compatibility with older deployments that may have set it manually.
#
# Resolution order (first hit wins):
#   1. `DATA_DIR` env var (canonical OWUI, set in upstream env.py)
#   2. `OPEN_WEBUI_DATA_DIR` env var (legacy plugin-only var)
#   3. `/app/backend/data` — official OWUI Docker image default (Dockerfile)
#   4. `/data` — common bind-mount path used by self-hosted setups
#   5. OS-specific userspace fallback (XDG/APPDATA/~/Library/~/)
# ---------------------------------------------------------------------------

def _candidate_owui_data_dirs() -> list[str]:
    """Return a list of candidate Open WebUI data directories, in priority order.

    Only directories that *might* exist are returned. The first one that exists
    (and contains webui.db, or is writable) is used by the callers.
    """
    candidates: list[str] = []

    # 1. Canonical OWUI env var
    data_dir = os.environ.get("DATA_DIR", "").strip()
    if data_dir:
        candidates.append(data_dir)

    # 2. Legacy plugin-only env var
    legacy = os.environ.get("OPEN_WEBUI_DATA_DIR", "").strip()
    if legacy and legacy not in candidates:
        candidates.append(legacy)

    # 3. Official OWUI Docker default (upstream Dockerfile creates this)
    if "/app/backend/data" not in candidates and not os.path.exists("/app/backend"):
        pass  # not running inside the OWUI container — skip
    if "/app/backend/data" not in candidates:
        candidates.append("/app/backend/data")

    # 4. Common bind-mount path
    if "/data" not in candidates:
        candidates.append("/data")

    # 5. OS-specific userspace fallback
    home = os.path.expanduser("~")
    system = platform.system()
    if system == "Windows":
        candidates.append(os.path.join(os.environ.get("APPDATA", home), "open-webui", "data"))
    elif system == "Darwin":
        candidates.append(os.path.join(home, "Library", "Application Support", "open-webui", "data"))
    else:  # Linux / Unix
        xdg_data_home = os.environ.get("XDG_DATA_HOME", "").strip()
        if xdg_data_home:
            candidates.append(os.path.join(xdg_data_home, "open-webui", "data"))
        candidates.append(os.path.join(home, ".open-webui", "data"))

    return candidates


def _get_owui_data_dir() -> str:
    """Return the Open WebUI data directory for the current OS.

    Resolution order (first hit wins):
        1. `DATA_DIR` env var (canonical OWUI)
        2. `OPEN_WEBUI_DATA_DIR` env var (legacy)
        3. `/app/backend/data` (official OWUI Docker default)
        4. `/data` (common bind-mount)
        5. OS-specific userspace fallback

    A directory is "good" if it exists. The first existing candidate is returned.
    If none exist, the canonical default (`/app/backend/data`) is returned so
    that callers can still attempt operations and surface a meaningful error.
    """
    candidates = _candidate_owui_data_dirs()
    for d in candidates:
        if d and os.path.isdir(d):
            return d
    # None exist — return canonical default (caller will get a clean error)
    return "/app/backend/data" if not os.name == "nt" else candidates[-1]


def _get_owui_uploads_dir() -> str:
    """Return the Open WebUI uploads directory for the current OS.

    The uploads directory is conventionally `${DATA_DIR}/uploads` in upstream
    OWUI (env.py:222, DATA_DIR + `/uploads`). For backwards compatibility,
    we also probe the legacy userspace paths where a prior v4.x might have
    written files.
    """
    # First try the canonical location: <DATA_DIR>/uploads
    data_dir = os.environ.get("DATA_DIR", "").strip()
    if data_dir:
        cand = os.path.join(data_dir, "uploads")
        if os.path.isdir(cand):
            return cand
        if os.path.isdir(data_dir):
            # DATA_DIR exists but no uploads/ subdir yet — return the expected
            # path so callers can create it. This is the common cold-start case.
            return cand

    legacy = os.environ.get("OPEN_WEBUI_DATA_DIR", "").strip()
    if legacy:
        return os.path.join(legacy, "data", "uploads")

    # Official Docker default
    docker_uploads = "/app/backend/data/uploads"
    if os.path.isdir(docker_uploads):
        return docker_uploads
    if os.path.isdir("/app/backend/data"):
        return docker_uploads

    # Common bind mount
    bind_uploads = "/data/uploads"
    if os.path.isdir(bind_uploads):
        return bind_uploads

    # OS-specific userspace fallback
    home = os.path.expanduser("~")
    system = platform.system()
    if system == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "open-webui", "data", "uploads")
    if system == "Darwin":
        return os.path.join(home, "Library", "Application Support", "open-webui", "data", "uploads")
    xdg_data_home = os.environ.get("XDG_DATA_HOME", "").strip()
    if xdg_data_home:
        return os.path.join(xdg_data_home, "open-webui", "data", "uploads")
    return os.path.join(home, ".open-webui", "data", "uploads")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
# Pick the first existing candidate as `_data_dir`, falling back to canonical.
_data_dir = ""
for _cand in _candidate_owui_data_dirs():
    if _cand and os.path.isdir(_cand):
        _data_dir = _cand
        break
if not _data_dir:
    # None exist yet — use canonical default so _DB_PATH is well-formed
    _data_dir = "/app/backend/data" if os.name != "nt" else _candidate_owui_data_dirs()[-1]

_DB_PATH = os.path.join(_data_dir, "webui.db")
_UPLOAD_DIR = os.path.join(_data_dir, "uploads")

_EXPORT_DIR = os.environ.get("OWUI_EXPORTS_DIR", os.path.join(_data_dir, "exports"))

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
    "_candidate_owui_data_dirs", "_get_owui_data_dir", "_get_owui_uploads_dir",
    "_data_dir", "_DB_PATH", "_UPLOAD_DIR", "_EXPORT_DIR",
    "_P_NS", "_P14_NS", "_R_NS", "_PKG_REL_NS", "_CT_NS",
    "_CM_REL_TYPE", "_CT_MODERN", "_CT_AUTHORS",
]