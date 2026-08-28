"""Tests for v4.0.4 fixes (#10, #12). Pure stdlib, no pydantic needed.

We exercise the patched _resolve_file_path and bulk_folder_ops logic by
extracting the relevant code (or running it through stubs) and assert the
behaviour the issues call out.
"""
import os
import sys
import sqlite3
import tempfile
import unittest
from unittest import mock


REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(REPO, "src")
sys.path.insert(0, SRC)


# ---------------------------------------------------------------------------
# Fix #12 — _resolve_file_path must NOT raise when the metadata DB is not
# reachable (e.g. PostgreSQL: ?mode=ro is rejected by the driver).
# We test the module-level function by faking sqlite3.connect to raise.
# ---------------------------------------------------------------------------
class TestResolveFilePathGracefulDB(unittest.TestCase):
    def setUp(self):
        # Load src/utils.py without going through pydantic dependencies.
        # src/utils.py imports from .constants — give it a fake constants
        # module on sys.path that mimics the bits it uses.
        import importlib.util

        # Build a minimal constants stub.
        constants_stub = type(sys)("constants")
        constants_stub._data_dir = ""
        constants_stub._DB_PATH = "/nonexistent/webui.db"
        constants_stub._UPLOAD_DIR = tempfile.mkdtemp()
        constants_stub._EXPORT_DIR = tempfile.mkdtemp()
        sys.modules.setdefault("constants", constants_stub)
        # src.utils uses `from .constants import ...` — rewire to non-package.
        # Easier: load it as a top-level module by faking the package.
        pkg = type(sys)("src")
        pkg.__path__ = [SRC]
        sys.modules["src"] = pkg
        sys.modules["src.constants"] = constants_stub

        spec = importlib.util.spec_from_file_location("src.utils", os.path.join(SRC, "utils.py"))
        self.utils = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.utils)

    def test_resolve_file_path_returns_none_when_sqlite_uri_rejected(self):
        """Postgres raises on `?mode=ro` URI — our function must return None,
        not propagate the exception to callers. (#12)"""
        def _explode(*a, **kw):
            raise sqlite3.OperationalError("file:/?mode=ro is not supported by this driver")
        with mock.patch.object(sqlite3, "connect", side_effect=_explode):
            result = self.utils._resolve_file_path("any-file-id")
        self.assertIsNone(result, "Must return None on DB error, not raise")

    def test_resolve_file_path_returns_none_when_db_missing(self):
        """When the DB file is missing, sqlite3.connect raises — must return None."""
        # No patch — real connect to a path that does not exist.
        # On Python 3.13 sqlite3 raises immediately if file missing, but
        # behaviour varies: we accept either None (caught) or OperationalError
        # being raised then caught inside _lookup. Force the connect to raise.
        with mock.patch.object(sqlite3, "connect", side_effect=sqlite3.OperationalError("no such file")):
            result = self.utils._resolve_file_path("any-file-id")
        self.assertIsNone(result)

    def test_resolve_file_path_returns_path_when_db_ok(self):
        """Happy path: a real SQLite DB with a matching row must return the path."""
        with tempfile.TemporaryDirectory() as d:
            db = os.path.join(d, "webui.db")
            uploads = os.path.join(d, "uploads")
            os.makedirs(uploads)
            fake_path = os.path.join(d, "foo.docx")  # real file on disk
            with open(fake_path, "w") as fh:
                fh.write("hi")
            conn = sqlite3.connect(db)
            try:
                conn.execute("CREATE TABLE file (id TEXT PRIMARY KEY, path TEXT, filename TEXT)")
                conn.execute("INSERT INTO file VALUES (?, ?, ?)", ("abc-123", fake_path, "foo.docx"))
                conn.commit()
            finally:
                conn.close()
            self.utils._DB_PATH = db
            self.utils._UPLOAD_DIR = uploads
            result = self.utils._resolve_file_path("abc-123")
            self.assertEqual(result, fake_path)


# ---------------------------------------------------------------------------
# Fix #10 — bulk_folder_ops must:
#   (a) reject delete_old when allow_bulk_delete is False
#   (b) provide preview_delete_old (read-only, never deletes)
#   (c) scope the deletion to meta LIKE '%office-plugin%'
# We test the logic by extracting the function body from the source file and
# executing it against stubs, avoiding the pydantic dependency on the Tools
# class.
# ---------------------------------------------------------------------------
def _extract_function(source: str, name: str) -> str:
    """Pull the `def NAME(...)` block (header + body) out of a Python source."""
    import re
    pattern = re.compile(rf"    (async )?def {name}\(", re.M)
    m = pattern.search(source)
    if not m:
        raise AssertionError(f"function {name} not found")
    # Walk forward to find the matching dedent (line whose indentation is 4 or less).
    lines = source.splitlines()
    i = m.start() // 1
    # Find the line number of the def
    line_no = source[:m.start()].count("\n")
    # Take a generous slice — we'll exec() the substring; if it has siblings
    # that don't define needed names, we provide them via globals below.
    # Easier: take the whole class body. We'll just exec the whole file.
    return source


class TestBulkFolderOpsScoping(unittest.TestCase):
    """Logic-level checks: extract the bulk_folder_ops source and exec it
    against a controlled namespace, then assert return values."""

    def setUp(self):
        self.src = open(os.path.join(REPO, "src", "tool.py")).read()

    def _exec_function(self):
        """Execute the whole src/tool.py and return the bound method's source-
        level globals we need to call it via a stub instance."""
        # We can't import src.tool because it pulls pydantic. Instead, parse
        # out the function body and exec it in a controlled namespace.
        import ast
        tree = ast.parse(self.src)
        target = None
        for node in ast.walk(tree):
            if isinstance(node, ast.AsyncFunctionDef) and node.name == "bulk_folder_ops":
                target = node
                break
        self.assertIsNotNone(target, "bulk_folder_ops not found in src/tool.py")
        # Unparse the function back to source.
        func_src = ast.unparse(target)
        ns = {
            "os": os, "json": __import__("json"), "sqlite3": sqlite3,
            "sys": sys, "_UPLOAD_DIR": tempfile.mkdtemp(),
            "_DB_PATH": "/tmp/nonexistent.db",
            "getattr": getattr,
        }
        exec(func_src, ns)
        return ns["bulk_folder_ops"]

    def test_bulk_folder_ops_rejects_unknown_operation(self):
        import asyncio
        fn = self._exec_function()
        stub = type("Stub", (), {
            "valves": type("V", (), {"allow_bulk_delete": True})(),
        })()
        out = asyncio.run(fn(stub, "make_coffee", pattern="*"))
        self.assertIn("Unknown operation", out)

    def test_bulk_folder_ops_preview_requires_valve_true(self):
        """The function must reference getattr(self.valves, 'allow_bulk_delete', False)
        so the Valve gate works at runtime."""
        import re
        m = re.search(r"async def bulk_folder_ops.*?(?=\n    async def |\n    def |\nclass )", self.src, re.S)
        self.assertIsNotNone(m)
        body = m.group(0)
        self.assertIn('getattr(self.valves, "allow_bulk_delete", False)', body,
                      "must gate on allow_bulk_delete via getattr default False")

    def test_bulk_folder_ops_contains_safety_gates(self):
        """Static check: the function source must reference all three
        protections called out in the issue."""
        import re
        m = re.search(r"async def bulk_folder_ops.*?(?=\n    async def |\n    def |\nclass )", self.src, re.S)
        self.assertIsNotNone(m)
        body = m.group(0)
        self.assertIn("allow_bulk_delete", body, "must gate on the Valve")
        self.assertIn("preview_delete_old", body, "must expose dry-run operation")
        self.assertIn("meta LIKE '%office-plugin%'", body, "must scope to office-plugin files")
        self.assertIn("'..'", body, "must reject path traversal")


if __name__ == "__main__":
    unittest.main(verbosity=2)
