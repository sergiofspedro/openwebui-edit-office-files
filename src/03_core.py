
class Tools:
    class Valves(BaseModel):
        base_url: Optional[str] = Field(
            default=None,
            description="Override the base URL for download links. Auto-detected from X-Original-Host header or WEBUI_URL env var if unset.",
        )
        templates: Optional[str] = Field(default="{}", description="JSON map of template names to content strings.")
        cleanup_schedule: Optional[str] = Field(default="{}", description="JSON schedule for auto-cleanup.")
        language: Optional[str] = Field(default="en", description="Language for error messages: en, pt, es, fr, de.")
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
    async def _progress(self, current: int, total: int, operation: str = "Processing") -> None:
        """Emit progress via __event_emitter__ if available."""
        try:
            await __event_emitter__({"type": "status", "data": {"description": f"{operation}: {current}/{total}", "done": current >= total}})
        except Exception:
            pass


    # --- v3.3.0: Document Comparison ---
