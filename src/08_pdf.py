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

