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

