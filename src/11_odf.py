    async def create_odf(self, content: str, filename: str = "document", format: str = "odt", __user__=None, __request__=None) -> str:
        """Create a new ODF file (.odt, .ods, .odp)."""
        from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet, OpenDocumentPresentation
        from odf.text import P, H
        from odf.table import Table, TableRow, TableCell
        import io
        
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
                        cell.addElement(P(text=_format_text(c)))
                        row.addElement(cell)
                    table.addElement(row)
                doc.spreadsheet.addElement(table)
            elif format == "odp":
                doc = OpenDocumentPresentation()
                for line in content.split('\n'):
                    line = line.strip()
                    if not line: continue
                    if line.startswith('# '):
                        doc.presentation.addElement(H(outlinelevel=1, text=_format_text(line[2:])))
                    else:
                        doc.presentation.addElement(P(text=_format_text(line)))
            else:
                doc = OpenDocumentText()
                for line in content.split('\n'):
                    line = line.strip()
                    if not line: continue
                    if line.startswith('# '):
                        doc.text.addElement(H(outlinelevel=1, text=_format_text(line[2:])))
                    elif line.startswith('## '):
                        doc.text.addElement(H(outlinelevel=2, text=_format_text(line[3:])))
                    else:
                        doc.text.addElement(P(text=_format_text(line)))
            
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
