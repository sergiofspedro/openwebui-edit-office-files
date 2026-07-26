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

    async def schedule_cleanup(self, days_old: int = 30, interval_hours: int = 24) -> str:
        """Schedule automatic cleanup every N hours. Set interval_hours=0 to disable."""
        schedule = {"days_old": days_old, "interval_hours": interval_hours, "enabled": interval_hours > 0}
        self.valves.cleanup_schedule = json.dumps(schedule)
        if interval_hours > 0:
            return f"Cleanup scheduled: remove files older than {days_old} days, every {interval_hours} hours."
        return "Scheduled cleanup disabled."

    # --- v3.2.0: Mail Merge ---
