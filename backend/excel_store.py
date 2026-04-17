"""Excel 结构化存储（Sheet / Row）"""
from __future__ import annotations

from datetime import datetime
from typing import Any

from cache import cache
from database import SessionLocal
from models import ExcelRow, ExcelSheet


class ExcelKnowledgeStore:
    @staticmethod
    def _row_cache_key(chunk_id: str) -> str:
        return f"excel_row:{chunk_id}"

    @staticmethod
    def _sheet_cache_key(chunk_id: str) -> str:
        return f"excel_sheet:{chunk_id}"

    @staticmethod
    def _row_to_dict(item: ExcelRow) -> dict[str, Any]:
        return {
            "chunk_id": item.chunk_id,
            "filename": item.filename,
            "file_path": item.file_path,
            "sheet_name": item.sheet_name,
            "sheet_index": item.sheet_index,
            "row_index": item.row_index,
            "row_text": item.row_text,
            "row_obj": item.row_obj,
            "headers": item.headers,
            "sheet_chunk_id": item.sheet_chunk_id,
        }

    @staticmethod
    def _sheet_to_dict(item: ExcelSheet) -> dict[str, Any]:
        return {
            "chunk_id": item.chunk_id,
            "filename": item.filename,
            "file_path": item.file_path,
            "sheet_name": item.sheet_name,
            "sheet_index": item.sheet_index,
            "headers": item.headers,
            "sheet_text": item.sheet_text,
            "sheet_html": item.sheet_html,
            "row_count": item.row_count,
            "column_count": item.column_count,
        }

    def replace_workbook(self, workbook: dict[str, Any]) -> dict[str, int]:
        filename = (workbook.get("filename") or "").strip()
        if not filename:
            return {"sheets": 0, "rows": 0}

        sheet_records = workbook.get("sheets", []) or []
        row_records = workbook.get("rows", []) or []

        db = SessionLocal()
        try:
            db.query(ExcelRow).filter(ExcelRow.filename == filename).delete(synchronize_session=False)
            db.query(ExcelSheet).filter(ExcelSheet.filename == filename).delete(synchronize_session=False)

            now = datetime.utcnow()
            for sheet in sheet_records:
                db.add(
                    ExcelSheet(
                        chunk_id=sheet["chunk_id"],
                        filename=filename,
                        file_path=sheet.get("file_path", ""),
                        sheet_name=sheet.get("sheet_name", ""),
                        sheet_index=int(sheet.get("sheet_index", 0) or 0),
                        headers=sheet.get("headers", []),
                        sheet_text=sheet.get("sheet_text", ""),
                        sheet_html=sheet.get("sheet_html", ""),
                        row_count=int(sheet.get("row_count", 0) or 0),
                        column_count=int(sheet.get("column_count", 0) or 0),
                        updated_at=now,
                    )
                )
                cache.set_json(self._sheet_cache_key(sheet["chunk_id"]), sheet)

            for row in row_records:
                db.add(
                    ExcelRow(
                        chunk_id=row["chunk_id"],
                        filename=filename,
                        file_path=row.get("file_path", ""),
                        sheet_name=row.get("sheet_name", ""),
                        sheet_index=int(row.get("sheet_index", 0) or 0),
                        row_index=int(row.get("row_index", 0) or 0),
                        row_text=row.get("row_text", ""),
                        row_obj=row.get("row_obj", {}),
                        headers=row.get("headers", []),
                        sheet_chunk_id=row.get("sheet_chunk_id", ""),
                        updated_at=now,
                    )
                )
                cache.set_json(self._row_cache_key(row["chunk_id"]), row)

            db.commit()
        finally:
            db.close()

        return {"sheets": len(sheet_records), "rows": len(row_records)}

    def get_rows_by_chunk_ids(self, chunk_ids: list[str]) -> list[dict[str, Any]]:
        if not chunk_ids:
            return []

        ordered_results: dict[str, dict[str, Any]] = {}
        missing_ids: list[str] = []
        for chunk_id in chunk_ids:
            key = (chunk_id or "").strip()
            if not key:
                continue
            cached = cache.get_json(self._row_cache_key(key))
            if cached:
                ordered_results[key] = cached
            else:
                missing_ids.append(key)

        if missing_ids:
            db = SessionLocal()
            try:
                rows = db.query(ExcelRow).filter(ExcelRow.chunk_id.in_(missing_ids)).all()
                for row in rows:
                    payload = self._row_to_dict(row)
                    ordered_results[row.chunk_id] = payload
                    cache.set_json(self._row_cache_key(row.chunk_id), payload)
            finally:
                db.close()

        return [ordered_results[item] for item in chunk_ids if item in ordered_results]

    def get_sheets_by_chunk_ids(self, chunk_ids: list[str]) -> list[dict[str, Any]]:
        if not chunk_ids:
            return []

        ordered_results: dict[str, dict[str, Any]] = {}
        missing_ids: list[str] = []
        for chunk_id in chunk_ids:
            key = (chunk_id or "").strip()
            if not key:
                continue
            cached = cache.get_json(self._sheet_cache_key(key))
            if cached:
                ordered_results[key] = cached
            else:
                missing_ids.append(key)

        if missing_ids:
            db = SessionLocal()
            try:
                rows = db.query(ExcelSheet).filter(ExcelSheet.chunk_id.in_(missing_ids)).all()
                for row in rows:
                    payload = self._sheet_to_dict(row)
                    ordered_results[row.chunk_id] = payload
                    cache.set_json(self._sheet_cache_key(row.chunk_id), payload)
            finally:
                db.close()

        return [ordered_results[item] for item in chunk_ids if item in ordered_results]

    def delete_by_filename(self, filename: str) -> dict[str, int]:
        if not filename:
            return {"sheets": 0, "rows": 0}

        db = SessionLocal()
        try:
            row_items = db.query(ExcelRow).filter(ExcelRow.filename == filename).all()
            sheet_items = db.query(ExcelSheet).filter(ExcelSheet.filename == filename).all()
            row_ids = [item.chunk_id for item in row_items]
            sheet_ids = [item.chunk_id for item in sheet_items]

            if row_ids:
                db.query(ExcelRow).filter(ExcelRow.filename == filename).delete(synchronize_session=False)
            if sheet_ids:
                db.query(ExcelSheet).filter(ExcelSheet.filename == filename).delete(synchronize_session=False)
            db.commit()

            for chunk_id in row_ids:
                cache.delete(self._row_cache_key(chunk_id))
            for chunk_id in sheet_ids:
                cache.delete(self._sheet_cache_key(chunk_id))

            return {"sheets": len(sheet_ids), "rows": len(row_ids)}
        finally:
            db.close()
