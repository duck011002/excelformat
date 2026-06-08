from __future__ import annotations

import json
from typing import Any

import pandas as pd

from .excel_reader import WorkbookData
from .utils import clean_cell


HEADER_KEYWORDS = [
    "学号",
    "姓名",
    "学生",
    "课程",
    "班级",
    "成绩",
    "总评",
    "总成绩",
    "最终",
]


def _row_values(frame: pd.DataFrame, row_index: int) -> list[str]:
    if row_index >= len(frame):
        return []
    return [clean_cell(value) for value in frame.iloc[row_index].tolist()]


def _non_empty_count(values: list[str]) -> int:
    return sum(1 for value in values if value)


def _guess_header_rows(frame: pd.DataFrame, limit: int) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for row_index in range(min(limit, len(frame))):
        values = _row_values(frame, row_index)
        joined = " ".join(values)
        keyword_hits = sum(1 for keyword in HEADER_KEYWORDS if keyword in joined)
        non_empty = _non_empty_count(values)
        score = keyword_hits * 3 + min(non_empty, 8)
        if keyword_hits or non_empty >= 3:
            candidates.append(
                {
                    "row_index": row_index,
                    "score": score,
                    "non_empty_count": non_empty,
                    "values": [value for value in values if value][:30],
                }
            )
    return sorted(candidates, key=lambda item: item["score"], reverse=True)[:5]


def build_workbook_preview(workbook: WorkbookData, preview_rows: int = 25) -> dict[str, Any]:
    sheets: list[dict[str, Any]] = []
    for sheet in workbook.sheets.values():
        frame = sheet.frame
        rows: list[dict[str, Any]] = []
        for row_index in range(min(preview_rows, len(frame))):
            values = _row_values(frame, row_index)
            rows.append(
                {
                    "row_index": row_index,
                    "non_empty_count": _non_empty_count(values),
                    "values": values,
                }
            )
        candidates = _guess_header_rows(frame, preview_rows)
        sheets.append(
            {
                "sheet_name": sheet.name,
                "row_count": int(frame.shape[0]),
                "column_count": int(frame.shape[1]),
                "top_rows": rows,
                "candidate_header_rows": candidates,
                "candidate_column_names": candidates[0]["values"] if candidates else [],
            }
        )
    return {
        "filename": workbook.filename,
        "engine_used": workbook.engine_used,
        "read_warnings": workbook.read_warnings,
        "filename_hints": _filename_hints(workbook.filename),
        "sheets": sheets,
    }


def preview_to_json(preview: dict[str, Any]) -> str:
    return json.dumps(preview, ensure_ascii=False, indent=2)


def _filename_hints(filename: str) -> dict[str, str]:
    stem = filename.rsplit(".", 1)[0]
    return {
        "raw": stem,
        "course_or_class_hint": stem,
        "note": "文件名可能包含课程名、班级名、学期或教师信息，需要结合 sheet 和前几行判断。",
    }
