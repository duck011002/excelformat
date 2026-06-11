from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

from ..excel_reader import WorkbookData
from ..utils import clean_cell
from .schemas import EvidenceItem, WorkbookEvidence


STRONG_KEYWORDS = (
    "教学班",
    "课程班",
    "班级",
    "运行学年",
    "学年",
    "学期",
    "任课",
    "教师",
    "老师",
    "工号",
    "课程名称",
    "课程名",
)
EXPANDED_KEYWORDS = STRONG_KEYWORDS + (
    "班",
    "专业",
    "年级",
    "学号",
    "姓名",
    "秋",
    "春",
)
TERM_PATTERN = re.compile(
    r"(?:20\d{2}\s*[-/—]\s*20\d{2}(?:学年)?|"
    r"\d{2}\s*[-/—]\s*\d{2}\s*[-/—]\s*[123]|20\d{2}\s*[春秋])"
)


def scan_workbook(
    workbook: WorkbookData,
    *,
    expanded: bool = False,
    max_items: int = 80,
) -> WorkbookEvidence:
    collector = _EvidenceCollector(max_items=max(1, max_items))
    collector.add(EvidenceItem(workbook.filename, "filename", score=1.0))

    for sheet_name, sheet in workbook.sheets.items():
        collector.add(EvidenceItem(sheet_name, "sheet_name", sheet_name=sheet_name, score=0.95))
        frame = sheet.frame
        non_empty_rows = _non_empty_row_indexes(frame)
        boundary_rows = non_empty_rows[:12] + non_empty_rows[-8:]
        for row_index in dict.fromkeys(boundary_rows):
            text = _row_context(frame, row_index)
            if text:
                collector.add(
                    EvidenceItem(
                        text,
                        "top_rows" if row_index in non_empty_rows[:12] else "bottom_rows",
                        sheet_name=sheet_name,
                        row_index=row_index,
                        score=0.85,
                    )
                )

        keywords = EXPANDED_KEYWORDS if expanded else STRONG_KEYWORDS
        for row_index, column_index, value in _sparse_matches(frame, keywords, expanded):
            text = _cell_neighborhood(frame, row_index, column_index, expanded=expanded)
            collector.add(
                EvidenceItem(
                    text,
                    "keyword_neighborhood",
                    sheet_name=sheet_name,
                    row_index=row_index,
                    column_index=column_index,
                    score=0.92 if not expanded else 0.82,
                )
            )

    return WorkbookEvidence(filename=workbook.filename, items=collector.items, expanded=expanded)


class _EvidenceCollector:
    def __init__(self, *, max_items: int) -> None:
        self.max_items = max_items
        self.items: list[EvidenceItem] = []
        self._seen: set[tuple[str, str]] = set()

    def add(self, item: EvidenceItem) -> None:
        text = re.sub(r"\s+", " ", clean_cell(item.text))[:500]
        key = (item.source_type, text)
        if not text or key in self._seen:
            return
        item.text = text
        if len(self.items) >= self.max_items:
            lowest_index = min(range(len(self.items)), key=lambda index: self.items[index].score)
            if item.score <= self.items[lowest_index].score:
                return
            removed = self.items.pop(lowest_index)
            self._seen.discard((removed.source_type, removed.text))
        self.items.append(item)
        self._seen.add(key)


def _non_empty_row_indexes(frame: pd.DataFrame) -> list[int]:
    indexes: list[int] = []
    for row_index in range(len(frame)):
        if any(clean_cell(value) for value in frame.iloc[row_index].tolist()):
            indexes.append(row_index)
    return indexes


def _sparse_matches(
    frame: pd.DataFrame,
    keywords: tuple[str, ...],
    expanded: bool,
) -> Iterable[tuple[int, int, str]]:
    for row_index in range(len(frame)):
        for column_index in range(frame.shape[1]):
            value = clean_cell(frame.iat[row_index, column_index])
            if not value:
                continue
            if any(keyword in value for keyword in keywords) or TERM_PATTERN.search(value):
                yield row_index, column_index, value
            elif expanded and _looks_like_class_token(value):
                yield row_index, column_index, value


def _row_context(frame: pd.DataFrame, row_index: int) -> str:
    values = [clean_cell(value) for value in frame.iloc[row_index].tolist()]
    return " | ".join(value for value in values if value)


def _cell_neighborhood(
    frame: pd.DataFrame,
    row_index: int,
    column_index: int,
    *,
    expanded: bool,
) -> str:
    column_radius = 2 if expanded else 1
    row_radius = 1 if expanded else 0
    parts: list[str] = []
    for nearby_row in range(max(0, row_index - row_radius), min(len(frame), row_index + row_radius + 1)):
        row_values: list[str] = []
        for nearby_column in range(
            max(0, column_index - column_radius),
            min(frame.shape[1], column_index + column_radius + 1),
        ):
            value = clean_cell(frame.iat[nearby_row, nearby_column])
            if value:
                row_values.append(value)
        if row_values:
            parts.append(" | ".join(row_values))
    return " || ".join(parts)


def _looks_like_class_token(value: str) -> bool:
    compact = re.sub(r"\s+", "", value)
    return bool(
        re.search(r"(?:\d{2,4}级?.{0,10}班|.{0,8}\d{1,3}班)$", compact)
        and len(compact) <= 40
    )
