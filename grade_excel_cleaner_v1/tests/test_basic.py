from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from grade_excel_cleaner.excel_reader import SheetData, WorkbookData
from grade_excel_cleaner.executor import execute_plan
from grade_excel_cleaner.schemas import ColumnMapping, ExtractionPlan, MetadataItem, MetadataMapping
from grade_excel_cleaner.utils import extract_json_object
from grade_excel_cleaner.validator import validate_output


def test_extract_json_from_markdown_block():
    payload = extract_json_object('```json\n{"sheet_name":"Sheet1"}\n```')
    assert payload["sheet_name"] == "Sheet1"


def test_execute_plan_with_metadata_course_name():
    raw = pd.DataFrame(
        [
            ["标题", None, None],
            ["学号", "姓名", "总成绩"],
            ["20230001", "张三", 91],
            ["20230002.0", "李四", "良好"],
        ]
    )
    workbook = WorkbookData(
        filename="绩效管理.xlsx",
        sheets={"Sheet1": SheetData(name="Sheet1", frame=raw)},
        engine_used="mock",
        read_warnings=[],
    )
    plan = ExtractionPlan(
        sheet_name="Sheet1",
        table_type_guess="single_course",
        header_row_index=1,
        data_start_row_index=2,
        data_end_strategy="all_rows_after_header",
        column_mapping=ColumnMapping(
            student_id="学号",
            student_name="姓名",
            course_name=None,
            class_name=None,
            final_score="总成绩",
        ),
        metadata_mapping=MetadataMapping(
            course_name=MetadataItem(source="filename", value="绩效管理", reason="来自文件名"),
            class_name=MetadataItem(source="unknown", value="", reason="未发现"),
        ),
        final_score_reason="总成绩优先级最高",
        confidence=0.9,
    )

    result = execute_plan(workbook, plan)
    validate_output(result.output)

    assert list(result.output.columns) == ["学号", "学生姓名", "课程名", "班级名", "最终成绩"]
    assert result.output.loc[0, "课程名"] == "绩效管理"
    assert result.output.loc[1, "学号"] == "20230002"
    assert result.table_type_detected == "single_course_like"


if __name__ == "__main__":
    test_extract_json_from_markdown_block()
    test_execute_plan_with_metadata_course_name()
    print("basic tests passed")
