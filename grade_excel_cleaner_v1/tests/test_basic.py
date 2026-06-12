from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from grade_excel_cleaner.excel_reader import SheetData, WorkbookData
from grade_excel_cleaner.executor import execute_plan
from grade_excel_cleaner.schemas import ColumnMapping, ExtractionPlan, MetadataItem, MetadataMapping
from grade_excel_cleaner.target_workflow import execute_target_item_workflow, execute_target_workflow
from grade_excel_cleaner.utils import extract_json_object
from grade_excel_cleaner.validator import validate_output
from app import _build_audit_issues


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


def test_execute_target_workflow_direct_target_columns():
    raw = pd.DataFrame(
        [
            ["学号", "姓名", "课程目标1", "课程目标2", "卷面成绩"],
            ["20230001", "张三", 0.8, 0.9, 85],
            ["20230002.0", "李四", 1, 0.75, 88],
        ]
    )
    workbook = WorkbookData(
        filename="目标成绩.xlsx",
        sheets={"综合测评": SheetData(name="综合测评", frame=raw)},
        engine_used="mock",
        read_warnings=[],
    )

    result = execute_target_workflow(workbook)

    assert list(result.output.columns) == ["学号", "姓名", "课程目标1", "课程目标2", "总分"]
    assert result.output.loc[0, "课程目标1"] == 80
    assert result.output.loc[1, "课程目标2"] == 75
    assert result.output.loc[0, "总分"] == 85


def test_execute_target_workflow_multirow_grouped_targets():
    raw = pd.DataFrame(
        [
            ["序号", "学号", "学生姓名", "毕业要求1", None, "毕业要求2", None, "总评成绩"],
            [None, None, None, "课程目标1", None, "课程目标2", None, None],
            [None, None, None, "学习表现", "期末考试", "作业", "期末考试", None],
            [None, None, None, 5, 15, 10, 20, 100],
            [1, "20230001", "张三", 4, 12, 8, 18, 84],
            [2, "20230002", "李四", 5, 15, 10, 20, 100],
        ]
    )
    workbook = WorkbookData(
        filename="多行目标.xlsx",
        sheets={"Sheet1": SheetData(name="Sheet1", frame=raw)},
        engine_used="mock",
        read_warnings=[],
    )

    result = execute_target_workflow(workbook)

    assert list(result.output.columns) == ["学号", "姓名", "课程目标1", "课程目标2", "总分"]
    assert result.output.loc[0, "课程目标1"] == 80
    assert result.output.loc[0, "课程目标2"] == 86.67
    assert result.output.loc[1, "总分"] == 100


def test_execute_target_item_workflow_keeps_assessment_items():
    raw = pd.DataFrame(
        [
            ["序号", "学号", "学生姓名", "毕业要求1", None, "毕业要求2", None, "总评成绩"],
            [None, None, None, "课程目标1", None, "课程目标2", None, None],
            [None, None, None, "学习表现", "期末考试", "作业", "期末考试", None],
            [None, None, None, 5, 15, 10, 20, 100],
            [1, "20230001", "张三", 4, 12, 8, 18, 84],
            [2, "20230002", "李四", 5, 15, 10, 20, 100],
        ]
    )
    workbook = WorkbookData(
        filename="多行目标含考核项.xlsx",
        sheets={"Sheet1": SheetData(name="Sheet1", frame=raw)},
        engine_used="mock",
        read_warnings=[],
    )

    result = execute_target_item_workflow(workbook)

    assert list(result.output.columns) == [
        "学号",
        "姓名",
        "课程目标1-学习表现",
        "课程目标1-期末考试",
        "课程目标1",
        "课程目标2-作业",
        "课程目标2-期末考试",
        "课程目标2",
        "总分",
    ]
    assert result.output.loc[0, "课程目标1-学习表现"] == 4
    assert result.output.loc[0, "课程目标1-期末考试"] == 12
    assert result.output.loc[0, "课程目标1"] == 80
    assert result.output.loc[0, "课程目标2-作业"] == 8
    assert result.output.loc[0, "课程目标2-期末考试"] == 18
    assert result.output.loc[0, "课程目标2"] == 86.67


def test_audit_does_not_count_blank_student_ids_as_duplicates():
    output = pd.DataFrame(
        {
            "学号": ["", "", "20230001", "20230002"],
            "学生姓名": ["张三", "李四", "王五", "赵六"],
            "最终成绩": [80, 81, 82, 83],
        }
    )

    issues = _build_audit_issues(output, [])

    issue_types = [issue["类型"] for issue in issues]
    assert "学号 空值" in issue_types
    assert "学号重复" not in issue_types


if __name__ == "__main__":
    test_extract_json_from_markdown_block()
    test_execute_plan_with_metadata_course_name()
    test_execute_target_workflow_direct_target_columns()
    test_execute_target_workflow_multirow_grouped_targets()
    test_execute_target_item_workflow_keeps_assessment_items()
    test_audit_does_not_count_blank_student_ids_as_duplicates()
    print("basic tests passed")
