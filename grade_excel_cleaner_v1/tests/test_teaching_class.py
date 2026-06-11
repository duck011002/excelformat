from __future__ import annotations

import json
import sys
import unittest
from io import BytesIO
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from grade_excel_cleaner.excel_reader import SheetData, WorkbookData
from grade_excel_cleaner.teaching_class import (
    TEACHING_CLASS_COLUMNS,
    build_llm_context,
    build_teaching_class_result,
    flatten_groups,
    normalize_academic_term,
    resolve_uncertain_groups,
    scan_workbook,
    teaching_class_to_xlsx,
)
from app import _apply_teaching_group_edits


def workbook_from_rows(
    rows: list[list[object]],
    *,
    filename: str = "绩效管理-2023秋.xlsx",
    sheet_name: str = "成绩",
) -> WorkbookData:
    return WorkbookData(
        filename=filename,
        sheets={sheet_name: SheetData(name=sheet_name, frame=pd.DataFrame(rows))},
        engine_used="mock",
        read_warnings=[],
    )


class AcademicTermTests(unittest.TestCase):
    def test_normalizes_autumn_spring_and_short_forms(self) -> None:
        autumn = normalize_academic_term("2023秋")
        spring = normalize_academic_term("2024春")
        short = normalize_academic_term("23-24-2")

        self.assertEqual((autumn.school_year, autumn.term), ("2023-2024学年", "第一学期"))
        self.assertEqual((spring.school_year, spring.term), ("2023-2024学年", "第二学期"))
        self.assertEqual((short.school_year, short.term), ("2023-2024学年", "第二学期"))

    def test_normalizes_explicit_school_year(self) -> None:
        result = normalize_academic_term("2022/2023学年第1学期")
        self.assertEqual(result.school_year, "2022-2023学年")
        self.assertEqual(result.term, "第一学期")

    def test_unrelated_text_stays_blank(self) -> None:
        result = normalize_academic_term("课程成绩汇总")
        self.assertEqual((result.school_year, result.term), ("", ""))


class ScannerTests(unittest.TestCase):
    def test_sparse_scan_finds_filename_sheet_and_far_keyword_neighbors(self) -> None:
        rows = [[""] for _ in range(80)]
        rows[0] = ["课程名称", "绩效管理", ""]
        rows[37] = ["备注", "教学班：人力22-1班", "任课教师：张三"]
        rows[79] = ["运行学年", "2023-2024学年", "第2学期"]
        workbook = workbook_from_rows(rows, filename="人力22绩效管理-2023秋.xlsx", sheet_name="人力22成绩")

        evidence = scan_workbook(workbook, expanded=False)
        texts = [item.text for item in evidence.items]

        self.assertIn("人力22绩效管理-2023秋.xlsx", texts)
        self.assertIn("人力22成绩", texts)
        self.assertTrue(any("教学班" in text and "人力22-1班" in text for text in texts))
        self.assertTrue(any("2023-2024学年" in text and "第2学期" in text for text in texts))

    def test_evidence_is_bounded_and_expanded_mode_adds_more_context(self) -> None:
        rows = [[f"普通内容{i}", f"班{i}", ""] for i in range(200)]
        rows[100] = ["课程班别名", "A01", "相邻值"]
        workbook = workbook_from_rows(rows)

        compact = scan_workbook(workbook, expanded=False, max_items=20)
        expanded = scan_workbook(workbook, expanded=True, max_items=30)

        self.assertLessEqual(len(compact.items), 20)
        self.assertLessEqual(len(expanded.items), 30)
        self.assertGreaterEqual(len(expanded.items), len(compact.items))
        self.assertTrue(any("课程班别名" in item.text for item in compact.items))


class WorkflowTests(unittest.TestCase):
    def test_groups_by_course_class_alias_before_administrative_class(self) -> None:
        workbook = workbook_from_rows([["课程", "绩效管理"], ["运行学年", "2023-2024学年", "第一学期"]])
        score_output = pd.DataFrame(
            {
                "学号": ["20230001", "20230002", "20230003"],
                "学生姓名": ["张三", "李四", "王五"],
                "课程名": ["绩效管理"] * 3,
                "课程班别名": ["A班", "B班", "A班"],
                "行政班": ["人力22-1班"] * 3,
                "最终成绩": [90, 80, 70],
            }
        )

        result = build_teaching_class_result(workbook, score_output)

        self.assertEqual([group.class_name.value for group in result.groups], ["A班", "B班"])
        self.assertTrue(all(group.class_name.source_type == "course_class_alias" for group in result.groups))
        self.assertEqual(len(result.groups[0].students), 2)

    def test_duplicate_names_from_distinct_courses_keep_independent_group_ids(self) -> None:
        workbook = workbook_from_rows([["2023秋"]])
        score_output = pd.DataFrame(
            {"学号": ["1", "2"], "姓名": ["甲", "乙"], "课程名": ["课程一", "课程二"], "课程班别名": ["A班", "A班"]}
        )

        result = build_teaching_class_result(workbook, score_output)

        self.assertEqual(len(result.groups), 2)
        self.assertNotEqual(result.groups[0].group_id, result.groups[1].group_id)

    def test_same_course_and_name_with_different_codes_are_separate_groups(self) -> None:
        workbook = workbook_from_rows([["2023秋"]])
        score_output = pd.DataFrame(
            {
                "学号": ["1", "2"],
                "姓名": ["甲", "乙"],
                "课程名": ["课程一", "课程一"],
                "课程班别名": ["A班", "A班"],
                "教学班编号": ["A001", "A002"],
            }
        )

        result = build_teaching_class_result(workbook, score_output)

        self.assertEqual(len(result.groups), 2)
        self.assertEqual([group.class_code.value for group in result.groups], ["A001", "A002"])

    def test_generated_codes_are_stable_and_original_alphanumeric_codes_are_preserved(self) -> None:
        workbook = workbook_from_rows([["2023秋"]])
        score_output = pd.DataFrame(
            {"学号": ["1", "2"], "姓名": ["甲", "乙"], "课程班别名": ["A班", "B班"], "教学班编号": ["0807S102", ""]}
        )

        first = build_teaching_class_result(workbook, score_output)
        second = build_teaching_class_result(workbook, score_output)

        self.assertEqual(first.groups[0].class_code.value, "0807S102")
        self.assertEqual(first.groups[1].class_code.value, second.groups[1].class_code.value)
        self.assertTrue(first.groups[1].class_code.value.isdigit())

    def test_blank_fields_still_export_one_row_with_warnings(self) -> None:
        workbook = workbook_from_rows([["课程成绩"]], filename="未知.xlsx")
        result = build_teaching_class_result(workbook, pd.DataFrame())
        output = flatten_groups(result.groups)

        self.assertEqual(list(output.columns), TEACHING_CLASS_COLUMNS)
        self.assertEqual(len(output), 1)
        self.assertEqual(output.loc[0, "教学班学生学号"], "")
        self.assertTrue(result.warnings)

    def test_generic_header_neighbor_is_not_used_as_class_name(self) -> None:
        workbook = workbook_from_rows([["班级", "成绩"], ["2022级2班", 90]], filename="2022级2+3+6+11班-2023年.xlsb")

        result = build_teaching_class_result(workbook, pd.DataFrame())

        self.assertEqual(result.groups[0].class_name.value, "2022级2+3+6+11班")

    def test_filename_extracts_compact_class_and_optional_teacher(self) -> None:
        workbook = workbook_from_rows([["学号", "成绩"]], filename="test测控电路-20测控1~6-丁炯（22-23-2）.xlsx")

        result = build_teaching_class_result(workbook, pd.DataFrame())

        self.assertEqual(result.groups[0].class_name.value, "20测控1~6")
        self.assertEqual(result.groups[0].teacher_name.value, "丁炯")


class PlannerTests(unittest.TestCase):
    def setUp(self) -> None:
        workbook = workbook_from_rows([["教学班", "人力22-1班"], ["运行学年", "2023-2024学年", "第一学期"]])
        score_output = pd.DataFrame({"学号": ["1"], "姓名": ["甲"]})
        self.result = build_teaching_class_result(workbook, score_output)

    def test_llm_context_is_bounded(self) -> None:
        context = build_llm_context(self.result.groups, self.result.evidence, enhanced=False)
        self.assertLessEqual(len(context["evidence"]), 40)
        for group in context["groups"]:
            for values in group["candidates"].values():
                self.assertLessEqual(len(values), 5)

    def test_reanalysis_preserves_locked_fields_and_updates_unlocked_fields(self) -> None:
        group = self.result.groups[0]
        group.class_name.value = "人工班名"
        group.class_name.locked_by_user = True
        group.school_year.confidence = 0.2

        payload = {
            "groups": [
                {
                    "group_id": group.group_id,
                    "教学班名称": {"value": "AI班名", "confidence": 0.99},
                    "运行学年": {"value": "2024-2025学年", "confidence": 0.9},
                }
            ]
        }
        updated = resolve_uncertain_groups(self.result, caller=lambda **_: json.dumps(payload, ensure_ascii=False))

        self.assertEqual(updated.groups[0].class_name.value, "人工班名")
        self.assertEqual(updated.groups[0].school_year.value, "2024-2025学年")
        self.assertTrue(updated.groups[0].revisions)

    def test_invalid_json_keeps_python_result_and_adds_warning(self) -> None:
        original = self.result.groups[0].class_name.value
        updated = resolve_uncertain_groups(self.result, caller=lambda **_: "not json")

        self.assertEqual(updated.groups[0].class_name.value, original)
        self.assertTrue(any("AI" in warning for warning in updated.warnings))


class ExporterTests(unittest.TestCase):
    def test_xlsx_export_keeps_exact_columns_and_blank_cells(self) -> None:
        workbook = workbook_from_rows([["未知"]], filename="未知.xlsx")
        result = build_teaching_class_result(workbook, pd.DataFrame())

        payload = teaching_class_to_xlsx(result.groups)
        exported = pd.read_excel(BytesIO(payload), dtype=object)

        self.assertEqual(list(exported.columns), TEACHING_CLASS_COLUMNS)
        self.assertEqual(len(exported), 1)
        self.assertTrue(pd.isna(exported.loc[0, "教学班学生学号"]))

    def test_group_edits_lock_only_changed_group_fields(self) -> None:
        workbook = workbook_from_rows([["2023秋"]])
        score_output = pd.DataFrame(
            {
                "学号": ["1", "2"],
                "姓名": ["甲", "乙"],
                "课程班别名": ["A班", "B班"],
            }
        )
        result = build_teaching_class_result(workbook, score_output)
        edited = pd.DataFrame(
            [
                {
                    "_group_id": result.groups[0].group_id,
                    "教学班名称": "人工A班",
                    "运行学年": result.groups[0].school_year.value,
                },
                {
                    "_group_id": result.groups[1].group_id,
                    "教学班名称": result.groups[1].class_name.value,
                    "运行学年": result.groups[1].school_year.value,
                },
            ]
        )

        changed = _apply_teaching_group_edits(result, edited)

        self.assertTrue(changed)
        self.assertEqual(result.groups[0].class_name.value, "人工A班")
        self.assertTrue(result.groups[0].class_name.locked_by_user)
        self.assertFalse(result.groups[1].class_name.locked_by_user)

    def test_group_edit_refreshes_missing_field_warnings(self) -> None:
        result = build_teaching_class_result(
            workbook_from_rows([["未知"]], filename="未知.xlsx"),
            pd.DataFrame(),
        )
        self.assertTrue(any("未找到教学班名称" in warning for warning in result.warnings))
        edited = pd.DataFrame(
            [
                {
                    "_group_id": result.groups[0].group_id,
                    "教学班名称": "人工教学班",
                }
            ]
        )

        _apply_teaching_group_edits(result, edited)

        self.assertFalse(any("未找到教学班名称" in warning for warning in result.warnings))


if __name__ == "__main__":
    unittest.main()
