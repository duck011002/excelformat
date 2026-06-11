from __future__ import annotations

import os
import sys
import unittest
from io import BytesIO
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from grade_excel_cleaner.excel_reader import read_workbook
from grade_excel_cleaner.teaching_class import (
    TEACHING_CLASS_COLUMNS,
    build_teaching_class_result,
    teaching_class_to_xlsx,
)
from grade_excel_cleaner.utils import clean_cell


EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsb"}


def _database_root() -> Path | None:
    configured = os.getenv("GRADE_CLEANER_DATABASE")
    if configured:
        path = Path(configured)
        return path if path.exists() else None
    candidate = ROOT.parent / "database"
    return candidate if candidate.exists() else None


def _template_path() -> Path | None:
    configured = os.getenv("GRADE_CLEANER_TEACHING_TEMPLATE")
    if not configured:
        return None
    path = Path(configured)
    return path if path.exists() else None


def _excel_files(folder: Path, *, recursive: bool = False) -> list[Path]:
    iterator = folder.rglob("*") if recursive else folder.iterdir()
    return sorted(
        path
        for path in iterator
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    )


class RealSampleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.database = _database_root()

    def _assert_readable_and_exportable(self, path: Path) -> None:
        workbook = read_workbook(path)
        result = build_teaching_class_result(workbook, pd.DataFrame())
        payload = teaching_class_to_xlsx(result.groups)
        exported = pd.read_excel(BytesIO(payload), dtype=object)

        self.assertTrue(workbook.sheets, path.name)
        self.assertEqual(list(exported.columns), TEACHING_CLASS_COLUMNS, path.name)
        self.assertGreaterEqual(len(exported), 1, path.name)

    def test_all_root_database_workbooks(self) -> None:
        if self.database is None:
            self.skipTest("database directory is not available")
        samples = _excel_files(self.database)
        self.assertEqual(len(samples), 6)
        for path in samples:
            with self.subTest(path=path.name):
                self._assert_readable_and_exportable(path)

    def test_one_workbook_per_score_subdirectory(self) -> None:
        if self.database is None:
            self.skipTest("database directory is not available")
        score_root = self.database / "成绩"
        if not score_root.exists():
            self.skipTest("database/成绩 is not available")

        samples: list[Path] = []
        for folder in sorted(path for path in score_root.iterdir() if path.is_dir()):
            choices = _excel_files(folder, recursive=True)
            self.assertTrue(choices, folder.name)
            samples.append(choices[0])

        for path in samples:
            with self.subTest(path=str(path.relative_to(score_root))):
                self._assert_readable_and_exportable(path)

    def test_template_headers_match_export_contract(self) -> None:
        template = _template_path()
        if template is None:
            self.skipTest("teaching-class template is not configured")
        workbook = read_workbook(template)
        matched = False
        for sheet in workbook.sheets.values():
            for row_index in range(len(sheet.frame)):
                values = [
                    clean_cell(value).replace("*", "").replace("\n", "").strip()
                    for value in sheet.frame.iloc[row_index].tolist()
                ]
                if len(values) >= len(TEACHING_CLASS_COLUMNS) and all(
                    actual.startswith(expected)
                    for actual, expected in zip(
                        values[: len(TEACHING_CLASS_COLUMNS)],
                        TEACHING_CLASS_COLUMNS,
                    )
                ):
                    matched = True
                    break
        self.assertTrue(matched)


if __name__ == "__main__":
    unittest.main()
