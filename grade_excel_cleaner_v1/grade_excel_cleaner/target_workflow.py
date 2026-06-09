from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from .excel_reader import WorkbookData, read_workbook
from .utils import clean_cell


ID_COLUMN = "学号"
NAME_COLUMN = "姓名"
TOTAL_COLUMN = "总分"
TARGET_PREFIX = "课程目标"

ID_HEADERS = {"学号", "学生学号"}
NAME_HEADERS = {"姓名", "学生姓名"}
IGNORED_HEADERS = {"序号", "班级", "性别", "专业", "学院"}
TOTAL_KEYWORDS = ["总评成绩", "总成绩", "总分", "最终成绩", "综合成绩", "期末总评", "卷面成绩"]
TARGET_RE = re.compile(r"课程目标\s*([一二三四五六七八九十\d]+)")
SEQUENTIAL_TARGET_HEADERS = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


@dataclass
class TargetGroup:
    number: int
    source_columns: list[int]
    denominator: float | None = None
    source_headers: list[str] = field(default_factory=list)


@dataclass
class TargetTablePlan:
    sheet_name: str
    header_row_index: int
    data_start_row_index: int
    student_id_column: int
    student_name_column: int
    target_groups: list[TargetGroup]
    total_column: int | None


@dataclass
class TargetWorkflowOutput:
    workbook: WorkbookData
    output: pd.DataFrame
    plan: TargetTablePlan
    warnings: list[str] = field(default_factory=list)
    sample_rows: pd.DataFrame = field(default_factory=pd.DataFrame)


def run_target_workflow(*, file_path: str) -> TargetWorkflowOutput:
    workbook = read_workbook(file_path)
    return execute_target_workflow(workbook)


def execute_target_workflow(workbook: WorkbookData) -> TargetWorkflowOutput:
    plan = _find_best_plan(workbook)
    raw = workbook.sheets[plan.sheet_name].frame
    data = raw.iloc[plan.data_start_row_index :].reset_index(drop=True)
    data = _trim_student_rows(data, plan.student_id_column, plan.student_name_column)
    if data.empty:
        raise ValueError("识别到表头后没有找到有效学生数据。")

    output = pd.DataFrame()
    output[ID_COLUMN] = data.iloc[:, plan.student_id_column].map(_clean_student_id)
    output[NAME_COLUMN] = data.iloc[:, plan.student_name_column].map(lambda value: clean_cell(value).replace(" ", ""))

    warnings: list[str] = []
    target_columns: list[str] = []
    for group in sorted(plan.target_groups, key=lambda item: item.number):
        column_name = f"{TARGET_PREFIX}{group.number}"
        target_columns.append(column_name)
        output[column_name] = _build_target_score(data, group)

    if plan.total_column is not None:
        output[TOTAL_COLUMN] = data.iloc[:, plan.total_column].map(_score_to_number)
    else:
        output[TOTAL_COLUMN] = _fallback_total(output[target_columns])
        warnings.append("未识别到总分列，已按课程目标均权计算百分制总分。")

    output = _clean_output(output, target_columns)
    if output.empty:
        raise ValueError("清洗后没有有效成绩数据。")

    return TargetWorkflowOutput(
        workbook=workbook,
        output=output[[ID_COLUMN, NAME_COLUMN, *target_columns, TOTAL_COLUMN]],
        plan=plan,
        warnings=warnings,
        sample_rows=data.head(10),
    )


def _find_best_plan(workbook: WorkbookData) -> TargetTablePlan:
    candidates: list[tuple[float, TargetTablePlan]] = []
    for sheet_name, sheet in workbook.sheets.items():
        frame = sheet.frame
        for header_row_index in range(min(len(frame), 20)):
            plan = _build_plan_for_row(sheet_name, frame, header_row_index)
            if plan is None:
                continue
            valid_rows = _count_student_rows(
                frame.iloc[plan.data_start_row_index :],
                plan.student_id_column,
                plan.student_name_column,
            )
            if valid_rows == 0:
                continue
            score = (
                100
                + len(plan.target_groups) * 20
                + min(valid_rows, 80)
                + (35 if plan.total_column is not None else 0)
                - frame.shape[1] * 0.3
            )
            candidates.append((score, plan))
    if not candidates:
        raise ValueError("未能找到包含学号、姓名和课程目标成绩的表头。")
    return max(candidates, key=lambda item: item[0])[1]


def _build_plan_for_row(sheet_name: str, frame: pd.DataFrame, header_row_index: int) -> TargetTablePlan | None:
    row = _row_values(frame, header_row_index)
    student_id_column = _find_header_column(row, ID_HEADERS)
    student_name_column = _find_header_column(row, NAME_HEADERS)
    if student_id_column is None or student_name_column is None:
        return None

    data_start_row_index = _find_data_start(frame, header_row_index + 1, student_id_column, student_name_column)
    if data_start_row_index is None:
        return None

    total_column = _find_total_column(frame, header_row_index, data_start_row_index)
    target_groups = _find_target_groups(frame, header_row_index, data_start_row_index, total_column)
    if not target_groups:
        target_groups = _find_sequential_target_groups(frame, header_row_index, data_start_row_index, total_column)
    if not target_groups:
        return None

    target_groups = [
        group
        for group in target_groups
        if group.source_columns
        and all(column not in {student_id_column, student_name_column, total_column} for column in group.source_columns)
    ]
    if not target_groups:
        return None

    return TargetTablePlan(
        sheet_name=sheet_name,
        header_row_index=header_row_index,
        data_start_row_index=data_start_row_index,
        student_id_column=student_id_column,
        student_name_column=student_name_column,
        target_groups=target_groups,
        total_column=total_column,
    )


def _find_target_groups(
    frame: pd.DataFrame,
    header_row_index: int,
    data_start_row_index: int,
    total_column: int | None,
) -> list[TargetGroup]:
    best_row_index: int | None = None
    best_targets: list[int | None] = []
    best_count = 0
    for row_index in range(header_row_index, min(data_start_row_index, header_row_index + 5, len(frame))):
        filled = _fill_forward(_row_values(frame, row_index))
        targets = [_target_number(value) for value in filled]
        count = sum(1 for value in targets if value is not None)
        if count > best_count:
            best_count = count
            best_row_index = row_index
            best_targets = targets

    if best_row_index is None or best_count == 0:
        return []

    raw_groups: dict[int, list[int]] = {}
    for column_index, number in enumerate(best_targets):
        if number is None or column_index == total_column:
            continue
        raw_groups.setdefault(number, []).append(column_index)

    groups: list[TargetGroup] = []
    for number, columns in raw_groups.items():
        source_columns, denominator, headers = _select_group_source_columns(
            frame,
            columns,
            best_row_index,
            data_start_row_index,
        )
        groups.append(TargetGroup(number=number, source_columns=source_columns, denominator=denominator, source_headers=headers))
    return sorted(groups, key=lambda item: item.number)


def _find_sequential_target_groups(
    frame: pd.DataFrame,
    header_row_index: int,
    data_start_row_index: int,
    total_column: int | None,
) -> list[TargetGroup]:
    row = _row_values(frame, header_row_index)
    groups: list[TargetGroup] = []
    for column_index, value in enumerate(row):
        if column_index == total_column:
            continue
        normalized = _normalize(value)
        number = SEQUENTIAL_TARGET_HEADERS.get(normalized)
        if number is None and normalized.isdigit():
            parsed = int(normalized)
            if 1 <= parsed <= 20:
                number = parsed
        if number is None:
            continue
        if not _has_numeric_data(frame, data_start_row_index, column_index):
            continue
        groups.append(TargetGroup(number=number, source_columns=[column_index], source_headers=[clean_cell(value)]))
    return sorted(groups, key=lambda item: item.number)


def _select_group_source_columns(
    frame: pd.DataFrame,
    columns: list[int],
    target_row_index: int,
    data_start_row_index: int,
) -> tuple[list[int], float | None, list[str]]:
    ratio_columns = [
        column
        for column in columns
        if _column_header_text(frame, target_row_index + 1, data_start_row_index, column).find("达成度") >= 0
        or _column_header_text(frame, target_row_index + 1, data_start_row_index, column).find("比例") >= 0
    ]
    if ratio_columns:
        return ratio_columns, None, [_column_header_text(frame, target_row_index, data_start_row_index, column) for column in ratio_columns]

    max_row_index = _find_max_score_row(frame, columns, target_row_index + 1, data_start_row_index)
    numeric_columns = [column for column in columns if _has_numeric_data(frame, data_start_row_index, column)]
    if max_row_index is None:
        return numeric_columns, None, [_column_header_text(frame, target_row_index, data_start_row_index, column) for column in numeric_columns]

    max_values = {column: _numeric_value(frame.iat[max_row_index, column]) for column in numeric_columns}
    component_columns = [
        column
        for column in numeric_columns
        if max_values.get(column) is not None
        and max_values[column] > 0
        and _column_header_text(frame, target_row_index + 1, max_row_index, column)
    ]
    if len(component_columns) >= 2:
        total_like = []
        for column in component_columns:
            other_sum = sum(float(max_values[other]) for other in component_columns if other != column)
            if abs(float(max_values[column]) - other_sum) < 0.000001:
                total_like.append(column)
        component_columns = [column for column in component_columns if column not in total_like] or component_columns

    if not component_columns:
        component_columns = numeric_columns
    denominator = sum(float(max_values[column]) for column in component_columns if max_values.get(column) is not None)
    if denominator <= 0:
        denominator = None
    headers = [_column_header_text(frame, target_row_index, max_row_index, column) for column in component_columns]
    return component_columns, denominator, headers


def _find_total_column(frame: pd.DataFrame, header_row_index: int, data_start_row_index: int) -> int | None:
    best: tuple[int, int] | None = None
    for row_index in range(header_row_index, min(data_start_row_index, header_row_index + 5, len(frame))):
        for column_index, value in enumerate(_row_values(frame, row_index)):
            text = _normalize(value)
            if not text:
                continue
            for rank, keyword in enumerate(TOTAL_KEYWORDS):
                if keyword in text:
                    current = (rank, column_index)
                    if best is None or current < best:
                        best = current
                    break
    return None if best is None else best[1]


def _find_data_start(frame: pd.DataFrame, start_row_index: int, student_id_column: int, student_name_column: int) -> int | None:
    for row_index in range(start_row_index, len(frame)):
        if _looks_like_student_id(frame.iat[row_index, student_id_column]) and clean_cell(frame.iat[row_index, student_name_column]):
            return row_index
    return None


def _find_max_score_row(frame: pd.DataFrame, columns: list[int], start_row_index: int, end_row_index: int) -> int | None:
    best: tuple[int, int] | None = None
    for row_index in range(start_row_index, end_row_index):
        count = sum(1 for column in columns if _numeric_value(frame.iat[row_index, column]) is not None)
        if count > 0 and (best is None or count >= best[0]):
            best = (count, row_index)
    return None if best is None else best[1]


def _build_target_score(data: pd.DataFrame, group: TargetGroup) -> pd.Series:
    if group.denominator is not None:
        return data.apply(lambda row: _round_or_none(_row_sum(row, group.source_columns) / group.denominator * 100), axis=1)

    if len(group.source_columns) == 1:
        column = group.source_columns[0]
        return data.iloc[:, column].map(_score_to_percent_if_ratio)

    return data.apply(lambda row: _round_or_none(_row_sum(row, group.source_columns)), axis=1)


def _fallback_total(target_scores: pd.DataFrame) -> pd.Series:
    totals: list[float | None] = []
    for _, row in target_scores.iterrows():
        values = [float(value) for value in row.tolist() if pd.notna(value)]
        if not values:
            totals.append(None)
        else:
            totals.append(_round_or_none(sum(values) / len(values)))
    return pd.Series(totals, index=target_scores.index)


def _trim_student_rows(data: pd.DataFrame, student_id_column: int, student_name_column: int) -> pd.DataFrame:
    valid = data.apply(
        lambda row: _looks_like_student_id(row.iloc[student_id_column]) and bool(clean_cell(row.iloc[student_name_column])),
        axis=1,
    )
    return data.loc[valid].reset_index(drop=True)


def _clean_output(output: pd.DataFrame, target_columns: list[str]) -> pd.DataFrame:
    cleaned = output.copy()
    cleaned = cleaned.loc[
        cleaned[ID_COLUMN].map(lambda value: bool(clean_cell(value)))
        & cleaned[NAME_COLUMN].map(lambda value: bool(clean_cell(value)))
    ]
    score_columns = [*target_columns, TOTAL_COLUMN]
    cleaned = cleaned.loc[
        cleaned[score_columns].apply(lambda row: any(pd.notna(value) for value in row.tolist()), axis=1)
    ]
    return cleaned.drop_duplicates().reset_index(drop=True)


def _count_student_rows(data: pd.DataFrame, student_id_column: int, student_name_column: int) -> int:
    return int(
        data.apply(
            lambda row: _looks_like_student_id(row.iloc[student_id_column])
            and bool(clean_cell(row.iloc[student_name_column])),
            axis=1,
        ).sum()
    )


def _row_values(frame: pd.DataFrame, row_index: int) -> list[Any]:
    return frame.iloc[row_index].tolist()


def _find_header_column(row: list[Any], candidates: set[str]) -> int | None:
    for index, value in enumerate(row):
        if _normalize(value) in candidates:
            return index
    return None


def _fill_forward(values: list[Any]) -> list[str]:
    filled: list[str] = []
    current = ""
    for value in values:
        text = clean_cell(value)
        if text:
            current = text
        filled.append(current)
    return filled


def _target_number(value: Any) -> int | None:
    text = _normalize(value)
    match = TARGET_RE.search(text)
    if not match:
        return None
    raw = match.group(1)
    if raw.isdigit():
        return int(raw)
    return _chinese_number_to_int(raw)


def _chinese_number_to_int(text: str) -> int | None:
    digits = SEQUENTIAL_TARGET_HEADERS
    if text in digits:
        return digits[text]
    if text == "十":
        return 10
    if text.startswith("十"):
        return 10 + digits.get(text[1:], 0)
    if "十" in text:
        left, right = text.split("十", 1)
        return digits.get(left, 1) * 10 + digits.get(right, 0)
    return None


def _column_header_text(frame: pd.DataFrame, start_row_index: int, end_row_index: int, column: int) -> str:
    parts = []
    for row_index in range(start_row_index, min(end_row_index, len(frame))):
        text = clean_cell(frame.iat[row_index, column])
        if text:
            parts.append(text)
    return " ".join(parts)


def _has_numeric_data(frame: pd.DataFrame, data_start_row_index: int, column: int) -> bool:
    sample = frame.iloc[data_start_row_index : data_start_row_index + 20, column]
    return any(_numeric_value(value) is not None for value in sample.tolist())


def _row_sum(row: pd.Series, columns: list[int]) -> float:
    values = [_numeric_value(row.iloc[column]) for column in columns]
    return sum(float(value) for value in values if value is not None)


def _score_to_percent_if_ratio(value: Any) -> float | None:
    number = _numeric_value(value)
    if number is None:
        return None
    if 0 <= number <= 1:
        number *= 100
    return _round_or_none(number)


def _score_to_number(value: Any) -> float | None:
    return _round_or_none(_numeric_value(value))


def _numeric_value(value: Any) -> float | None:
    text = clean_cell(value)
    if not text:
        return None
    percent = "%" in text
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(match.group(0))
    return number / 100 if percent else number


def _round_or_none(value: float | None) -> float | None:
    if value is None:
        return None
    return round(float(value), 2)


def _clean_student_id(value: Any) -> str:
    text = clean_cell(value).replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _looks_like_student_id(value: Any) -> bool:
    text = _clean_student_id(value)
    if not text or _normalize(text) in ID_HEADERS:
        return False
    return sum(char.isdigit() for char in text) >= 4


def _normalize(value: Any) -> str:
    return re.sub(r"\s+", "", clean_cell(value))
