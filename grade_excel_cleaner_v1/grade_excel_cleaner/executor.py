from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from .excel_reader import WorkbookData
from .schemas import ExtractionPlan
from .utils import clean_cell, make_unique_headers, normalize_header


OUTPUT_COLUMNS = ["学号", "学生姓名", "课程名", "班级名", "最终成绩"]


@dataclass
class ExecutionResult:
    output: pd.DataFrame
    table_type_detected: str
    warnings: list[str] = field(default_factory=list)
    source_columns: list[str] = field(default_factory=list)
    sample_rows: pd.DataFrame = field(default_factory=pd.DataFrame)


def execute_plan(workbook: WorkbookData, plan: ExtractionPlan) -> ExecutionResult:
    if plan.sheet_name not in workbook.sheets:
        raise ValueError(f"未找到 sheet：{plan.sheet_name}。可用 sheet：{list(workbook.sheets)}")

    raw = workbook.sheets[plan.sheet_name].frame.copy()
    if plan.header_row_index >= len(raw):
        raise ValueError("header_row_index 超出表格行数。")
    if plan.data_start_row_index >= len(raw):
        raise ValueError("data_start_row_index 超出表格行数。")

    headers = make_unique_headers(raw.iloc[plan.header_row_index].tolist())
    data = raw.iloc[plan.data_start_row_index :].copy()
    data.columns = headers[: len(data.columns)]
    data = data.reset_index(drop=True)

    data = _apply_end_strategy(data, plan)
    column_lookup = _build_column_lookup(list(data.columns))
    output = pd.DataFrame()

    output["学号"] = _read_required(data, column_lookup, plan.column_mapping.student_id, "学号")
    output["学生姓名"] = _read_required(
        data, column_lookup, plan.column_mapping.student_name, "学生姓名"
    )
    output["课程名"] = _read_optional_or_metadata(
        data,
        column_lookup,
        plan.column_mapping.course_name,
        plan.metadata_mapping.course_name.value,
        "未知课程",
    )
    output["班级名"] = _read_optional_or_metadata(
        data,
        column_lookup,
        plan.column_mapping.class_name,
        plan.metadata_mapping.class_name.value,
        "",
    )
    output["最终成绩"] = _read_required(data, column_lookup, plan.column_mapping.final_score, "最终成绩")

    output = _clean_output(output)
    table_type = _detect_table_type(output)
    warnings = list(plan.warnings)
    if plan.metadata_mapping.course_name.source != "column" and not plan.column_mapping.course_name:
        warnings.append(f"课程名来自 {plan.metadata_mapping.course_name.source}。")
    if output.empty:
        raise ValueError("按 extraction_plan 清洗后没有有效数据。")

    return ExecutionResult(
        output=output[OUTPUT_COLUMNS],
        table_type_detected=table_type,
        warnings=warnings,
        source_columns=list(data.columns),
        sample_rows=data.head(10),
    )


def _build_column_lookup(columns: list[str]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for column in columns:
        lookup[normalize_header(column)] = column
    return lookup


def _resolve_column(lookup: dict[str, str], mapped: str | None) -> str | None:
    if not mapped:
        return None
    key = normalize_header(mapped)
    if key in lookup:
        return lookup[key]
    for normalized, original in lookup.items():
        if key and (key in normalized or normalized in key):
            return original
    return None


def _read_required(
    data: pd.DataFrame, lookup: dict[str, str], mapped: str | None, target_name: str
) -> pd.Series:
    column = _resolve_column(lookup, mapped)
    if column is None:
        raise ValueError(f"必须字段“{target_name}”未能映射到有效原始列：{mapped}")
    return data[column]


def _read_optional_or_metadata(
    data: pd.DataFrame,
    lookup: dict[str, str],
    mapped: str | None,
    metadata_value: str | None,
    fallback: str,
) -> pd.Series:
    column = _resolve_column(lookup, mapped)
    if column is not None:
        return data[column]
    value = clean_cell(metadata_value) or fallback
    return pd.Series([value] * len(data), index=data.index)


def _apply_end_strategy(data: pd.DataFrame, plan: ExtractionPlan) -> pd.DataFrame:
    if plan.data_end_strategy == "all_rows_after_header":
        return data
    if plan.data_end_strategy == "until_empty_row":
        empty_mask = data.apply(lambda row: all(not clean_cell(value) for value in row), axis=1)
        if empty_mask.any():
            return data.iloc[: empty_mask.idxmax()]
        return data
    if plan.data_end_strategy == "until_empty_student_id":
        lookup = _build_column_lookup(list(data.columns))
        column = _resolve_column(lookup, plan.column_mapping.student_id)
        if column:
            empty_mask = data[column].map(lambda value: not clean_cell(value))
            if empty_mask.any():
                first_empty = empty_mask[empty_mask].index[0]
                return data.loc[: first_empty - 1] if first_empty > 0 else data.iloc[:0]
    return data


def _clean_output(output: pd.DataFrame) -> pd.DataFrame:
    cleaned = output.copy()
    cleaned["学号"] = cleaned["学号"].map(_clean_student_id)
    cleaned["学生姓名"] = cleaned["学生姓名"].map(lambda value: clean_cell(value).replace(" ", ""))
    cleaned["课程名"] = cleaned["课程名"].map(clean_cell).replace("", "未知课程")
    cleaned["班级名"] = cleaned["班级名"].map(clean_cell)
    cleaned["最终成绩"] = cleaned["最终成绩"].map(_clean_score)

    invalid = (
        cleaned["学号"].eq("")
        | cleaned["学号"].eq("学号")
        | cleaned["学生姓名"].eq("")
        | cleaned["最终成绩"].eq("")
    )
    cleaned = cleaned.loc[~invalid].drop_duplicates().reset_index(drop=True)
    return cleaned


def _clean_student_id(value: Any) -> str:
    text = clean_cell(value).replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _clean_score(value: Any) -> str:
    text = clean_cell(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _detect_table_type(output: pd.DataFrame) -> str:
    if "课程名" not in output.columns or output.empty:
        return "unknown"
    course_count = output["课程名"].map(clean_cell).replace("", pd.NA).dropna().nunique()
    if course_count <= 2:
        return "single_course_like"
    return "multi_course"
