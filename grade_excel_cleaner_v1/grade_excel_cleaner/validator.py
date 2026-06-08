from __future__ import annotations

import pandas as pd

from .executor import OUTPUT_COLUMNS
from .utils import clean_cell


def validate_output(frame: pd.DataFrame) -> None:
    missing = [column for column in OUTPUT_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError(f"输出缺少固定列：{missing}")
    if frame.empty:
        raise ValueError("输出 DataFrame 为空。")

    required = ["学号", "学生姓名", "最终成绩"]
    for column in required:
        empty_ratio = frame[column].map(lambda value: not clean_cell(value)).mean()
        if empty_ratio > 0.35:
            raise ValueError(f"字段“{column}”空值比例过高：{empty_ratio:.0%}")

    valid_student_id_ratio = frame["学号"].map(_looks_like_student_id).mean()
    if valid_student_id_ratio < 0.5:
        raise ValueError(
            f"学号列有效值比例过低：{valid_student_id_ratio:.0%}，可能表头行或字段映射错误。"
        )


def _looks_like_student_id(value: object) -> bool:
    text = clean_cell(value)
    if not text or text == "学号":
        return False
    digit_count = sum(char.isdigit() for char in text)
    return digit_count >= 4
