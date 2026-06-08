from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


READ_ERROR_MESSAGE = (
    "文件读取失败，请确认文件格式是否为 .xls/.xlsx/.xlsb，"
    "或尝试另存为 .xlsx 后重试。"
)


@dataclass
class SheetData:
    name: str
    frame: pd.DataFrame


@dataclass
class WorkbookData:
    filename: str
    sheets: dict[str, SheetData]
    engine_used: str
    read_warnings: list[str]


def _candidate_engines(path: Path) -> list[str | None]:
    suffix = path.suffix.lower()
    if suffix == ".xlsb":
        return ["calamine", "pyxlsb"]
    if suffix == ".xls":
        return ["calamine", "xlrd"]
    if suffix == ".xlsx":
        return ["calamine", "openpyxl"]
    return ["calamine", None]


def read_workbook(path: str | Path) -> WorkbookData:
    source = Path(path)
    warnings: list[str] = []
    last_error: Exception | None = None

    for engine in _candidate_engines(source):
        try:
            kwargs = {"sheet_name": None, "header": None, "dtype": object}
            if engine:
                kwargs["engine"] = engine
            raw = pd.read_excel(source, **kwargs)
            sheets = {
                name: SheetData(name=name, frame=frame)
                for name, frame in raw.items()
                if frame is not None and not frame.empty
            }
            if not sheets:
                raise ValueError("workbook contains no readable non-empty sheets")
            return WorkbookData(
                filename=source.name,
                sheets=sheets,
                engine_used=engine or "pandas-default",
                read_warnings=warnings,
            )
        except Exception as exc:
            last_error = exc
            warnings.append(f"{engine or 'pandas-default'} 读取失败：{exc}")

    raise ValueError(f"{READ_ERROR_MESSAGE}\n最后一次错误：{last_error}")
