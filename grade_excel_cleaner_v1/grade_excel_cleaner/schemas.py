from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, field_validator


TableTypeGuess = Literal["single_course", "multi_course", "unknown"]
DataEndStrategy = Literal[
    "until_empty_student_id",
    "until_empty_row",
    "all_rows_after_header",
]
MetadataSource = Literal[
    "column",
    "filename",
    "sheet_name",
    "top_rows",
    "constant",
    "unknown",
]


class ColumnMapping(BaseModel):
    student_id: str | None = None
    student_name: str | None = None
    course_name: str | None = None
    class_name: str | None = None
    final_score: str | None = None


class MetadataItem(BaseModel):
    source: MetadataSource = "unknown"
    value: str | None = None
    reason: str = ""


class MetadataMapping(BaseModel):
    course_name: MetadataItem = Field(default_factory=MetadataItem)
    class_name: MetadataItem = Field(default_factory=MetadataItem)


class ExtractionPlan(BaseModel):
    sheet_name: str
    table_type_guess: TableTypeGuess = "unknown"
    header_row_index: int = Field(ge=0)
    data_start_row_index: int = Field(ge=0)
    data_end_strategy: DataEndStrategy = "until_empty_student_id"
    column_mapping: ColumnMapping
    metadata_mapping: MetadataMapping = Field(default_factory=MetadataMapping)
    final_score_reason: str = ""
    confidence: float = Field(default=0.0, ge=0.0, le=1.0)
    warnings: list[str] = Field(default_factory=list)

    @field_validator("data_start_row_index")
    @classmethod
    def data_starts_after_header(cls, value: int, info):
        header = info.data.get("header_row_index")
        if header is not None and value <= header:
            raise ValueError("data_start_row_index must be greater than header_row_index")
        return value


class WorkflowResult(BaseModel):
    plan: ExtractionPlan
    rows: int
    table_type_detected: str
    warnings: list[str] = Field(default_factory=list)
    repaired: bool = False
