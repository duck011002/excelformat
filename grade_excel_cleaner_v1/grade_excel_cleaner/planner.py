from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .excel_reader import WorkbookData, read_workbook
from .executor import ExecutionResult, execute_plan
from .llm_client import call_openai_compatible
from .preview_builder import build_workbook_preview, preview_to_json
from .schemas import ExtractionPlan
from .utils import extract_json_object
from .validator import validate_output


SYSTEM_PROMPT = """
你是高校成绩 Excel 表格结构分析专家。
你的任务不是输出最终成绩数据，也不是生成 CSV，而是根据 workbook preview 输出一个 extraction_plan JSON。
你要识别 sheet、表头行、数据起始行、字段映射、课程名来源、班级名来源和最终成绩列。
必须只输出 JSON，不要输出 markdown，不要解释，不要包含额外文本。
""".strip()


PLAN_INSTRUCTIONS = """
目标输出字段固定为：学号、学生姓名、课程名、班级名、最终成绩。

核心规则：
1. 如果原表存在课程名、课程名称、课程、科目、课程号+课程名、教学班课程等语义等价列，column_mapping.course_name 必须直接映射该列。
2. 如果没有课程名列，通常视为单课程表，课程名从前几行标题、sheet 名、文件名中提取；无法判断时填“未知课程”，并在 warnings 说明。
3. 如果没有班级名列，优先从前几行、sheet 名、文件名提取；无法判断时可为空。
4. 最终成绩优先级：总评成绩、总成绩、最终成绩、成绩、期末总评、综合成绩。
5. 不要随便把平时成绩、卷面成绩当作最终成绩，除非表格明确只有该成绩，且标题或文件名表明它就是本表目标成绩。
6. 如果存在多个最终成绩候选列，必须在 final_score_reason 中说明选择理由。
7. row_index 使用 0-based index。
8. column_mapping 中写原始表头名，不写列号；字段来自元数据时对应 column_mapping 字段为 null。

严格输出 JSON，结构如下：
{
  "sheet_name": "string",
  "table_type_guess": "single_course | multi_course | unknown",
  "header_row_index": 0,
  "data_start_row_index": 1,
  "data_end_strategy": "until_empty_student_id | until_empty_row | all_rows_after_header",
  "column_mapping": {
    "student_id": "string | null",
    "student_name": "string | null",
    "course_name": "string | null",
    "class_name": "string | null",
    "final_score": "string | null"
  },
  "metadata_mapping": {
    "course_name": {
      "source": "column | filename | sheet_name | top_rows | constant | unknown",
      "value": "string | null",
      "reason": "string"
    },
    "class_name": {
      "source": "column | filename | sheet_name | top_rows | constant | unknown",
      "value": "string | null",
      "reason": "string"
    }
  },
  "final_score_reason": "string",
  "confidence": 0.0,
  "warnings": ["string"]
}
""".strip()


@dataclass
class WorkflowOutput:
    workbook: WorkbookData
    preview: dict
    plan: ExtractionPlan
    execution: ExecutionResult
    repaired: bool
    raw_plan_response: str


def call_llm_for_plan(
    *,
    preview: dict,
    base_url: str,
    api_key: str,
    model: str,
) -> tuple[ExtractionPlan, str]:
    user_prompt = (
        f"{PLAN_INSTRUCTIONS}\n\n"
        "下面是 workbook preview：\n"
        f"{preview_to_json(preview)}"
    )
    raw = call_openai_compatible(
        base_url=base_url,
        api_key=api_key,
        model=model,
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
    )
    return ExtractionPlan.model_validate(extract_json_object(raw)), raw


def call_llm_for_repair(
    *,
    preview: dict,
    previous_plan: ExtractionPlan,
    error: str,
    columns: list[str],
    sample_rows: pd.DataFrame,
    base_url: str,
    api_key: str,
    model: str,
) -> tuple[ExtractionPlan, str]:
    repair_context = {
        "previous_plan": previous_plan.model_dump(),
        "validation_or_execution_error": error,
        "available_columns_after_header": columns,
        "sample_rows": sample_rows.fillna("").head(10).to_dict(orient="records"),
        "original_preview": preview,
    }
    user_prompt = (
        f"{PLAN_INSTRUCTIONS}\n\n"
        "上一次 extraction_plan 执行或校验失败。请基于以下反馈重新输出完整 extraction_plan JSON：\n"
        f"{preview_to_json(repair_context)}"
    )
    raw = call_openai_compatible(
        base_url=base_url,
        api_key=api_key,
        model=model,
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
    )
    return ExtractionPlan.model_validate(extract_json_object(raw)), raw


def run_workflow(
    *,
    file_path: str,
    preview_rows: int,
    base_url: str,
    api_key: str,
    model: str,
    enable_repair: bool = True,
) -> WorkflowOutput:
    workbook = read_workbook(file_path)
    preview = build_workbook_preview(workbook, preview_rows=preview_rows)
    plan, raw = call_llm_for_plan(
        preview=preview,
        base_url=base_url,
        api_key=api_key,
        model=model,
    )

    try:
        execution = execute_plan(workbook, plan)
        validate_output(execution.output)
        return WorkflowOutput(workbook, preview, plan, execution, False, raw)
    except Exception as exc:
        if not enable_repair:
            raise
        columns, sample_rows = _plan_columns_and_sample(workbook, plan)
        repaired_plan, repaired_raw = call_llm_for_repair(
            preview=preview,
            previous_plan=plan,
            error=str(exc),
            columns=columns,
            sample_rows=sample_rows,
            base_url=base_url,
            api_key=api_key,
            model=model,
        )
        execution = execute_plan(workbook, repaired_plan)
        validate_output(execution.output)
        return WorkflowOutput(workbook, preview, repaired_plan, execution, True, repaired_raw)


def _plan_columns_and_sample(
    workbook: WorkbookData, plan: ExtractionPlan
) -> tuple[list[str], pd.DataFrame]:
    sheet = workbook.sheets.get(plan.sheet_name)
    if sheet is None:
        return [], pd.DataFrame()
    frame = sheet.frame
    if plan.header_row_index >= len(frame):
        return [], pd.DataFrame()
    from .utils import make_unique_headers

    columns = make_unique_headers(frame.iloc[plan.header_row_index].tolist())
    sample = frame.iloc[plan.data_start_row_index : plan.data_start_row_index + 10].copy()
    sample.columns = columns[: len(sample.columns)]
    return columns, sample
