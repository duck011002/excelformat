from __future__ import annotations

import copy
import json
from typing import Any, Callable

from ..utils import extract_json_object
from .schemas import ResolvedField, TeachingClassGroup, TeachingClassResult, WorkbookEvidence


FIELD_MAP = {
    "教学班编号": "class_code",
    "教学班名称": "class_name",
    "运行学年": "school_year",
    "学年内学期": "term",
    "任课老师名称": "teacher_name",
    "任课老师工号": "teacher_id",
}
SYSTEM_PROMPT = """
你是高校教学班 Excel 元数据复核助手。只判断不确定字段，不处理整张成绩表。
必须只输出 JSON；不得编造学生、课程或教师信息。证据不足时 value 为空。
""".strip()


def build_llm_context(
    groups: list[TeachingClassGroup],
    evidence: WorkbookEvidence,
    *,
    enhanced: bool = False,
) -> dict[str, Any]:
    candidate_limit = 8 if enhanced else 5
    evidence_limit = 80 if enhanced else 40
    context_groups: list[dict[str, Any]] = []
    for group in groups:
        candidates: dict[str, list[str]] = {}
        current: dict[str, dict[str, Any]] = {}
        for output_name, attr_name in FIELD_MAP.items():
            field: ResolvedField = getattr(group, attr_name)
            values = [field.value, *field.candidates, *field.evidence]
            candidates[output_name] = _unique_non_empty(values)[:candidate_limit]
            current[output_name] = {
                "value": field.value,
                "confidence": round(field.confidence, 3),
                "locked_by_user": field.locked_by_user,
            }
        context_groups.append(
            {
                "group_id": group.group_id,
                "course_name": group.course_name,
                "student_count": len(group.students),
                "current": current,
                "candidates": candidates,
            }
        )
    return {
        "mode": "enhanced" if enhanced else "compact",
        "instruction": "只返回有充分证据且需要修正的未锁定字段。学期格式统一为第一学期、第二学期或第三学期。",
        "groups": context_groups,
        "evidence": [
            {
                "source": item.source_type,
                "sheet": item.sheet_name,
                "cell": _cell_label(item.row_index, item.column_index),
                "text": item.text,
            }
            for item in evidence.items[:evidence_limit]
        ],
        "response_schema": {
            "groups": [
                {
                    "group_id": "string",
                    "教学班名称": {"value": "string", "confidence": 0.0},
                }
            ]
        },
    }


def resolve_uncertain_groups(
    result: TeachingClassResult,
    *,
    caller: Callable[..., str],
    enhanced: bool = False,
) -> TeachingClassResult:
    updated = copy.deepcopy(result)
    context = build_llm_context(updated.groups, updated.evidence, enhanced=enhanced)
    try:
        raw = caller(
            system_prompt=SYSTEM_PROMPT,
            user_prompt=json.dumps(context, ensure_ascii=False, separators=(",", ":")),
        )
        payload = extract_json_object(raw)
        response_groups = payload.get("groups")
        if not isinstance(response_groups, list):
            raise ValueError("响应缺少 groups 数组")
        group_lookup = {group.group_id: group for group in updated.groups}
        for response_group in response_groups:
            if not isinstance(response_group, dict):
                continue
            group = group_lookup.get(str(response_group.get("group_id", "")))
            if group is None:
                continue
            _apply_group_resolution(group, response_group)
    except Exception as exc:
        updated.warnings.append(f"AI 再次解析失败，已保留 Python 结果：{exc}")
    updated.warnings = list(dict.fromkeys(updated.warnings))
    return updated


def _apply_group_resolution(group: TeachingClassGroup, response: dict[str, Any]) -> None:
    before = _snapshot(group)
    changed = False
    for output_name, attr_name in FIELD_MAP.items():
        field: ResolvedField = getattr(group, attr_name)
        if field.locked_by_user or (field.value and field.confidence >= 0.75):
            continue
        parsed = _parse_field_response(response.get(output_name))
        if parsed is None:
            continue
        value, confidence = parsed
        if not value or confidence < 0.45:
            continue
        field.value = value
        field.confidence = min(max(confidence, 0.0), 1.0)
        field.source_type = "llm_reanalysis"
        field.evidence.append("AI 基于候选证据再次解析")
        field.candidates = _unique_non_empty([value, *field.candidates])
        changed = True
    if changed:
        group.revisions.append(before)


def _parse_field_response(value: Any) -> tuple[str, float] | None:
    if isinstance(value, str):
        return value.strip(), 0.6
    if not isinstance(value, dict):
        return None
    text = str(value.get("value", "")).strip()
    try:
        confidence = float(value.get("confidence", 0.6))
    except (TypeError, ValueError):
        confidence = 0.6
    return text, confidence


def _snapshot(group: TeachingClassGroup) -> dict[str, str]:
    return {output_name: getattr(group, attr_name).value for output_name, attr_name in FIELD_MAP.items()}


def _unique_non_empty(values: list[str]) -> list[str]:
    return list(dict.fromkeys(str(value).strip() for value in values if str(value).strip()))


def _cell_label(row_index: int | None, column_index: int | None) -> str:
    if row_index is None:
        return ""
    return f"R{row_index + 1}C{(column_index or 0) + 1}"
