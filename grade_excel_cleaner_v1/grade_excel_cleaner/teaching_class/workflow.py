from __future__ import annotations

import hashlib
import re
from collections import OrderedDict
from typing import Any

import pandas as pd

from ..excel_reader import WorkbookData
from ..utils import clean_cell
from .academic_term import normalize_academic_term
from .scanner import scan_workbook
from .schemas import (
    ResolvedField,
    StudentRecord,
    TeachingClassGroup,
    TeachingClassResult,
    WorkbookEvidence,
)


TEACHING_CLASS_COLUMNS = [
    "教学班编号",
    "教学班名称",
    "课程名",
    "班级名",
    "运行学年",
    "学年内学期",
    "任课老师名称",
    "任课老师工号",
    "教学班学生姓名",
    "教学班学生学号",
]

STUDENT_ID_COLUMNS = ("教学班学生学号", "学生学号", "学号")
STUDENT_NAME_COLUMNS = ("教学班学生姓名", "学生姓名", "姓名")
COURSE_COLUMNS = ("课程名", "课程名称", "课程")
COURSE_ALIAS_COLUMNS = ("课程班别名", "课程班", "班级名")
EXPLICIT_CLASS_COLUMNS = ("教学班名称", "教学班")
ADMIN_CLASS_COLUMNS = ("行政班", "行政班级", "班级")
CLASS_CODE_COLUMNS = ("教学班编号", "教学班代码", "课程班编号")
TEACHER_NAME_COLUMNS = ("任课老师名称", "任课教师", "教师姓名", "老师姓名")
TEACHER_ID_COLUMNS = ("任课老师工号", "教师工号", "工号")


def build_teaching_class_result(
    workbook: WorkbookData,
    score_output: pd.DataFrame | None,
    *,
    expanded_scan: bool = False,
) -> TeachingClassResult:
    evidence = scan_workbook(workbook, expanded=expanded_scan)
    metadata = _resolve_metadata(evidence)
    rows = _normalize_score_rows(score_output)
    grouped_rows = _group_rows(rows, metadata["class_name"].value)
    groups: list[TeachingClassGroup] = []

    for group_key, group_rows in grouped_rows.items():
        first = group_rows[0] if group_rows else {}
        class_name, grouping_source = _row_group_name(first, metadata["class_name"].value)
        course_name = clean_cell(first.get("course_name")) or metadata.get("course_name", ResolvedField()).value or "未知课程"
        group_id = _stable_group_id(group_key, course_name)
        source_code = _first_value(group_rows, "class_code")
        if source_code:
            class_code = ResolvedField(
                value=source_code,
                confidence=0.98,
                source_type="source_column",
                evidence=[source_code],
                candidates=[source_code],
            )
        else:
            generated = _stable_numeric_code(group_id)
            class_code = ResolvedField(
                value=generated,
                confidence=0.35,
                source_type="generated",
                evidence=["源表未发现教学班编号，已生成稳定编号"],
                candidates=[generated],
                warning="教学班编号为稳定兜底值，建议人工确认。",
            )

        if class_name:
            class_name_field = ResolvedField(
                value=class_name,
                confidence=0.96 if grouping_source != "metadata" else metadata["class_name"].confidence,
                source_type=grouping_source,
                evidence=[class_name],
                candidates=[class_name, *metadata["class_name"].candidates],
            )
        else:
            class_name_field = ResolvedField(
                warning="未找到教学班名称，导出时将保留为空。",
                candidates=metadata["class_name"].candidates,
            )

        teacher_name = _first_value(group_rows, "teacher_name")
        teacher_id = _first_value(group_rows, "teacher_id")
        groups.append(
            TeachingClassGroup(
                group_id=group_id,
                course_name=course_name,
                grouping_source=grouping_source,
                class_code=class_code,
                class_name=class_name_field,
                school_year=_copy_field(metadata["school_year"]),
                term=_copy_field(metadata["term"]),
                teacher_name=_resolved_optional(teacher_name, "source_column", metadata["teacher_name"]),
                teacher_id=_resolved_optional(teacher_id, "source_column", metadata["teacher_id"]),
                students=[
                    StudentRecord(
                        student_id=clean_cell(row.get("student_id")),
                        student_name=clean_cell(row.get("student_name")),
                        course_name=clean_cell(row.get("course_name")),
                        class_name=clean_cell(row.get("course_class_alias")),
                    )
                    for row in group_rows
                ]
                or [StudentRecord()],
            )
        )

    warnings = _build_warnings(groups)
    return TeachingClassResult(groups=groups, evidence=evidence, warnings=warnings)


def flatten_groups(groups: list[TeachingClassGroup]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for group in groups:
        students = group.students or [StudentRecord()]
        for student in students:
            rows.append(
                {
                    "教学班编号": group.class_code.value,
                    "教学班名称": group.class_name.value,
                    "课程名": student.course_name or group.course_name,
                    "班级名": student.class_name or group.class_name.value,
                    "运行学年": group.school_year.value,
                    "学年内学期": group.term.value,
                    "任课老师名称": group.teacher_name.value,
                    "任课老师工号": group.teacher_id.value,
                    "教学班学生姓名": student.student_name,
                    "教学班学生学号": student.student_id,
                }
            )
    return pd.DataFrame(rows, columns=TEACHING_CLASS_COLUMNS)


def refresh_teaching_warnings(result: TeachingClassResult) -> None:
    ai_warnings = [warning for warning in result.warnings if warning.startswith("AI ")]
    result.warnings = list(dict.fromkeys([*_build_warnings(result.groups), *ai_warnings]))


def _normalize_score_rows(score_output: pd.DataFrame | None) -> list[dict[str, str]]:
    if score_output is None or score_output.empty:
        return [{}]
    columns = {str(column).strip(): column for column in score_output.columns}
    normalized: list[dict[str, str]] = []
    for _, row in score_output.iterrows():
        normalized.append(
            {
                "student_id": _row_value(row, columns, STUDENT_ID_COLUMNS),
                "student_name": _row_value(row, columns, STUDENT_NAME_COLUMNS),
                "course_name": _row_value(row, columns, COURSE_COLUMNS),
                "course_class_alias": _row_value(row, columns, COURSE_ALIAS_COLUMNS),
                "explicit_class": _row_value(row, columns, EXPLICIT_CLASS_COLUMNS),
                "admin_class": _row_value(row, columns, ADMIN_CLASS_COLUMNS),
                "class_code": _row_value(row, columns, CLASS_CODE_COLUMNS),
                "teacher_name": _row_value(row, columns, TEACHER_NAME_COLUMNS),
                "teacher_id": _row_value(row, columns, TEACHER_ID_COLUMNS),
            }
        )
    return normalized or [{}]


def _row_value(row: pd.Series, columns: dict[str, Any], candidates: tuple[str, ...]) -> str:
    for candidate in candidates:
        column = columns.get(candidate)
        if column is not None:
            value = clean_cell(row[column])
            if value:
                return value
    return ""


def _group_rows(rows: list[dict[str, str]], metadata_class_name: str) -> OrderedDict[str, list[dict[str, str]]]:
    groups: OrderedDict[str, list[dict[str, str]]] = OrderedDict()
    for row in rows:
        course = clean_cell(row.get("course_name"))
        class_code = clean_cell(row.get("class_code"))
        if row.get("course_class_alias"):
            identity = f"course_alias|{course}|{row['course_class_alias']}|{class_code}"
        elif row.get("explicit_class"):
            identity = f"explicit|{course}|{row['explicit_class']}|{class_code}"
        elif row.get("admin_class"):
            identity = f"admin|{course}|{row['admin_class']}|{class_code}"
        elif metadata_class_name:
            identity = f"metadata|{course}|{metadata_class_name}|{class_code}"
        else:
            identity = f"fallback|{course or 'unknown'}|{class_code}"
        groups.setdefault(identity, []).append(row)
    return groups


def _row_group_name(row: dict[str, str], metadata_class_name: str) -> tuple[str, str]:
    if row.get("course_class_alias"):
        return row["course_class_alias"], "course_class_alias"
    if row.get("explicit_class"):
        return row["explicit_class"], "explicit_teaching_class"
    if row.get("admin_class"):
        return row["admin_class"], "administrative_class"
    if metadata_class_name:
        return metadata_class_name, "metadata"
    return "", "fallback"


def _resolve_metadata(evidence: WorkbookEvidence) -> dict[str, ResolvedField]:
    class_candidates: list[tuple[str, float, str]] = []
    course_candidates: list[tuple[str, float, str]] = []
    teacher_candidates: list[tuple[str, float, str]] = []
    teacher_id_candidates: list[tuple[str, float, str]] = []
    term_candidates: list[tuple[float, str, str, str]] = []

    for item in evidence.items:
        source_confidence = 0.96 if item.source_type == "keyword_neighborhood" else 0.82
        for value in _extract_labeled_values(item.text, ("教学班名称", "教学班", "课程班别名", "课程班", "班级名称", "班级")):
            if _looks_like_class_candidate(value):
                class_candidates.append((value, source_confidence, item.text))
        for value in _extract_labeled_values(item.text, ("课程名称", "课程名", "课程")):
            if _valid_candidate(value):
                course_candidates.append((value, source_confidence, item.text))
        for value in _extract_labeled_values(item.text, ("任课老师名称", "任课教师", "教师姓名", "任课老师", "老师")):
            if _valid_candidate(value):
                teacher_candidates.append((value, source_confidence, item.text))
        for value in _extract_labeled_values(item.text, ("任课老师工号", "教师工号", "工号")):
            if _valid_candidate(value):
                teacher_id_candidates.append((value, source_confidence, item.text))

        normalized_term = normalize_academic_term(item.text)
        if normalized_term.school_year:
            term_candidates.append(
                (
                    normalized_term.confidence * max(item.score, 0.7),
                    normalized_term.school_year,
                    normalized_term.term,
                    item.text,
                )
            )

        if item.source_type in {"filename", "sheet_name"}:
            for value in _extract_unlabeled_class_tokens(item.text):
                class_candidates.append((value, 0.62, item.text))
        if item.source_type == "filename":
            for value in _extract_filename_teacher(item.text):
                teacher_candidates.append((value, 0.68, item.text))

    class_name = _field_from_candidates(class_candidates, "metadata")
    course_name = _field_from_candidates(course_candidates, "metadata")
    teacher_name = _field_from_candidates(teacher_candidates, "metadata")
    teacher_id = _field_from_candidates(teacher_id_candidates, "metadata")
    school_year = ResolvedField()
    term = ResolvedField()
    if term_candidates:
        confidence, year_value, term_value, source = max(term_candidates, key=lambda item: item[0])
        school_year = ResolvedField(
            value=year_value,
            confidence=confidence,
            source_type="metadata",
            evidence=[source],
            candidates=list(dict.fromkeys(candidate[1] for candidate in term_candidates)),
        )
        term = ResolvedField(
            value=term_value,
            confidence=confidence,
            source_type="metadata",
            evidence=[source],
            candidates=list(dict.fromkeys(candidate[2] for candidate in term_candidates if candidate[2])),
        )
    return {
        "class_name": class_name,
        "course_name": course_name,
        "school_year": school_year,
        "term": term,
        "teacher_name": teacher_name,
        "teacher_id": teacher_id,
    }


def _extract_labeled_values(text: str, labels: tuple[str, ...]) -> list[str]:
    values: list[str] = []
    for label in labels:
        pattern = re.compile(rf"{re.escape(label)}\s*(?:[:：=]\s*|\|\s*)([^|｜,，;；]{{1,50}})")
        values.extend(match.group(1).strip() for match in pattern.finditer(text))
    return values


def _extract_unlabeled_class_tokens(text: str) -> list[str]:
    compact = re.sub(r"\s+", "", text)
    patterns = (
        r"20\d{2}级[0-9+~～\-、至]+班",
        r"\d{2}[\u4e00-\u9fff]{1,6}[0-9+~～\-、至]+(?:班)?",
        r"[\u4e00-\u9fff]{1,8}\d{2}(?:[-、]\d+)*(?:班)?",
    )
    candidates: list[str] = []
    for pattern in patterns:
        candidates.extend(re.findall(pattern, compact))
    cleaned = [candidate.rstrip("-_－、~～+") for candidate in dict.fromkeys(candidates)]
    return [candidate for candidate in cleaned if _looks_like_class_candidate(candidate)]


def _extract_filename_teacher(text: str) -> list[str]:
    return [
        match.group(1)
        for match in re.finditer(r"[-_－]([\u4e00-\u9fff]{2,4})[（(]", text)
        if match.group(1) not in {"学年第", "成绩表"}
    ]


def _valid_candidate(value: str) -> bool:
    value = value.strip()
    return bool(value and len(value) <= 50 and value not in {"名称", "姓名", "编号"})


def _looks_like_class_candidate(value: str) -> bool:
    compact = re.sub(r"\s+", "", value)
    if not _valid_candidate(compact) or compact in {
        "班级",
        "教学班",
        "课程班",
        "成绩",
        "总成绩",
        "学号",
        "姓名",
        "学生姓名",
        "一",
        "二",
        "三",
    }:
        return False
    if any(keyword in compact for keyword in ("选择题", "填空题", "判断题", "得分", "总分", "成绩")):
        return False
    if "班" in compact:
        return True
    if re.fullmatch(r"[A-Za-z]*\d+[A-Za-z0-9_-]*", compact):
        return True
    return bool(re.search(r"\d", compact) and re.search(r"[\u4e00-\u9fffA-Za-z]", compact) and len(compact) <= 35)


def _field_from_candidates(candidates: list[tuple[str, float, str]], source_type: str) -> ResolvedField:
    if not candidates:
        return ResolvedField()
    deduplicated: dict[str, tuple[float, str]] = {}
    for value, confidence, evidence in candidates:
        current = deduplicated.get(value)
        if current is None or confidence > current[0]:
            deduplicated[value] = (confidence, evidence)
    best_value, (confidence, evidence) = max(deduplicated.items(), key=lambda item: item[1][0])
    return ResolvedField(
        value=best_value,
        confidence=confidence,
        source_type=source_type,
        evidence=[evidence],
        candidates=list(deduplicated)[:10],
    )


def _resolved_optional(row_value: str, row_source: str, metadata_field: ResolvedField) -> ResolvedField:
    if row_value:
        return ResolvedField(
            value=row_value,
            confidence=0.98,
            source_type=row_source,
            evidence=[row_value],
            candidates=[row_value, *metadata_field.candidates],
        )
    return _copy_field(metadata_field)


def _copy_field(field: ResolvedField) -> ResolvedField:
    return ResolvedField(
        value=field.value,
        confidence=field.confidence,
        source_type=field.source_type,
        evidence=list(field.evidence),
        candidates=list(field.candidates),
        warning=field.warning,
        locked_by_user=field.locked_by_user,
    )


def _first_value(rows: list[dict[str, str]], key: str) -> str:
    for row in rows:
        value = clean_cell(row.get(key))
        if value:
            return value
    return ""


def _stable_group_id(group_key: str, course_name: str) -> str:
    digest = hashlib.sha1(f"{course_name}|{group_key}".encode("utf-8")).hexdigest()[:12]
    return f"tc_{digest}"


def _stable_numeric_code(group_id: str) -> str:
    number = int(hashlib.sha1(group_id.encode("utf-8")).hexdigest()[:12], 16)
    return str(100000000 + number % 900000000)


def _build_warnings(groups: list[TeachingClassGroup]) -> list[str]:
    warnings: list[str] = []
    for index, group in enumerate(groups, start=1):
        label = group.class_name.value or f"第{index}个教学班"
        for field_name, field in (
            ("教学班名称", group.class_name),
            ("运行学年", group.school_year),
            ("学年内学期", group.term),
        ):
            if not field.value:
                warnings.append(f"{label}：未找到{field_name}，导出列将保留为空。")
        if group.class_code.warning:
            warnings.append(f"{label}：{group.class_code.warning}")
        if not any(student.student_name for student in group.students):
            warnings.append(f"{label}：未找到学生姓名，导出列将保留为空。")
        if not any(student.student_id for student in group.students):
            warnings.append(f"{label}：未找到学生学号，导出列将保留为空。")
    return list(dict.fromkeys(warnings))
