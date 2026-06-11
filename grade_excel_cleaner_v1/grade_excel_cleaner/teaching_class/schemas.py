from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class AcademicTerm:
    school_year: str = ""
    term: str = ""
    confidence: float = 0.0
    evidence: str = ""


@dataclass
class EvidenceItem:
    text: str
    source_type: str
    sheet_name: str = ""
    row_index: int | None = None
    column_index: int | None = None
    score: float = 0.0


@dataclass
class WorkbookEvidence:
    filename: str
    items: list[EvidenceItem] = field(default_factory=list)
    expanded: bool = False


@dataclass
class ResolvedField:
    value: str = ""
    confidence: float = 0.0
    source_type: str = "unknown"
    evidence: list[str] = field(default_factory=list)
    candidates: list[str] = field(default_factory=list)
    warning: str = ""
    locked_by_user: bool = False


@dataclass
class StudentRecord:
    student_id: str = ""
    student_name: str = ""


@dataclass
class TeachingClassGroup:
    group_id: str
    course_name: str = ""
    grouping_source: str = "fallback"
    class_code: ResolvedField = field(default_factory=ResolvedField)
    class_name: ResolvedField = field(default_factory=ResolvedField)
    school_year: ResolvedField = field(default_factory=ResolvedField)
    term: ResolvedField = field(default_factory=ResolvedField)
    teacher_name: ResolvedField = field(default_factory=ResolvedField)
    teacher_id: ResolvedField = field(default_factory=ResolvedField)
    students: list[StudentRecord] = field(default_factory=list)
    revisions: list[dict[str, str]] = field(default_factory=list)


@dataclass
class TeachingClassResult:
    groups: list[TeachingClassGroup]
    evidence: WorkbookEvidence
    warnings: list[str] = field(default_factory=list)

