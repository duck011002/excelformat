from __future__ import annotations

import re

from .schemas import AcademicTerm


TERM_NAMES = {
    "1": "第一学期",
    "一": "第一学期",
    "第一": "第一学期",
    "2": "第二学期",
    "二": "第二学期",
    "第二": "第二学期",
    "3": "第三学期",
    "三": "第三学期",
    "第三": "第三学期",
}


def normalize_academic_term(value: object) -> AcademicTerm:
    text = re.sub(r"\s+", "", str(value or ""))
    if not text:
        return AcademicTerm()

    explicit = re.search(
        r"(?P<start>20\d{2})\s*[-/—至]\s*(?P<end>20\d{2})"
        r"(?:学年)?(?:第)?(?P<term>[一二三123]|第一|第二|第三)(?:学期)?",
        text,
    )
    if explicit:
        return _term(int(explicit.group("start")), int(explicit.group("end")), explicit.group("term"), 0.98, text)

    full_short = re.search(
        r"(?P<start>20\d{2})\s*[-/—]\s*(?P<end>20\d{2})"
        r"(?:学年)?\s*[-/—]?\s*(?:第)?(?P<term>[一二三123])(?:学期)?",
        text,
    )
    if full_short:
        return _term(
            int(full_short.group("start")),
            int(full_short.group("end")),
            full_short.group("term"),
            0.96,
            text,
        )

    short = re.search(
        r"(?<!\d)(?P<start>\d{2})\s*[-/—]\s*(?P<end>\d{2})"
        r"\s*[-/—]\s*(?P<term>[123])(?!\d)",
        text,
    )
    if short:
        start = _expand_two_digit_year(int(short.group("start")))
        end = _expand_two_digit_year(int(short.group("end")))
        return _term(start, end, short.group("term"), 0.94, text)

    autumn = re.search(r"(?<!\d)(20\d{2})\s*(?:年)?(?:秋|秋季)", text)
    if autumn:
        start = int(autumn.group(1))
        return _term(start, start + 1, "1", 0.92, text)

    spring = re.search(r"(?<!\d)(20\d{2})\s*(?:年)?(?:春|春季)", text)
    if spring:
        calendar_year = int(spring.group(1))
        return _term(calendar_year - 1, calendar_year, "2", 0.92, text)

    school_year = re.search(r"(?P<start>20\d{2})\s*[-/—至]\s*(?P<end>20\d{2})学年", text)
    term_only = re.search(r"(?:第)?(?P<term>[一二三123]|第一|第二|第三)学期", text)
    if school_year and term_only:
        return _term(
            int(school_year.group("start")),
            int(school_year.group("end")),
            term_only.group("term"),
            0.98,
            text,
        )
    return AcademicTerm()


def _term(start: int, end: int, raw_term: str, confidence: float, evidence: str) -> AcademicTerm:
    term = TERM_NAMES.get(raw_term, "")
    if not term or end < start or end - start > 2:
        return AcademicTerm()
    return AcademicTerm(
        school_year=f"{start:04d}-{end:04d}学年",
        term=term,
        confidence=confidence,
        evidence=evidence,
    )


def _expand_two_digit_year(value: int) -> int:
    return 2000 + value if value < 80 else 1900 + value
