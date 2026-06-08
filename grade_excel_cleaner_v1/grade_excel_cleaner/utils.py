from __future__ import annotations

import json
import math
import re
from typing import Any


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and math.isnan(value):
            return ""
    except TypeError:
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def normalize_header(value: Any) -> str:
    text = clean_cell(value)
    text = re.sub(r"\s+", "", text)
    return text


def make_unique_headers(values: list[Any]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}
    for index, value in enumerate(values):
        base = clean_cell(value) or f"未命名列{index + 1}"
        count = seen.get(base, 0)
        seen[base] = count + 1
        headers.append(base if count == 0 else f"{base}.{count + 1}")
    return headers


def strip_markdown_json(text: str) -> str:
    text = text.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.S | re.I)
    if fenced:
        return fenced.group(1).strip()
    return text


def extract_json_object(text: str) -> dict[str, Any]:
    text = strip_markdown_json(text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    if start < 0:
        raise ValueError("LLM response does not contain a JSON object")

    depth = 0
    in_string = False
    escape = False
    for position, char in enumerate(text[start:], start=start):
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start : position + 1])
    raise ValueError("LLM response JSON object is incomplete")
