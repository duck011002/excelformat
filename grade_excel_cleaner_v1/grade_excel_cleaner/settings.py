from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


DEFAULT_BASE_URL = "https://llm-service.polymas.com/api/openai/v1"
DEFAULT_MODEL = "gpt-5.4"
DEFAULT_PREVIEW_ROWS = 25
DEFAULT_ENABLE_REPAIR = True

PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOCAL_SETTINGS_PATH = PROJECT_ROOT / "config" / "local_settings.json"


@dataclass
class AppSettings:
    base_url: str = DEFAULT_BASE_URL
    api_key: str = ""
    model: str = DEFAULT_MODEL
    preview_rows: int = DEFAULT_PREVIEW_ROWS
    enable_repair: bool = DEFAULT_ENABLE_REPAIR

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def load_app_settings() -> AppSettings:
    values: dict[str, Any] = {}
    if LOCAL_SETTINGS_PATH.exists():
        values.update(json.loads(LOCAL_SETTINGS_PATH.read_text(encoding="utf-8")))

    env_map = {
        "base_url": ("OPENAI_BASE_URL", "LLM_BASE_URL"),
        "api_key": ("OPENAI_API_KEY",),
        "model": ("OPENAI_MODEL", "MODEL_NAME"),
        "preview_rows": ("GRADE_CLEANER_PREVIEW_ROWS",),
        "enable_repair": ("GRADE_CLEANER_ENABLE_REPAIR",),
    }
    for key, env_names in env_map.items():
        for env_name in env_names:
            if os.getenv(env_name):
                values[key] = os.environ[env_name]
                break

    settings = AppSettings(**{key: value for key, value in values.items() if key in AppSettings.__annotations__})
    settings.base_url = normalize_base_url(settings.base_url)
    settings.preview_rows = int(settings.preview_rows)
    settings.enable_repair = parse_bool(settings.enable_repair)
    return settings


def normalize_base_url(base_url: str) -> str:
    base_url = str(base_url or "").strip().rstrip("/")
    marker = "/https://"
    if marker in base_url:
        base_url = base_url.split(marker, 1)[0]
    return base_url


def parse_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() not in {"0", "false", "no", "off", ""}
