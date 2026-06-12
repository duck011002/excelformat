from __future__ import annotations

import json
import http.client
import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path
from typing import Any


DEFAULT_PORT = 8501
PORT_SCAN_LIMIT = 20

CONFIG_ENV_MAP = {
    "base_url": "OPENAI_BASE_URL",
    "api_key": "OPENAI_API_KEY",
    "model": "OPENAI_MODEL",
    "preview_rows": "GRADE_CLEANER_PREVIEW_ROWS",
    "enable_repair": "GRADE_CLEANER_ENABLE_REPAIR",
}


def log_status(message: str) -> None:
    print(message, flush=True)


def bundled_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def external_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def apply_external_config() -> None:
    config_path = external_root() / "config" / "local_settings.json"
    if not config_path.exists():
        return

    try:
        values = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception as exc:  # pragma: no cover - console-only diagnostic
        print(f"Warning: failed to read {config_path}: {exc}")
        return

    if not isinstance(values, dict):
        print(f"Warning: ignored non-object config file: {config_path}")
        return

    for key, env_name in CONFIG_ENV_MAP.items():
        value = values.get(key)
        if value in (None, ""):
            continue
        if isinstance(value, bool):
            value = "true" if value else "false"
        os.environ.setdefault(env_name, str(value))


def choose_port(start: int = DEFAULT_PORT, attempts: int = PORT_SCAN_LIMIT) -> int:
    for port in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError(f"No available localhost port in {start}-{start + attempts - 1}")


def wait_for_http_ready(url: str, *, timeout: float = 90.0, poll_interval: float = 0.25) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            host_port = url.removeprefix("http://").split("/", 1)[0]
            host, port_text = host_port.split(":", 1)
            connection = http.client.HTTPConnection(host, int(port_text), timeout=2)
            try:
                connection.request("GET", "/_stcore/health")
                response = connection.getresponse()
                response.read()
            finally:
                connection.close()
            if response.status < 500:
                return True
        except OSError:
            pass
        time.sleep(poll_interval)
    return False


def open_browser_when_ready(
    url: str,
    *,
    timeout: float = 90.0,
    poll_interval: float = 0.25,
    opener=None,
    status_callback=None,
) -> threading.Thread | None:
    if os.getenv("GRADE_CLEANER_NO_BROWSER", "").strip().lower() in {"1", "true", "yes", "on"}:
        return None

    open_target = opener or webbrowser.open
    report = status_callback or log_status

    def runner() -> None:
        if wait_for_http_ready(url, timeout=timeout, poll_interval=poll_interval):
            report(f"本地服务已启动：{url}")
            open_target(url)
        else:
            report(f"等待服务启动超时：{url}")

    thread = threading.Thread(target=runner, daemon=True)
    thread.start()
    return thread


def main() -> int:
    apply_external_config()

    app_path = bundled_root() / "app.py"
    if not app_path.exists():
        print(f"Cannot find bundled Streamlit app: {app_path}")
        return 1

    port = choose_port()
    url = f"http://localhost:{port}"
    log_status("正在启动 成绩 Excel 智能清洗 v2.2 ...")
    log_status("正在加载运行环境，首次启动可能需要 10-20 秒。")
    log_status(f"服务准备完成后会自动打开浏览器：{url}")
    open_browser_when_ready(url)

    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    from streamlit import config
    from streamlit.web import bootstrap

    config.set_option("global.developmentMode", False)
    config.set_option("server.address", "localhost")
    config.set_option("server.port", port)
    config.set_option("server.headless", True)
    config.set_option("browser.serverAddress", "localhost")
    config.set_option("browser.serverPort", port)
    config.set_option("browser.gatherUsageStats", False)

    bootstrap.run(
        str(app_path),
        False,
        [],
        {},
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
