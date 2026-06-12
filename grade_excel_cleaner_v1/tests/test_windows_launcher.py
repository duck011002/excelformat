from __future__ import annotations

import socket
import sys
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import windows_launcher


class _OkHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def log_message(self, format: str, *args: object) -> None:  # noqa: A003
        return


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


class BrowserOpenTests(unittest.TestCase):
    def test_opens_browser_only_after_server_is_ready(self) -> None:
        port = _free_port()
        url = f"http://127.0.0.1:{port}"
        opened: list[float] = []
        statuses: list[str] = []
        server_started = threading.Event()

        def start_server_later() -> None:
            time.sleep(0.25)
            server = ThreadingHTTPServer(("127.0.0.1", port), _OkHandler)
            server_started.set()
            try:
                server.handle_request()
            finally:
                server.server_close()

        thread = threading.Thread(target=start_server_later, daemon=True)
        thread.start()

        browser_thread = windows_launcher.open_browser_when_ready(
            url,
            timeout=2.0,
            poll_interval=0.05,
            opener=lambda target: opened.append(time.monotonic()),
            status_callback=statuses.append,
        )

        browser_thread.join(timeout=3.0)

        self.assertTrue(server_started.is_set())
        self.assertEqual(len(opened), 1)
        self.assertTrue(any("http://127.0.0.1:" in status for status in statuses))


if __name__ == "__main__":
    unittest.main()
