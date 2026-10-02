"""UTF-8 safety of the stdio path (issue #5): Hebrew text must survive
COM -> Python -> MCP transport even when the console code page can't
represent it."""

import io
import json
import os
import subprocess
import sys
import threading
from pathlib import Path

from outlook_mcp import server

REPO_ROOT = Path(__file__).resolve().parent.parent
HEBREW_SUBJECT = "מייל שיקוף"
HEBREW_BODY = "סיכום עשייה"

# Runs the real server.main() over stdio with a fake client that returns
# (and logs) Hebrew, echoing back the query it was given.
SERVER_SCRIPT = f"""
import logging
from outlook_mcp import server
from tests.conftest import FakeOutlookClient

class HebrewClient(FakeOutlookClient):
    def list_emails(self, query=None, **kwargs):
        logging.getLogger("outlook-mcp").info("listing %s", {HEBREW_SUBJECT!r})
        return [{{"id": "e1|s1", "subject": {HEBREW_SUBJECT!r},
                  "body": {HEBREW_BODY!r}, "query": query}}]

server.set_client(HebrewClient())
server.main()
"""


def _text_wrapper(encoding):
    return io.TextIOWrapper(io.BytesIO(), encoding=encoding)


def test_force_utf8_stdio_reconfigures_streams(monkeypatch):
    for name in ("stdin", "stdout", "stderr"):
        monkeypatch.setattr(sys, name, _text_wrapper("cp1252"))

    server._force_utf8_stdio()

    for name in ("stdin", "stdout", "stderr"):
        assert getattr(sys, name).encoding == "utf-8"
    assert sys.stderr.errors == "backslashreplace"
    sys.stdout.write(HEBREW_SUBJECT)  # would raise under cp1252


def test_force_utf8_stdio_tolerates_missing_streams(monkeypatch):
    # pythonw.exe / detached processes have no console streams at all.
    for name in ("stdin", "stdout", "stderr"):
        monkeypatch.setattr(sys, name, None)
    server._force_utf8_stdio()


def _send(proc, message):
    proc.stdin.write((json.dumps(message) + "\n").encode("utf-8"))
    proc.stdin.flush()


def _read_response(proc, request_id):
    while True:
        line = proc.stdout.readline()
        assert line, "server closed stdout before responding"
        message = json.loads(line.decode("utf-8"))  # strict UTF-8
        if message.get("id") == request_id:
            return message


def test_hebrew_round_trip_over_stdio_with_legacy_code_page():
    env = dict(os.environ)
    env.pop("PYTHONUTF8", None)
    # Simulate a Windows console code page that can't encode Hebrew.
    env["PYTHONIOENCODING"] = "cp1252"
    env["PYTHONPATH"] = os.pathsep.join(
        p for p in (str(REPO_ROOT), env.get("PYTHONPATH")) if p)

    proc = subprocess.Popen(
        [sys.executable, "-c", SERVER_SCRIPT], cwd=str(REPO_ROOT), env=env,
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    watchdog = threading.Timer(30, proc.kill)
    watchdog.start()
    try:
        _send(proc, {"jsonrpc": "2.0", "id": 1, "method": "initialize",
                     "params": {"protocolVersion": "2024-11-05",
                                "capabilities": {},
                                "clientInfo": {"name": "test",
                                               "version": "0"}}})
        _read_response(proc, 1)
        _send(proc, {"jsonrpc": "2.0",
                     "method": "notifications/initialized"})
        _send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/call",
                     "params": {"name": "list_emails",
                                "arguments": {"query": HEBREW_BODY}}})
        response = _read_response(proc, 2)
        _, stderr = proc.communicate(timeout=30)  # closes stdin -> exit
    finally:
        watchdog.cancel()
        if proc.poll() is None:
            proc.kill()
            proc.communicate()

    result = response["result"]
    assert not result.get("isError"), result
    email = json.loads(result["content"][0]["text"])
    assert email["subject"] == HEBREW_SUBJECT
    assert email["body"] == HEBREW_BODY
    assert email["query"] == HEBREW_BODY  # stdin decoded as UTF-8 too

    log = stderr.decode("utf-8")
    assert "Logging error" not in log, log
    assert HEBREW_SUBJECT in log
