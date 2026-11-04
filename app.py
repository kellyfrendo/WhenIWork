"""Local page for checking a WhenIWork timesheet against a clock rule."""

from __future__ import annotations

import base64
import json
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from timesheet import TimesheetError, evaluate, parse_clock, read_timesheet, Rule

ROOT = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(ROOT, "static")
MAX_BYTES = 20_000_000


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        path = self.path.split("?", 1)[0]
        if path == "/":
            path = "/index.html"
        file_path = _static_path(path)
        if file_path is None or not os.path.isfile(file_path):
            self.send_error(404)
            return
        with open(file_path, "rb") as handle:
            data = handle.read()
        self.send_response(200)
        self.send_header("Content-Type", _mime(file_path))
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        if self.path.split("?", 1)[0] != "/api/check":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self._send_json(400, {"error": "The upload could not be read."})
            return
        if length <= 0 or length > MAX_BYTES:
            self._send_json(400, {"error": "Choose a timesheet smaller than 20 MB."})
            return
        try:
            payload = json.loads(self.rfile.read(length))
            result = _check(payload)
        except TimesheetError as error:
            self._send_json(400, {"error": str(error)})
            return
        except (json.JSONDecodeError, KeyError, TypeError, ValueError):
            self._send_json(400, {"error": "The upload could not be read."})
            return
        self._send_json(200, result)

    def log_message(self, fmt: str, *args) -> None:
        print("%s - %s" % (self.address_string(), fmt % args))

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def _check(payload: dict) -> dict:
    filename = str(payload.get("filename") or "")
    encoded = payload.get("file")
    if not isinstance(encoded, str) or not encoded:
        raise TimesheetError("Choose a timesheet file first.")
    try:
        data = base64.b64decode(encoded, validate=True)
    except ValueError as error:
        raise TimesheetError("The upload could not be read.") from error
    minimum = payload.get("minimumDays", 1)
    try:
        minimum_days = int(minimum)
    except (TypeError, ValueError) as error:
        raise TimesheetError("Enter a whole number of days.") from error
    rule = Rule(
        clock_in_before=parse_clock(payload.get("clockInBefore")),
        clock_out_after=parse_clock(payload.get("clockOutAfter")),
        minimum_days=minimum_days,
        schedule=str(payload.get("schedule") or "").strip(),
        exclude_note=str(payload.get("excludeNote") or "").strip(),
        period=str(payload.get("period") or "last-week").strip(),
    )
    punches = read_timesheet(data, filename)
    schedules = {punch["schedule"] for punch in punches}
    if rule.schedule and rule.schedule not in schedules:
        raise TimesheetError("That schedule is not in this timesheet.")
    return evaluate(punches, rule)


def _mime(file_path: str) -> str:
    if file_path.endswith(".css"):
        return "text/css; charset=utf-8"
    if file_path.endswith(".js"):
        return "text/javascript; charset=utf-8"
    if file_path.endswith(".html"):
        return "text/html; charset=utf-8"
    return "application/octet-stream"


def _static_path(url_path: str) -> str | None:
    relative = url_path.lstrip("/")
    if not relative or "\\" in relative:
        return None
    candidate = os.path.normpath(os.path.join(STATIC, relative))
    if os.path.commonpath([STATIC, candidate]) != STATIC:
        return None
    return candidate


def main() -> None:
    server = ThreadingHTTPServer(("127.0.0.1", 8765), Handler)
    print("Clock check is running at http://127.0.0.1:8765")
    print("Leave this window open, then open that address in your browser.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
