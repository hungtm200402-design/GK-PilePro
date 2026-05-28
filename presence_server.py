#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import sqlite3
import threading
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 8765
DEFAULT_DB = Path("presence_state.db")


DB_LOCK = threading.Lock()


def utc_now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_utc(value):
    s = str(value or "").strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def status_text(last_seen_at, timeout_seconds):
    dt = parse_utc(last_seen_at)
    if dt is None:
        return "Chưa có dữ liệu"
    delta = datetime.now(timezone.utc) - dt
    seconds = max(0, int(delta.total_seconds()))
    if seconds <= timeout_seconds:
        return "Đang hoạt động"
    minutes = seconds // 60
    if minutes < 60:
        return f"Truy cập {minutes} phút trước"
    hours = minutes // 60
    if hours < 24:
        return f"Truy cập {hours} giờ trước"
    days = hours // 24
    return f"Truy cập {days} ngày trước"


def init_db(db_path: Path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS machines (
                machine_code TEXT PRIMARY KEY,
                user_name TEXT DEFAULT '',
                role TEXT DEFAULT '',
                app_kind TEXT DEFAULT '',
                status TEXT DEFAULT '',
                first_seen_at TEXT DEFAULT '',
                last_seen_at TEXT DEFAULT '',
                payload_json TEXT DEFAULT ''
            )
            """
        )
        conn.commit()


def upsert_machine(db_path: Path, payload: dict):
    machine_code = str(payload.get("machine_code") or "").strip().upper()
    if not machine_code:
        return None
    now = utc_now()
    record = {
        "machine_code": machine_code,
        "user_name": str(payload.get("user_name") or "").strip(),
        "role": str(payload.get("role") or "").strip(),
        "app_kind": str(payload.get("app_kind") or "").strip(),
        "status": str(payload.get("status") or "online").strip(),
        "last_seen_at": now,
        "payload_json": json.dumps(payload, ensure_ascii=False),
    }
    with DB_LOCK, sqlite3.connect(db_path) as conn:
        cur = conn.execute(
            "SELECT first_seen_at FROM machines WHERE machine_code = ?",
            (machine_code,),
        )
        row = cur.fetchone()
        record["first_seen_at"] = row[0] if row and row[0] else now
        conn.execute(
            """
            INSERT INTO machines (
                machine_code, user_name, role, app_kind, status,
                first_seen_at, last_seen_at, payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(machine_code) DO UPDATE SET
                user_name=excluded.user_name,
                role=excluded.role,
                app_kind=excluded.app_kind,
                status=excluded.status,
                last_seen_at=excluded.last_seen_at,
                payload_json=excluded.payload_json
            """,
            (
                record["machine_code"],
                record["user_name"],
                record["role"],
                record["app_kind"],
                record["status"],
                record["first_seen_at"],
                record["last_seen_at"],
                record["payload_json"],
            ),
        )
        conn.commit()
    return record


def fetch_machines(db_path: Path):
    with DB_LOCK, sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT machine_code, user_name, role, app_kind, status, first_seen_at, last_seen_at, payload_json FROM machines ORDER BY last_seen_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


class PresenceHandler(BaseHTTPRequestHandler):
    server_version = "GKPresence/1.0"

    def _send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length > 0 else b"{}"
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    def log_message(self, format, *args):
        return

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        server = self.server  # type: ignore[attr-defined]
        timeout_seconds = getattr(server, "timeout_seconds", 20)
        path = urlparse(self.path).path.rstrip("/")
        if path in {"", "/"}:
            return self._send_json({"ok": True, "service": "presence"})
        if path == "/health":
            return self._send_json({"ok": True, "time": utc_now()})
        if path == "/machines":
            machines = fetch_machines(server.db_path)
            now = datetime.now(timezone.utc)
            for item in machines:
                dt = parse_utc(item.get("last_seen_at"))
                active = False
                if dt is not None:
                    active = (now - dt).total_seconds() <= timeout_seconds
                item["online"] = active
                item["status_text"] = "Đang hoạt động" if active else status_text(item.get("last_seen_at"), timeout_seconds)
            return self._send_json({"ok": True, "machines": machines, "timeout_seconds": timeout_seconds})
        if path.startswith("/machines/"):
            machine_code = path.split("/", 2)[2].strip().upper()
            machines = fetch_machines(server.db_path)
            for item in machines:
                if str(item.get("machine_code") or "").strip().upper() == machine_code:
                    dt = parse_utc(item.get("last_seen_at"))
                    active = False
                    if dt is not None:
                        active = (datetime.now(timezone.utc) - dt).total_seconds() <= timeout_seconds
                    item["online"] = active
                    item["status_text"] = "Đang hoạt động" if active else status_text(item.get("last_seen_at"), timeout_seconds)
                    return self._send_json({"ok": True, "machine": item})
            return self._send_json({"ok": False, "error": "not_found"}, status=404)
        return self._send_json({"ok": False, "error": "not_found"}, status=404)

    def do_POST(self):
        server = self.server  # type: ignore[attr-defined]
        path = urlparse(self.path).path.rstrip("/")
        if path == "/heartbeat":
            payload = self._read_json()
            record = upsert_machine(server.db_path, payload)
            if not record:
                return self._send_json({"ok": False, "error": "machine_code_required"}, status=400)
            return self._send_json({"ok": True, "machine": record})
        return self._send_json({"ok": False, "error": "not_found"}, status=404)


class PresenceHTTPServer(ThreadingHTTPServer):
    def __init__(self, server_address, RequestHandlerClass, db_path: Path, timeout_seconds: int = 20):
        super().__init__(server_address, RequestHandlerClass)
        self.db_path = db_path
        self.timeout_seconds = timeout_seconds


def main():
    parser = argparse.ArgumentParser(description="GK PilePro presence server")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--db", default=str(DEFAULT_DB))
    parser.add_argument("--timeout", type=int, default=20)
    args = parser.parse_args()

    db_path = Path(args.db).resolve()
    init_db(db_path)

    server = PresenceHTTPServer((args.host, args.port), PresenceHandler, db_path=db_path, timeout_seconds=max(5, int(args.timeout)))
    print(f"Presence server listening on http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
