#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import hashlib
import json
import sqlite3
import threading
import hashlib
import os
import uuid
from contextlib import closing
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


SERVER_OWNER_MACHINE_CODE = os.getenv("GK_PILEPRO_SERVER_OWNER", "").strip().upper()


def get_machine_code():
    raw = f"{os.environ.get('COMPUTERNAME', '')}|{os.environ.get('USERNAME', '')}|{uuid.getnode()}"
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()
    return "-".join([digest[i:i + 5] for i in range(0, 20, 5)])


DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 8765
DEFAULT_DB = Path("presence_state.db")
USER_EXE_NAME = "GK PilePro.exe"
USER_UPDATE_INFO_CACHE = {}


DB_LOCK = threading.Lock()


def file_sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def cached_file_sha256(path: Path):
    try:
        resolved = Path(path).resolve()
        stat = resolved.stat()
        signature = (str(resolved), int(stat.st_size), int(stat.st_mtime))
        cached = USER_UPDATE_INFO_CACHE.get("sha256")
        if cached and cached.get("signature") == signature:
            return cached.get("sha") or ""
        sha = file_sha256(resolved)
        USER_UPDATE_INFO_CACHE["sha256"] = {"signature": signature, "sha": sha}
        return sha
    except Exception:
        return file_sha256(path)


def user_update_exe_path():
    base = Path(sys_executable_dir()).resolve()
    candidate = base / USER_EXE_NAME
    if candidate.exists():
        return candidate
    fallback = Path(__file__).resolve().parent / "dist" / USER_EXE_NAME
    return fallback if fallback.exists() else candidate


def sys_executable_dir():
    import sys

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def utc_now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def log_record_hash(record):
    payload = {
        "machine_code": str(record.get("machine_code") or ""),
        "user_name": str(record.get("user_name") or ""),
        "windows_user": str(record.get("windows_user") or ""),
        "computer_name": str(record.get("computer_name") or ""),
        "role": str(record.get("role") or ""),
        "app_kind": str(record.get("app_kind") or ""),
        "message": str(record.get("message") or ""),
        "log_text": str(record.get("log_text") or ""),
        "created_at": str(record.get("created_at") or ""),
    }
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


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
    with closing(sqlite3.connect(db_path)) as conn:
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
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS approved_machines (
                machine_code TEXT PRIMARY KEY,
                approval_code TEXT DEFAULT '',
                approval_version INTEGER DEFAULT 1,
                approved_at TEXT DEFAULT '',
                last_seen_at TEXT DEFAULT '',
                user_name TEXT DEFAULT '',
                role TEXT DEFAULT '',
                app_kind TEXT DEFAULT '',
                status TEXT DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS error_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                machine_code TEXT DEFAULT '',
                user_name TEXT DEFAULT '',
                windows_user TEXT DEFAULT '',
                computer_name TEXT DEFAULT '',
                role TEXT DEFAULT '',
                app_kind TEXT DEFAULT '',
                message TEXT DEFAULT '',
                log_text TEXT DEFAULT '',
                created_at TEXT DEFAULT '',
                resolved_at TEXT DEFAULT '',
                log_hash TEXT DEFAULT '',
                payload_json TEXT DEFAULT ''
            )
            """
        )
        cols = {row[1] for row in conn.execute("PRAGMA table_info(error_logs)").fetchall()}
        if "resolved_at" not in cols:
            conn.execute("ALTER TABLE error_logs ADD COLUMN resolved_at TEXT DEFAULT ''")
        if "log_hash" not in cols:
            conn.execute("ALTER TABLE error_logs ADD COLUMN log_hash TEXT DEFAULT ''")
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
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
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
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT machine_code, user_name, role, app_kind, status, first_seen_at, last_seen_at, payload_json FROM machines ORDER BY last_seen_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def upsert_approved_machine(db_path: Path, payload: dict):
    machine_code = str(payload.get("machine_code") or "").strip().upper()
    if not machine_code:
        return None
    now = str(payload.get("approved_at") or utc_now()).strip() or utc_now()
    record = {
        "machine_code": machine_code,
        "approval_code": str(payload.get("approval_code") or "").strip().upper(),
        "approval_version": int(payload.get("approval_version") or 1),
        "approved_at": now,
        "last_seen_at": str(payload.get("last_seen_at") or now).strip() or now,
        "user_name": str(payload.get("user_name") or "").strip(),
        "role": str(payload.get("role") or "").strip(),
        "app_kind": str(payload.get("app_kind") or "").strip(),
        "status": str(payload.get("status") or "approved").strip(),
    }
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        conn.execute(
            """
            INSERT INTO approved_machines (
                machine_code, approval_code, approval_version, approved_at,
                last_seen_at, user_name, role, app_kind, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(machine_code) DO UPDATE SET
                approval_code=excluded.approval_code,
                approval_version=excluded.approval_version,
                approved_at=excluded.approved_at,
                last_seen_at=excluded.last_seen_at,
                user_name=excluded.user_name,
                role=excluded.role,
                app_kind=excluded.app_kind,
                status=excluded.status
            """,
            (
                record["machine_code"],
                record["approval_code"],
                record["approval_version"],
                record["approved_at"],
                record["last_seen_at"],
                record["user_name"],
                record["role"],
                record["app_kind"],
                record["status"],
            ),
        )
        conn.commit()
    return record


def fetch_approved_machines(db_path: Path):
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT machine_code, approval_code, approval_version, approved_at, last_seen_at, user_name, role, app_kind, status FROM approved_machines ORDER BY approved_at DESC, last_seen_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def delete_approved_machine(db_path: Path, machine_code: str):
    machine_code = str(machine_code or "").strip().upper()
    if not machine_code:
        return False
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        cur = conn.execute("DELETE FROM approved_machines WHERE machine_code = ?", (machine_code,))
        conn.commit()
    return cur.rowcount > 0


def insert_error_log(db_path: Path, payload: dict):
    machine_code = str(payload.get("machine_code") or "").strip().upper()
    now = utc_now()
    record = {
        "machine_code": machine_code,
        "user_name": str(payload.get("user_name") or "").strip(),
        "windows_user": str(payload.get("windows_user") or "").strip(),
        "computer_name": str(payload.get("computer_name") or "").strip(),
        "role": str(payload.get("role") or "").strip(),
        "app_kind": str(payload.get("app_kind") or "").strip(),
        "message": str(payload.get("message") or "").strip(),
        "log_text": str(payload.get("log_text") or "").strip()[-60000:],
        "created_at": now,
        "resolved_at": "",
        "payload_json": json.dumps(payload, ensure_ascii=False),
    }
    record["log_hash"] = log_record_hash(record)
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        cur = conn.execute(
            """
            INSERT INTO error_logs (
                machine_code, user_name, windows_user, computer_name,
                role, app_kind, message, log_text, created_at, resolved_at, log_hash, payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record["machine_code"],
                record["user_name"],
                record["windows_user"],
                record["computer_name"],
                record["role"],
                record["app_kind"],
                record["message"],
                record["log_text"],
                record["created_at"],
                record["resolved_at"],
                record["log_hash"],
                record["payload_json"],
            ),
        )
        conn.commit()
        record["id"] = cur.lastrowid
    return record


def fetch_error_logs(db_path: Path, limit=100, unresolved_only=False):
    try:
        limit = max(1, min(300, int(limit or 100)))
    except Exception:
        limit = 100
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        conn.row_factory = sqlite3.Row
        where_sql = "WHERE COALESCE(resolved_at, '') = ''" if unresolved_only else ""
        rows = conn.execute(
            f"""
            SELECT id, machine_code, user_name, windows_user, computer_name,
                   role, app_kind, message, log_text, created_at, resolved_at, log_hash
            FROM error_logs
            {where_sql}
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    out = []
    for r in rows:
        item = dict(r)
        item["hash_ok"] = bool(item.get("log_hash")) and item.get("log_hash") == log_record_hash(item)
        out.append(item)
    return out


def resolve_error_log(db_path: Path, log_id):
    try:
        log_id = int(log_id)
    except Exception:
        return False
    with DB_LOCK, closing(sqlite3.connect(db_path)) as conn:
        cur = conn.execute(
            "UPDATE error_logs SET resolved_at = ? WHERE id = ?",
            (utc_now(), log_id),
        )
        conn.commit()
    return cur.rowcount > 0


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

    def _send_file(self, path: Path):
        if not path.exists() or not path.is_file():
            return self._send_json({"ok": False, "error": "not_found"}, status=404)
        try:
            size = path.stat().st_size
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
            self.send_header("Content-Length", str(size))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            with path.open("rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    self.wfile.write(chunk)
        except Exception:
            return

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
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        server = self.server  # type: ignore[attr-defined]
        timeout_seconds = getattr(server, "timeout_seconds", 20)
        parsed_url = urlparse(self.path)
        path = parsed_url.path.rstrip("/")
        if path in {"", "/"}:
            return self._send_json({"ok": True, "service": "presence"})
        if path == "/health":
            return self._send_json({"ok": True, "time": utc_now()})
        if path == "/update-info":
            exe_path = user_update_exe_path()
            if not exe_path.exists():
                return self._send_json({"ok": False, "error": "user_exe_not_found"}, status=404)
            stat = exe_path.stat()
            return self._send_json(
                {
                    "ok": True,
                    "app": "GK PilePro",
                    "filename": USER_EXE_NAME,
                    "size": stat.st_size,
                    "mtime": int(stat.st_mtime),
                    "sha256": cached_file_sha256(exe_path),
                    "download_url": "/updates/user-exe",
                }
            )
        if path == "/updates/user-exe":
            return self._send_file(user_update_exe_path())
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
        if path == "/approved-machines":
            approved = fetch_approved_machines(server.db_path)
            return self._send_json({"ok": True, "approved_machines": approved})
        if path == "/error-logs":
            query = parse_qs(parsed_url.query or "")
            limit = (query.get("limit") or ["100"])[0]
            unresolved_only = str((query.get("unresolved") or [""])[0]).strip().lower() in {"1", "true", "yes"}
            logs = fetch_error_logs(server.db_path, limit=limit, unresolved_only=unresolved_only)
            return self._send_json({"ok": True, "error_logs": logs})
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
        if path == "/approved-machines":
            payload = self._read_json()
            record = upsert_approved_machine(server.db_path, payload)
            if not record:
                return self._send_json({"ok": False, "error": "machine_code_required"}, status=400)
            return self._send_json({"ok": True, "machine": record})
        if path == "/error-logs":
            payload = self._read_json()
            record = insert_error_log(server.db_path, payload)
            return self._send_json({"ok": True, "error_log": record})
        if path.startswith("/error-logs/") and path.endswith("/resolve"):
            parts = [p for p in path.split("/") if p]
            log_id = parts[1] if len(parts) >= 3 else ""
            if not resolve_error_log(server.db_path, log_id):
                return self._send_json({"ok": False, "error": "not_found"}, status=404)
            return self._send_json({"ok": True, "id": log_id})
        if path.startswith("/approved-machines/"):
            machine_code = path.split("/", 2)[2].strip().upper()
            deleted = delete_approved_machine(server.db_path, machine_code)
            if not deleted:
                return self._send_json({"ok": False, "error": "not_found"}, status=404)
            return self._send_json({"ok": True, "machine_code": machine_code})
        return self._send_json({"ok": False, "error": "not_found"}, status=404)

    def do_DELETE(self):
        server = self.server  # type: ignore[attr-defined]
        path = urlparse(self.path).path.rstrip("/")
        if path.startswith("/approved-machines/"):
            machine_code = path.split("/", 2)[2].strip().upper()
            deleted = delete_approved_machine(server.db_path, machine_code)
            if not deleted:
                return self._send_json({"ok": False, "error": "not_found"}, status=404)
            return self._send_json({"ok": True, "machine_code": machine_code})
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

    if SERVER_OWNER_MACHINE_CODE and get_machine_code() != SERVER_OWNER_MACHINE_CODE:
        raise SystemExit("This presence server is locked to the owner machine.")

    db_path = Path(args.db).resolve()
    init_db(db_path)

    host = str(args.host or DEFAULT_HOST).strip()

    server = PresenceHTTPServer((host, args.port), PresenceHandler, db_path=db_path, timeout_seconds=max(5, int(args.timeout)))
    print(f"Presence server listening on http://{host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
