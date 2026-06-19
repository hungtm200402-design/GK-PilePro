# -*- coding: utf-8 -*-

import hashlib
import json
import msvcrt
import os
import re
import shutil
import socket
import sys
import time
import traceback
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib import error as urllib_error, parse as urllib_parse, request as urllib_request

from dotenv import load_dotenv


APP_TITLE = "GK PilePro"
SERVER_OWNER_MACHINE_CODE = os.getenv("GK_PILEPRO_SERVER_OWNER", "").strip().upper()


DEFAULT_MODEL = "gemini-3.1-flash-lite"

FALLBACK_MODELS = ["gemini-3.1-flash-lite", "gemini-2.5-flash-lite", "gemini-2.5-flash"]

APPROVAL_SECRET = "GK_PILEPRO_APPROVAL_V1"



def app_dir():

    if getattr(sys, "frozen", False):

        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent



def _safe_path_part(value, fallback="default"):

    text = str(value or "").strip()

    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)

    text = text.strip("._-")

    return text or fallback



def app_data_root():

    base = os.environ.get("LOCALAPPDATA")

    if base:

        return Path(base) / APP_TITLE

    return Path.home() / "AppData" / "Local" / APP_TITLE



def app_data_dir():

    # Keep user data outside the install folder so every machine has its own
    # recent files, history, approvals, settings and OCR logs.

    machine_part = _safe_path_part(get_machine_code(), "unknown_machine")

    path = app_data_root() / machine_part

    path.mkdir(parents=True, exist_ok=True)

    return path



def app_data_path(*parts):

    path = app_data_dir().joinpath(*parts)

    path.parent.mkdir(parents=True, exist_ok=True)

    return path


def role_log_path():
    return app_data_path("logs", "admin_error.log" if is_admin_build() else "user_error.log")


def write_role_error_log(context, exc=None, extra=None):
    try:
        path = role_log_path()
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            "=" * 80,
            f"time: {stamp}",
            f"role: {'admin' if is_admin_build() else 'user'}",
            f"machine: {get_machine_code()}",
            f"context: {context}",
        ]
        if extra:
            try:
                lines.append("extra: " + json.dumps(extra, ensure_ascii=False, default=str))
            except Exception:
                lines.append(f"extra: {extra}")
        if exc is not None:
            lines.append(f"error: {exc}")
            lines.append(traceback.format_exc())
        with path.open("a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        return path
    except Exception:
        return None


def clean_role_log_text_for_report(log_text):
    text = str(log_text or "")
    empty_marker = "Chưa có lỗi nào được ghi nhận."
    separator = "=" * 80
    if separator in text:
        text = text[text.find(separator):].strip()
    return text or empty_marker


def report_runtime_error_to_admin(context, exc=None, extra=None, error_file_name=None, notify_server=True):
    """Record a real exception traceback and best-effort send it to the admin server."""
    trace_text = traceback.format_exc()
    if trace_text.strip() == "NoneType: None" and exc is not None:
        trace_text = "".join(traceback.format_exception(type(exc), exc, getattr(exc, "__traceback__", None)))

    out_path = None
    try:
        if error_file_name:
            out_dir = last_run_dir()
            out_dir.mkdir(exist_ok=True)
            out_path = out_dir / error_file_name
            out_path.write_text(trace_text, encoding="utf-8")
    except Exception:
        out_path = None

    log_path = write_role_error_log(context, exc, {
        **(extra or {}),
        "error_file": str(out_path or ""),
    })

    if notify_server:
        try:
            server_url = presence_server_url_from_env()
            if check_presence_server_alive(server_url, timeout=0.7):
                log_text = ""
                try:
                    if log_path and Path(log_path).exists():
                        log_text = Path(log_path).read_text(encoding="utf-8", errors="replace")
                except Exception:
                    log_text = ""
                log_text = clean_role_log_text_for_report(log_text)
                payload = {
                    "machine_code": get_machine_code(),
                    "user_name": resolve_member_display_name(get_machine_code()),
                    "windows_user": os.environ.get("USERNAME", ""),
                    "computer_name": os.environ.get("COMPUTERNAME", socket.gethostname()),
                    "role": "Admin" if is_admin_build() else "User",
                    "app_kind": "admin" if is_admin_build() else "user",
                    "message": f"Lỗi {context}",
                    "log_text": (log_text or trace_text)[-60000:],
                    "log_path": str(log_path or ""),
                    "error_file": str(out_path or ""),
                    "extra": extra or {},
                }
                send_presence_error_log(server_url, payload, timeout=3)
        except Exception:
            pass

    return out_path, log_path


def audit_log_path():
    return app_data_path("logs", "audit_events.jsonl")


def _audit_hash_payload(record):
    payload = dict(record or {})
    payload.pop("hash", None)
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)


def make_record_hash(record):
    return hashlib.sha256(_audit_hash_payload(record).encode("utf-8")).hexdigest()


def append_audit_event(action, status="success", file_path="", message="", extra=None):
    try:
        record = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "role": "admin" if is_admin_build() else "user",
            "machine": get_machine_code(),
            "windows_user": current_os_username(),
            "app_user": load_app_user_name_setting(),
            "action": str(action or "").strip(),
            "status": str(status or "").strip() or "success",
            "file_path": str(file_path or ""),
            "message": str(message or ""),
            "extra": extra or {},
        }
        record["hash"] = make_record_hash(record)
        path = audit_log_path()
        with path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")
        return record
    except Exception as exc:
        write_role_error_log("append_audit_event", exc, {"action": action, "file_path": file_path})
        return None


def verify_record_hash(record):
    try:
        return str((record or {}).get("hash") or "") == make_record_hash(record)
    except Exception:
        return False


@contextmanager
def locked_file_for_write(path, timeout=10):
    lock_path = Path(str(path) + ".gklock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    lock_file = lock_path.open("a+b")
    start = time.time()
    locked = False
    try:
        while True:
            try:
                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                locked = True
                break
            except OSError:
                if time.time() - start >= float(timeout or 10):
                    raise TimeoutError(f"File đang được máy khác ghi: {path}")
                time.sleep(0.15)
        yield lock_path
    finally:
        if locked:
            try:
                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
            except Exception:
                pass
        try:
            lock_file.close()
        except Exception:
            pass


def list_backup_files(category="excel", limit=200):
    try:
        root = app_data_path("backups", category)
        if not root.exists():
            return []
        rows = []
        for path in root.glob("*"):
            if not path.is_file():
                continue
            try:
                stat = path.stat()
                rows.append(
                    {
                        "path": str(path),
                        "name": path.name,
                        "size": stat.st_size,
                        "modified_at": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
            except Exception:
                pass
        rows.sort(key=lambda item: item.get("modified_at", ""), reverse=True)
        return rows[: max(1, int(limit or 200))]
    except Exception as exc:
        write_role_error_log("list_backup_files", exc, {"category": category})
        return []


def restore_backup_file(backup_path, target_path):
    src = Path(backup_path)
    dst = Path(target_path)
    if not src.exists():
        raise FileNotFoundError(f"Không tìm thấy backup: {src}")
    if not dst.parent.exists():
        dst.parent.mkdir(parents=True, exist_ok=True)
    pre_restore_backup = backup_file(dst, "excel_restore_before")
    shutil.copy2(src, dst)
    append_audit_event(
        "restore_excel_backup",
        file_path=str(dst),
        message="Admin khôi phục Excel từ backup.",
        extra={"backup_path": str(src), "pre_restore_backup": str(pre_restore_backup or "")},
    )
    return str(pre_restore_backup or "")


def backup_file(path, category="config"):
    try:
        src = Path(path)
        if not src.exists():
            return None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = app_data_path("backups", category)
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / f"{src.stem}_{stamp}{src.suffix}"
        shutil.copy2(src, out)
        return out
    except Exception as exc:
        write_role_error_log(f"backup_file:{path}", exc)
        return None


def validate_excel_before_write(wb):
    if wb is None:
        raise ValueError("Không đọc được workbook Excel.")
    if not getattr(wb, "sheetnames", None):
        raise ValueError("Workbook không có sheet.")
    for ws in wb.worksheets:
        if ws.max_row >= 1 and ws.max_column >= 1:
            return True
    raise ValueError("Workbook không có dữ liệu để ghi.")


def last_run_dir():

    path = app_data_path("last_run_v12")

    path.mkdir(parents=True, exist_ok=True)

    return path



def resource_dir():

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):

        return Path(sys._MEIPASS)

    return app_dir()



def resource_path(*parts):

    return resource_dir().joinpath(*parts)



def current_user_role_labels():

    if getattr(sys, "frozen", False):

        exe_name = Path(sys.executable).stem.lower()

        if "admin" in exe_name:

            return "Admin", "Quản trị viên"

    return "Thành viên", "Người dùng"



def current_os_username():

    username = str(os.environ.get("USERNAME") or "").strip()

    if username:

        return username

    try:

        return os.getlogin()

    except Exception:

        return "Unknown"



def default_app_user_name():

    if getattr(sys, "frozen", False):

        return "Admin" if is_admin_build() else "User"

    return current_os_username()



def load_app_user_name_setting():

    load_dotenv(env_path())

    name = str(os.getenv("APP_USER_NAME") or "").strip()

    if name:

        return name

    try:

        settings = json.loads(user_settings_path().read_text(encoding="utf-8"))

        name = str(settings.get("APP_USER_NAME") or "").strip()

        if name:

            return name

    except Exception:

        pass

    return default_app_user_name()



def current_app_user_name():

    return default_app_user_name()



def lookup_assigned_machine_user_name(machine_code):

    machine_code = str(machine_code or "").strip().upper()

    if not machine_code:

        return ""

    try:

        for item in load_admin_approved_machines():

            if str(item.get("machine_code", "")).strip().upper() == machine_code:

                return str(item.get("user_name") or "").strip()

    except Exception:

        pass

    return ""



def resolve_member_display_name(machine_code=None):

    machine_code = str(machine_code or get_machine_code() or "").strip().upper()

    name = lookup_assigned_machine_user_name(machine_code)

    if name:

        return name

    try:

        data = json.loads(approval_path().read_text(encoding="utf-8"))

        if str(data.get("machine_code", "")).strip().upper() == machine_code:

            name = str(data.get("user_name") or "").strip()

            if name:

                return name

    except Exception:

        pass

    return "Người dùng"



def is_admin_build():

    if not getattr(sys, "frozen", False):

        return False

    return "admin" in Path(sys.executable).stem.lower()



def approval_path():

    return app_data_path("gk_pilepro_approval.json")



def admin_approved_machines_path():

    return app_data_path("gk_pilepro_approved_machines.json")



def code_versions_path():

    return app_data_path("gk_pilepro_code_versions.json")



def legacy_revoked_machines_path():

    return app_data_path("gk_pilepro_revoked_machines.json")



def get_machine_code():

    raw = f"{os.environ.get('COMPUTERNAME', '')}|{os.environ.get('USERNAME', '')}|{uuid.getnode()}"

    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()

    return "-".join([digest[i:i + 5] for i in range(0, 20, 5)])


def is_server_owner_machine():
    return (not SERVER_OWNER_MACHINE_CODE) or get_machine_code() == SERVER_OWNER_MACHINE_CODE



def make_approval_code(machine_code, version=None):

    clean = re.sub(r"[^A-Z0-9]", "", str(machine_code or "").upper())

    if version is None:

        version = machine_approval_version(machine_code)

    if int(version or 1) <= 1:

        digest = hashlib.sha256(f"{APPROVAL_SECRET}:{clean}".encode("utf-8")).hexdigest().upper()

    else:

        digest = hashlib.sha256(f"{APPROVAL_SECRET}:{clean}:V{int(version or 1)}".encode("utf-8")).hexdigest().upper()

    return "-".join([digest[i:i + 4] for i in range(0, 16, 4)])



def load_revoked_machines():

    try:

        data = json.loads(code_versions_path().read_text(encoding="utf-8"))

        if isinstance(data, list):

            return data

    except Exception:

        pass

    try:

        data = json.loads(legacy_revoked_machines_path().read_text(encoding="utf-8"))

        if isinstance(data, list):

            return data

    except Exception:

        pass

    return []



def save_revoked_machines(items):

    code_versions_path().write_text(

        json.dumps(items or [], ensure_ascii=False, indent=2),

        encoding="utf-8",

    )



def machine_approval_version(machine_code):

    machine_code = str(machine_code or "").strip().upper()

    if not machine_code:

        return 1

    for item in load_revoked_machines():

        if str(item.get("machine_code", "")).strip().upper() == machine_code:

            try:

                return max(1, int(item.get("approval_version") or item.get("version") or 2))

            except Exception:

                return 2

    return 1



def invalidate_machine_approval_code(machine_code):

    machine_code = str(machine_code or "").strip().upper()

    if not machine_code:

        return 1

    items = load_revoked_machines()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for item in items:

        if str(item.get("machine_code", "")).strip().upper() == machine_code:

            try:

                next_version = max(2, int(item.get("approval_version") or item.get("version") or 1) + 1)

            except Exception:

                next_version = 2

            item["approval_version"] = next_version

            item["invalidated_at"] = now

            item.pop("approval_code", None)

            item.pop("revoked_at", None)

            save_revoked_machines(items)

            return next_version

    items.append({

        "machine_code": machine_code,

        "approval_version": 2,

        "invalidated_at": now,

    })

    save_revoked_machines(items)

    return 2



def is_machine_approved():

    machine_code = get_machine_code()
    current_version = machine_approval_version(machine_code)

    try:
        server_url = presence_server_url_from_env()
        if not server_url or not check_presence_server_alive(server_url, timeout=0.5):
            return False
        approved_items = fetch_presence_approved_machines(server_url, timeout=3)
        for item in approved_items:
            if not isinstance(item, dict):
                continue
            if str(item.get("machine_code") or "").strip().upper() != machine_code:
                continue
            saved_version = int(item.get("approval_version") or 1)
            saved_code = str(item.get("approval_code") or "").strip().upper()
            return (
                saved_version >= current_version
                and saved_code == make_approval_code(machine_code, saved_version)
            )
    except Exception:
        return False

    return False



def save_machine_approval(approval_code):

    machine_code = get_machine_code()

    current_version = machine_approval_version(machine_code)

    approval_code = str(approval_code or "").strip().upper()

    matched_version = None

    for version in [current_version] + [v for v in range(1, 51) if v != current_version]:

        if approval_code == make_approval_code(machine_code, version):

            matched_version = version

            break

    if matched_version is None:

        return False

    assigned_name = lookup_assigned_machine_user_name(machine_code) or "Người dùng"

    approval_path().write_text(

        json.dumps(

            {

                "machine_code": machine_code,

                "approval_code": make_approval_code(machine_code, matched_version),

                "approval_version": matched_version,

                "user_name": assigned_name,

            },

            ensure_ascii=False,

            indent=2,

        ),

        encoding="utf-8",

    )

    try:

        remember_admin_approved_machine(machine_code, assigned_name)

    except Exception:

        pass

    return True



def load_admin_approved_machines():

    local_items = []
    try:
        data = json.loads(admin_approved_machines_path().read_text(encoding="utf-8"))
        if isinstance(data, list):
            local_items = [item for item in data if isinstance(item, dict)]
    except Exception:
        local_items = []

    merged = {}
    for item in local_items:
        code = str(item.get("machine_code") or "").strip().upper()
        if code:
            merged[code] = dict(item)

    try:
        server_url = presence_server_url_from_env()
        remote_items = fetch_presence_approved_machines(server_url, timeout=3)
        for item in remote_items:
            if not isinstance(item, dict):
                continue
            code = str(item.get("machine_code") or "").strip().upper()
            if code:
                merged[code] = dict(item)
    except Exception:
        pass

    items = list(merged.values())
    try:
        items.sort(
            key=lambda x: (
                str(x.get("approved_at") or ""),
                str(x.get("last_seen_at") or ""),
                str(x.get("machine_code") or ""),
            ),
            reverse=True,
        )
    except Exception:
        pass
    return items



def save_admin_approved_machines(items):

    admin_approved_machines_path().write_text(

        json.dumps(items or [], ensure_ascii=False, indent=2),

        encoding="utf-8",

    )



def remember_admin_approved_machine(machine_code, user_name=None):

    machine_code = str(machine_code or "").strip().upper()

    if not machine_code:

        return None

    approval_version = machine_approval_version(machine_code)

    approval_code = make_approval_code(machine_code, approval_version)

    user_name = str(user_name or "").strip()

    items = load_admin_approved_machines()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    found = False

    for item in items:

        if item.get("machine_code") == machine_code:

            item["approval_code"] = approval_code

            item["approval_version"] = approval_version

            item["approved_at"] = now

            item["last_seen_at"] = item.get("last_seen_at") or now

            if user_name:

                item["user_name"] = user_name

            found = True

            break

    if not found:

        items.append({

            "machine_code": machine_code,

            "approval_code": approval_code,

            "approval_version": approval_version,

            "approved_at": now,

            "last_seen_at": now,

            "user_name": user_name,

        })

    save_admin_approved_machines(items)

    try:
        server_url = presence_server_url_from_env()
        if server_url:
            _payload = {
                "machine_code": machine_code,
                "approval_code": approval_code,
                "approval_version": approval_version,
                "approved_at": now,
                "last_seen_at": now,
                "user_name": user_name,
                "role": "",
                "app_kind": "",
                "status": "approved",
            }
            threading.Thread(target=lambda: send_presence_approved_machine(server_url, _payload, timeout=3), daemon=True).start()
    except Exception:
        pass

    return approval_code


def update_admin_machine_last_seen(machine_code, seen_at=None):
    machine_code = str(machine_code or "").strip().upper()
    if not machine_code:
        return False
    items = load_admin_approved_machines()
    now = seen_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    changed = False
    for item in items:
        if str(item.get("machine_code", "")).strip().upper() == machine_code:
            item["last_seen_at"] = now
            changed = True
            break
    if changed:
        save_admin_approved_machines(items)
    return changed


def parse_machine_datetime(text):
    s = str(text or "").strip()
    if not s:
        return None
    iso_candidate = s[:-1] + "+00:00" if s.endswith("Z") else s
    try:
        dt = datetime.fromisoformat(iso_candidate)
        if dt.tzinfo is not None:
            return dt.astimezone().replace(tzinfo=None)
        return dt
    except Exception:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def format_machine_last_seen(text, now=None):
    dt = parse_machine_datetime(text)
    if dt is None:
        return "Chưa có dữ liệu"
    now = now or datetime.now()
    delta = now - dt
    seconds = max(0, int(delta.total_seconds()))
    if seconds < 60:
        return "Đang hoạt động"
    minutes = seconds // 60
    if minutes < 60:
        return f"Truy cập {minutes} phút trước"
    hours = minutes // 60
    if hours < 24:
        return f"Truy cập {hours} giờ trước"
    days = hours // 24
    return f"Truy cập {days} ngày trước"



def is_machine_active_recently(text, now=None):
    dt = parse_machine_datetime(text)
    if dt is None:
        return False
    now = now or datetime.now()
    return (now - dt).total_seconds() < 60



def delete_admin_approved_machine(machine_code):

    machine_code = str(machine_code or "").strip().upper()

    invalidate_machine_approval_code(machine_code)

    items = [x for x in load_admin_approved_machines() if x.get("machine_code") != machine_code]

    save_admin_approved_machines(items)

    try:
        server_url = presence_server_url_from_env()
        if server_url:
            delete_presence_approved_machine(server_url, machine_code, timeout=3)
    except Exception:
        pass

    try:

        data = json.loads(approval_path().read_text(encoding="utf-8"))

        if str(data.get("machine_code", "")).strip().upper() == machine_code:

            approval_path().unlink(missing_ok=True)

    except Exception:

        pass

    return items



def import_local_approval_to_admin_list():

    try:

        data = json.loads(approval_path().read_text(encoding="utf-8"))

        machine_code = data.get("machine_code", "")

        approval_code = data.get("approval_code", "")

        approval_version = int(data.get("approval_version") or 1)

        user_name = str(data.get("user_name", "") or lookup_assigned_machine_user_name(machine_code) or "").strip()

        if (

            machine_code

            and approval_version == machine_approval_version(machine_code)

            and approval_code == make_approval_code(machine_code, approval_version)

        ):

            remember_admin_approved_machine(machine_code, user_name)

    except Exception:

        pass





def sync_presence_machines_to_admin_list(presence_rows):

    """
    Đồng bộ toàn bộ máy đang heartbeat từ presence server vào danh sách máy đã duyệt.
    """

    try:

        rows = [r for r in (presence_rows or []) if isinstance(r, dict)]

        if not rows:

            return 0

        items = load_admin_approved_machines()

        index_by_code = {}

        for idx, item in enumerate(items):

            code = str(item.get("machine_code") or "").strip().upper()

            if code:

                index_by_code[code] = idx

        changed = False

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for row in rows:

            machine_code = str(row.get("machine_code") or "").strip().upper()

            if not machine_code:

                continue

            user_name = str(row.get("user_name") or "").strip()

            if not user_name:

                user_name = str(lookup_assigned_machine_user_name(machine_code) or "").strip()

            last_seen_at = str(row.get("last_seen_at") or "").strip() or now

            app_kind = str(row.get("app_kind") or "").strip()

            role = str(row.get("role") or "").strip()

            status = str(row.get("status") or "").strip() or "online"

            if machine_code not in index_by_code:
                continue

            item = items[index_by_code[machine_code]]

            if user_name and str(item.get("user_name") or "").strip() != user_name:

                item["user_name"] = user_name

                changed = True

            if last_seen_at and str(item.get("last_seen_at") or "").strip() != last_seen_at:

                item["last_seen_at"] = last_seen_at

                changed = True

            if app_kind and str(item.get("app_kind") or "").strip() != app_kind:

                item["app_kind"] = app_kind

                changed = True

            if role and str(item.get("role") or "").strip() != role:

                item["role"] = role

                changed = True

            if status and str(item.get("status") or "").strip() != status:

                item["status"] = status

                changed = True

            if not str(item.get("approval_code") or "").strip():

                approval_version = int(item.get("approval_version") or machine_approval_version(machine_code) or 1)

                item["approval_code"] = make_approval_code(machine_code, approval_version)

                item["approval_version"] = approval_version

                changed = True

        if changed:

            save_admin_approved_machines(items)

        return len(rows)

    except Exception:

        return 0


def env_path():

    external = app_data_path(".env")

    if external.exists():

        return external

    bundled = resource_dir() / ".env"

    if bundled.exists():

        return bundled

    return external



def user_settings_path():

    return app_data_path("tool_kl_settings.json")



def selected_excel_files_path():

    return app_data_path("tool_kl_selected_excels.json")



def history_entries_path():

    return app_data_path("tool_kl_history.json")


def mapping_templates_path():

    return app_data_path("tool_kl_mapping_templates.json")


def formula_profiles_path():

    return app_data_path("tool_kl_formula_profiles.json")


def load_history_entries():

    try:

        data = json.loads(history_entries_path().read_text(encoding="utf-8"))

        if isinstance(data, list):

            out = []

            for item in data:

                if isinstance(item, dict):

                    out.append(item)

            return out

    except Exception:

        pass

    return []


def save_history_entries(entries):

    try:
        backup_file(history_entries_path(), "config")

        history_entries_path().write_text(

            json.dumps(entries or [], ensure_ascii=False, indent=2),

            encoding="utf-8",

        )

    except Exception:

        pass


def load_mapping_templates():

    try:

        data = json.loads(mapping_templates_path().read_text(encoding="utf-8"))

        if isinstance(data, list):

            return [item for item in data if isinstance(item, dict)]

    except Exception:

        pass

    return []


def save_mapping_templates(templates):

    try:
        backup_file(mapping_templates_path(), "config")

        mapping_templates_path().write_text(

            json.dumps(templates or [], ensure_ascii=False, indent=2),

            encoding="utf-8",

        )

    except Exception:

        pass


def load_formula_profiles():

    try:

        data = json.loads(formula_profiles_path().read_text(encoding="utf-8"))

        if isinstance(data, list):

            return [item for item in data if isinstance(item, dict)]

    except Exception:

        pass

    return []


def save_formula_profiles(profiles):

    try:
        backup_file(formula_profiles_path(), "config")

        formula_profiles_path().write_text(

            json.dumps(profiles or [], ensure_ascii=False, indent=2),

            encoding="utf-8",

        )

    except Exception:

        pass


def append_history_entry(entry):

    try:

        entries = load_history_entries()

        entries.append(entry)

        if len(entries) > 1000:

            entries = entries[-1000:]

        save_history_entries(entries)

    except Exception:

        pass


def new_workflow_id():

    return uuid.uuid4().hex[:12].upper()


def load_env_values():

    load_dotenv(env_path())

    values = {

        "GEMINI_API_KEY": os.getenv("GEMINI_API_KEY", ""),

        "GEMINI_MODEL": os.getenv("GEMINI_MODEL", DEFAULT_MODEL),

        "SCREEN_PROFILE": os.getenv("SCREEN_PROFILE", "auto") or "auto",

        "PRESENCE_SERVER_URL": os.getenv("PRESENCE_SERVER_URL", DEFAULT_PRESENCE_SERVER_URL).strip() or DEFAULT_PRESENCE_SERVER_URL,

    }

    try:

        settings = json.loads(user_settings_path().read_text(encoding="utf-8"))

        if settings.get("GEMINI_API_KEY"):

            values["GEMINI_API_KEY"] = settings["GEMINI_API_KEY"]

        if settings.get("GEMINI_MODEL"):

            values["GEMINI_MODEL"] = settings["GEMINI_MODEL"]

        if settings.get("SCREEN_PROFILE"):

            values["SCREEN_PROFILE"] = str(settings["SCREEN_PROFILE"])

        if settings.get("PRESENCE_SERVER_URL"):

            values["PRESENCE_SERVER_URL"] = str(settings["PRESENCE_SERVER_URL"]).strip() or values["PRESENCE_SERVER_URL"]

    except Exception:

        pass

    if not is_admin_build():

        values["PRESENCE_SERVER_URL"] = DEFAULT_PRESENCE_SERVER_URL

    return values



def save_env(api_key, model, screen_profile="auto", presence_server_url=None):

    resolved_presence_server_url = resolve_presence_server_url(
        presence_server_url if is_admin_build() else DEFAULT_PRESENCE_SERVER_URL
    ) or DEFAULT_PRESENCE_SERVER_URL

    payload = {

        "GEMINI_API_KEY": api_key.strip(),

        "GEMINI_MODEL": (model.strip() or DEFAULT_MODEL),

        "SCREEN_PROFILE": str(screen_profile or "auto").strip() or "auto",

        "PRESENCE_SERVER_URL": resolved_presence_server_url,

    }

    if getattr(sys, "frozen", False):
        backup_file(user_settings_path(), "config")

        user_settings_path().write_text(

            json.dumps(payload, ensure_ascii=False, indent=2),

            encoding="utf-8"

        )

    else:
        backup_file(env_path(), "config")

        env_path().write_text(

            f"GEMINI_API_KEY={api_key.strip()}\nGEMINI_MODEL={(model.strip() or DEFAULT_MODEL)}\nSCREEN_PROFILE={str(screen_profile or 'auto').strip() or 'auto'}\nPRESENCE_SERVER_URL={resolved_presence_server_url}\n",

            encoding="utf-8"

        )

    try:
        backup_file(user_settings_path(), "config")

        user_settings_path().write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    except Exception:

        pass


def normalize_presence_server_url(url):
    value = str(url or "").strip().rstrip("/")
    if not value:
        return ""
    if "://" not in value:
        value = "http://" + value
    return value


def _detect_local_ip_address():
    sock = None
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(0.2)
        try:
            sock.connect(("8.8.8.8", 80))
        except Exception:
            pass
        ip = sock.getsockname()[0]
        if ip and ip != "0.0.0.0":
            return ip
    except Exception:
        pass
    finally:
        try:
            if sock is not None:
                sock.close()
        except Exception:
            pass
    return "127.0.0.1"


def resolve_presence_server_url(url):
    base = normalize_presence_server_url(url)
    if not base:
        return ""
    try:
        parsed = urllib_parse.urlsplit(base)
        host = (parsed.hostname or "").strip().lower()
        if host in {"0.0.0.0", "127.0.0.1", "localhost", ""}:
            resolved_host = _detect_local_ip_address()
            port = f":{parsed.port}" if parsed.port else ""
            return urllib_parse.urlunsplit((parsed.scheme or "http", f"{resolved_host}{port}", parsed.path or "", parsed.query or "", parsed.fragment or ""))
    except Exception:
        pass
    return base


DEFAULT_PRESENCE_SERVER_URL = normalize_presence_server_url(os.getenv("PRESENCE_SERVER_URL_DEFAULT", "http://192.168.1.5:8765"))


def presence_server_url_from_env():
    env = load_env_values()
    return resolve_presence_server_url(env.get("PRESENCE_SERVER_URL") or DEFAULT_PRESENCE_SERVER_URL)


def _presence_client_request_json(url, payload=None, timeout=3, method=None):
    data = None
    headers = {"Content-Type": "application/json"}
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=data, headers=headers, method=method or ("POST" if data is not None else "GET"))
    with urllib_request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
        return json.loads(raw) if raw else None


def file_sha256(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def fetch_user_update_info(server_url, timeout=3):
    base = normalize_presence_server_url(server_url)
    if not base:
        return None
    try:
        data = _presence_client_request_json(f"{base}/update-info", timeout=timeout, method="GET")
        return data if isinstance(data, dict) and data.get("ok") else None
    except Exception:
        return None


def download_user_update(server_url, update_info, timeout=20, progress_callback=None):
    base = normalize_presence_server_url(server_url)
    download_url = str((update_info or {}).get("download_url") or "").strip()
    expected_sha = str((update_info or {}).get("sha256") or "").strip().lower()
    if not base or not download_url or not expected_sha:
        return None
    url = urllib_parse.urljoin(base.rstrip("/") + "/", download_url.lstrip("/"))
    out_dir = Path(tempfile.mkdtemp(prefix="gk_pilepro_update_"))
    out_path = out_dir / "GK PilePro.exe"
    try:
        with urllib_request.urlopen(url, timeout=timeout) as resp, out_path.open("wb") as f:
            total = int(resp.headers.get("Content-Length") or 0)
            downloaded = 0
            chunk_size = 65536  # 64 KB
            _cb_counter = 0
            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                _cb_counter += 1
                # Throttle: update UI every 4 chunks (~256 KB) to avoid flooding event loop
                if progress_callback and _cb_counter % 4 == 0:
                    try:
                        progress_callback(downloaded, total)
                    except Exception:
                        pass
            # Final callback with complete download size
            if progress_callback:
                try:
                    progress_callback(downloaded, total)
                except Exception:
                    pass
        actual_sha = file_sha256(out_path).lower()
        if actual_sha != expected_sha:
            try:
                shutil.rmtree(out_dir, ignore_errors=True)
            except Exception:
                pass
            return None
        return out_path
    except Exception:
        try:
            shutil.rmtree(out_dir, ignore_errors=True)
        except Exception:
            pass
        return None


def send_presence_heartbeat(server_url, payload, timeout=3):
    base = normalize_presence_server_url(server_url)
    if not base:
        return False
    try:
        _presence_client_request_json(f"{base}/heartbeat", payload=payload, timeout=timeout, method="POST")
        return True
    except Exception:
        return False


def check_presence_server_alive(server_url, timeout=2):
    base = normalize_presence_server_url(server_url)
    if not base:
        return False
    try:
        data = _presence_client_request_json(f"{base}/health", timeout=timeout, method="GET")
        return bool(isinstance(data, dict) and data.get("ok"))
    except Exception:
        return False


def fetch_presence_machines(server_url, timeout=3):
    base = normalize_presence_server_url(server_url)
    if not base:
        return []
    try:
        data = _presence_client_request_json(f"{base}/machines", timeout=timeout, method="GET")
        if isinstance(data, dict):
            machines = data.get("machines")
            if isinstance(machines, list):
                return machines
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def fetch_presence_approved_machines(server_url, timeout=3):
    base = normalize_presence_server_url(server_url)
    if not base:
        return []
    try:
        data = _presence_client_request_json(f"{base}/approved-machines", timeout=timeout, method="GET")
        if isinstance(data, dict):
            items = data.get("approved_machines")
            if isinstance(items, list):
                return items
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def send_presence_approved_machine(server_url, payload, timeout=3):
    base = normalize_presence_server_url(server_url)
    if not base:
        return False
    try:
        _presence_client_request_json(f"{base}/approved-machines", payload=payload, timeout=timeout, method="POST")
        return True
    except Exception:
        return False


def send_presence_error_log(server_url, payload, timeout=5):
    base = normalize_presence_server_url(server_url)
    if not base:
        return False
    try:
        data = _presence_client_request_json(f"{base}/error-logs", payload=payload, timeout=timeout, method="POST")
        return bool(isinstance(data, dict) and data.get("ok"))
    except Exception:
        return False


def fetch_presence_error_logs(server_url, limit=100, timeout=3, unresolved_only=False):
    base = normalize_presence_server_url(server_url)
    if not base:
        return []
    try:
        query = f"limit={int(limit or 100)}"
        if unresolved_only:
            query += "&unresolved=1"
        data = _presence_client_request_json(f"{base}/error-logs?{query}", timeout=timeout, method="GET")
        if isinstance(data, dict):
            logs = data.get("error_logs")
            if isinstance(logs, list):
                return logs
    except Exception:
        pass
    return []


def resolve_presence_error_log(server_url, log_id, timeout=3):
    base = normalize_presence_server_url(server_url)
    try:
        log_id = int(log_id)
    except Exception:
        return False
    if not base:
        return False
    try:
        data = _presence_client_request_json(f"{base}/error-logs/{log_id}/resolve", payload={}, timeout=timeout, method="POST")
        return bool(isinstance(data, dict) and data.get("ok"))
    except Exception:
        return False


def delete_presence_approved_machine(server_url, machine_code, timeout=3):
    base = normalize_presence_server_url(server_url)
    machine_code = str(machine_code or "").strip().upper()
    if not base or not machine_code:
        return False
    try:
        _presence_client_request_json(f"{base}/approved-machines/{urllib_parse.quote(machine_code)}", timeout=timeout, method="DELETE")
        return True
    except Exception:
        return False
