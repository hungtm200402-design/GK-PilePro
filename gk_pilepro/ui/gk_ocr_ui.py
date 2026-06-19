"""OCR workflow, image preview, viewer, and clipboard actions for the main window."""

import json
import math
import os
import shutil
import threading
import tkinter as tk
import traceback
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from gk_pilepro.gk_core import last_run_dir, new_workflow_id, resource_path, write_role_error_log
from gk_pilepro.gk_excel import (
    call_gemini,
    call_gemini_phieu_coc,
    merge_ocr_tables_for_continuous_read,
    norm,
    normalize_vietnam_date,
    postprocess_to_hop_coc_d1_d2,
)
from gk_pilepro.ui.gk_ui import (
    UI_BG,
    UI_BORDER,
    UI_ERROR,
    UI_MUTED,
    UI_PRIMARY,
    UI_SUCCESS,
    UI_SURFACE,
    UI_SURFACE_2,
    UI_TEXT,
    ui_button,
    ui_font,
)


def _valid_ocr_bbox(value):
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    try:
        x1, y1, x2, y2 = [max(0.0, min(1000.0, float(item))) for item in value]
    except (TypeError, ValueError):
        return None
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _strip_positive_number_sign(value):
    text = str(value or "").strip()
    if not text.startswith("+"):
        return value
    compact = text.replace(" ", "")
    if not compact[1:]:
        return value
    if all(ch.isdigit() or ch in {",", "."} for ch in compact[1:]) and any(ch.isdigit() for ch in compact[1:]):
        return compact[1:]
    return value


def _merge_line_candidates(candidates, max_gap=2):
    if not candidates:
        return []
    groups = []
    current = [candidates[0]]
    for value in candidates[1:]:
        if value - current[-1] <= max_gap:
            current.append(value)
        else:
            groups.append(current)
            current = [value]
    groups.append(current)
    return [sum(group) / len(group) for group in groups]


def _select_even_line_window(lines, needed, min_pos, max_pos):
    candidates = [line for line in lines if min_pos <= line <= max_pos]
    if len(candidates) < needed:
        return []
    best = None
    best_score = None
    for start in range(0, len(candidates) - needed + 1):
        window = candidates[start:start + needed]
        gaps = [window[i + 1] - window[i] for i in range(len(window) - 1)]
        if not gaps:
            continue
        sorted_gaps = sorted(gaps)
        median = sorted_gaps[len(sorted_gaps) // 2]
        if median <= 2:
            continue
        variance = sum(abs(gap - median) for gap in gaps) / len(gaps)
        score = variance + max(0, min_pos - window[0]) * 0.02
        if best is None or score < best_score:
            best = window
            best_score = score
    return best or []


def _select_grid_lines(lines, needed, min_pos, max_pos):
    candidates = [float(line) for line in lines if min_pos <= line <= max_pos]
    if len(candidates) < needed:
        return []
    candidates = sorted(candidates)
    best = None
    best_score = None
    for start in range(0, len(candidates) - needed + 1):
        window = candidates[start:start + needed]
        gaps = [window[i + 1] - window[i] for i in range(len(window) - 1)]
        if not gaps:
            continue
        sorted_gaps = sorted(gaps)
        median = sorted_gaps[len(sorted_gaps) // 2]
        if median <= 4:
            continue
        variance = sum(abs(gap - median) for gap in gaps) / len(gaps)
        span = window[-1] - window[0]
        edge_score = abs(window[0] - min_pos) * 0.005 + abs(max_pos - window[-1]) * 0.002
        score = variance - span * 0.01 + edge_score
        if best is None or score < best_score:
            best = window
            best_score = score
    return best or []


def _bbox_from_thumb(left, top, right, bottom, width, height):
    return [
        max(0.0, min(1000.0, left * 1000.0 / max(1, width))),
        max(0.0, min(1000.0, top * 1000.0 / max(1, height))),
        max(0.0, min(1000.0, right * 1000.0 / max(1, width))),
        max(0.0, min(1000.0, bottom * 1000.0 / max(1, height))),
    ]


def _opencv_table_grid_bboxes(image_path, row_count, column_count):
    try:
        import cv2
        import numpy as np

        img = cv2.imread(str(image_path), cv2.IMREAD_GRAYSCALE)
        if img is None:
            return [], []
        h, w = img.shape[:2]
        scale = min(1.0, 1600.0 / max(1, max(h, w)))
        if scale < 1:
            work = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
        else:
            work = img
        th = cv2.adaptiveThreshold(
            work,
            255,
            cv2.ADAPTIVE_THRESH_MEAN_C,
            cv2.THRESH_BINARY_INV,
            35,
            12,
        )
        wh, ww = work.shape[:2]
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, ww // 18), 1))
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, wh // 22)))
        horizontal = cv2.morphologyEx(th, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
        vertical = cv2.morphologyEx(th, cv2.MORPH_OPEN, vertical_kernel, iterations=1)

        h_profile = np.count_nonzero(horizontal, axis=1)
        v_profile = np.count_nonzero(vertical, axis=0)
        h_hits = [idx for idx, value in enumerate(h_profile) if value >= ww * 0.22]
        v_hits = [idx for idx, value in enumerate(v_profile) if value >= wh * 0.16]
        h_lines = _merge_line_candidates(h_hits, max_gap=max(2, int(4 * scale)))
        v_lines = _merge_line_candidates(v_hits, max_gap=max(2, int(4 * scale)))

        needed_y = row_count + 2
        y_lines = _select_grid_lines(h_lines, needed_y, wh * 0.08, wh * 0.98)
        if not y_lines:
            y_lines = _select_even_line_window(h_lines, needed_y, wh * 0.12, wh * 0.98)
        if not y_lines:
            return [], []

        needed_x = max(2, column_count + 1)
        x_lines = _select_grid_lines(v_lines, needed_x, ww * 0.01, ww * 0.99)
        if not x_lines:
            x_lines = _select_even_line_window(v_lines, needed_x, ww * 0.01, ww * 0.99)
        if not x_lines:
            contours, _hier = cv2.findContours(cv2.bitwise_or(horizontal, vertical), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            boxes = [cv2.boundingRect(c) for c in contours]
            boxes = [box for box in boxes if box[2] > ww * 0.25 and box[3] > wh * 0.20]
            if boxes:
                x, y, bw, bh = max(boxes, key=lambda box: box[2] * box[3])
                x_lines = [x, x + bw]
        if not x_lines:
            return [], []

        row_bboxes = []
        cell_bboxes = []
        pad_y = max(1.0, min(6.0, (y_lines[-1] - y_lines[0]) / max(1, row_count) * 0.08))
        for row_idx in range(row_count):
            top = y_lines[row_idx + 1]
            bottom = y_lines[row_idx + 2] if row_idx + 2 < len(y_lines) else y_lines[-1]
            if bottom <= top:
                continue
            safe_top = max(0.0, top - pad_y)
            safe_bottom = min(float(wh), bottom + pad_y)
            row_bboxes.append(_bbox_from_thumb(x_lines[0], safe_top, x_lines[-1], safe_bottom, ww, wh))
            row_cells = []
            if len(x_lines) >= column_count + 1:
                for col_idx in range(column_count):
                    left = x_lines[col_idx]
                    right = x_lines[col_idx + 1]
                    row_cells.append(_bbox_from_thumb(left, safe_top, right, safe_bottom, ww, wh))
            else:
                step = (x_lines[-1] - x_lines[0]) / max(1, column_count)
                for col_idx in range(column_count):
                    left = x_lines[0] + step * col_idx
                    right = x_lines[0] + step * (col_idx + 1)
                    row_cells.append(_bbox_from_thumb(left, safe_top, right, safe_bottom, ww, wh))
            cell_bboxes.append(row_cells)
        return row_bboxes, cell_bboxes
    except Exception as exc:
        write_role_error_log("opencv_table_grid_bboxes", exc)
        return [], []


def _fallback_table_grid_bboxes(image_path, row_count, column_count):
    if row_count <= 0:
        return [], []
    cv_rows, cv_cells = _opencv_table_grid_bboxes(image_path, row_count, column_count)
    if len(cv_rows) >= row_count:
        return cv_rows[:row_count], cv_cells[:row_count]
    try:
        img = Image.open(image_path).convert("L")
        thumb = img.copy()
        thumb.thumbnail((900, 1200), Image.Resampling.LANCZOS)
        threshold = 190
        pixels = thumb.load()
        width, height = thumb.size

        horizontal_hits = []
        for y in range(height):
            hits = 0
            for x in range(width):
                if pixels[x, y] < threshold:
                    hits += 1
            if hits >= width * 0.34:
                horizontal_hits.append(y)
        horizontal_lines = _merge_line_candidates(horizontal_hits, max_gap=3)
        needed_y = row_count + 2
        y_lines = _select_even_line_window(
            horizontal_lines,
            needed_y,
            height * 0.22,
            height * 0.98,
        )

        if y_lines:
            top = y_lines[0]
            bottom = y_lines[-1]
        else:
            lower_top = int(height * 0.30)
            xs = []
            ys = []
            for y in range(lower_top, height):
                for x in range(width):
                    if pixels[x, y] < 220:
                        xs.append(x)
                        ys.append(y)
            if xs and ys:
                top = max(0, min(ys) - 6)
                bottom = min(height, max(ys) + 6)
            else:
                top, bottom = int(height * 0.36), int(height * 0.92)
            header_h = max(12, (bottom - top) / max(5, row_count + 1))
            row_h = max(1, (bottom - top - header_h) / max(1, row_count))
            y_lines = [top, top + header_h] + [
                top + header_h + row_h * (idx + 1)
                for idx in range(row_count)
            ]

        vertical_hits = []
        y1_scan = int(max(0, y_lines[0]))
        y2_scan = int(min(height, y_lines[-1]))
        scan_h = max(1, y2_scan - y1_scan)
        for x in range(width):
            hits = 0
            for y in range(y1_scan, y2_scan):
                if pixels[x, y] < threshold:
                    hits += 1
            if hits >= scan_h * 0.34:
                vertical_hits.append(x)
        vertical_lines = _merge_line_candidates(vertical_hits, max_gap=3)
        needed_x = max(2, column_count + 1)
        x_lines = _select_even_line_window(
            vertical_lines,
            needed_x,
            width * 0.03,
            width * 0.98,
        )

        if not x_lines:
            xs = []
            for y in range(y1_scan, y2_scan):
                for x in range(width):
                    if pixels[x, y] < 220:
                        xs.append(x)
            if xs:
                left = max(0, min(xs) - 6)
                right = min(width, max(xs) + 6)
            else:
                left, right = int(width * 0.06), int(width * 0.94)
            step = (right - left) / max(1, column_count)
            x_lines = [left + step * idx for idx in range(column_count + 1)]

        row_bboxes = []
        cell_bboxes = []
        for row_idx in range(row_count):
            top = y_lines[row_idx + 1]
            bottom = y_lines[row_idx + 2] if row_idx + 2 < len(y_lines) else y_lines[-1]
            row_bboxes.append(_bbox_from_thumb(x_lines[0], top, x_lines[-1], bottom, width, height))
            row_cells = []
            for col_idx in range(column_count):
                left = x_lines[col_idx] if col_idx < len(x_lines) else x_lines[0]
                right = x_lines[col_idx + 1] if col_idx + 1 < len(x_lines) else x_lines[-1]
                row_cells.append(_bbox_from_thumb(left, top, right, bottom, width, height))
            cell_bboxes.append(row_cells)
        return row_bboxes, cell_bboxes
    except Exception:
        step = 560.0 / max(1, row_count)
        row_bboxes = [
            [70.0, 360.0 + row_idx * step, 930.0, 360.0 + (row_idx + 1) * step]
            for row_idx in range(row_count)
        ]
        cell_bboxes = []
        col_step = 860.0 / max(1, column_count)
        for row_bbox in row_bboxes:
            row_cells = []
            for col_idx in range(column_count):
                row_cells.append([
                    70.0 + col_idx * col_step,
                    row_bbox[1],
                    70.0 + (col_idx + 1) * col_step,
                    row_bbox[3],
                ])
            cell_bboxes.append(row_cells)
        return row_bboxes, cell_bboxes


def _fallback_row_bboxes_for_image(image_path, row_count, column_count=0):
    row_bboxes, _cell_bboxes = _fallback_table_grid_bboxes(image_path, row_count, column_count)
    return row_bboxes


def _fallback_cell_bboxes_for_image(image_path, row_count, column_count):
    _row_bboxes, cell_bboxes = _fallback_table_grid_bboxes(image_path, row_count, column_count)
    return cell_bboxes


def _bbox_needs_grid_fallback(bbox):
    if not bbox or len(bbox) != 4:
        return True
    try:
        x1, y1, x2, y2 = [float(value) for value in bbox]
        width = x2 - x1
        height = y2 - y1
        if width <= 120 or height <= 8:
            return True
        if height > 95:
            return True
        if y1 < 120 or y2 > 990:
            return True
        return False
    except Exception:
        return True


def ensure_table_image_metadata(table, image_index=0, image_path=None, allow_fallback=True):
    if not isinstance(table, dict):
        return table
    rows = [
        [_strip_positive_number_sign(cell) for cell in row]
        if isinstance(row, list)
        else row
        for row in (table.get("rows") or [])
    ]
    table["rows"] = rows
    row_count = len(rows)
    table["_source_image_index"] = image_index
    if image_path is not None:
        table["_source_image_path"] = str(image_path)
    table["_row_source_indexes"] = [image_index] * row_count

    raw_row_bboxes = list(table.get("_row_bboxes") or table.get("row_bboxes") or [])
    row_bboxes = [
        _valid_ocr_bbox(raw_row_bboxes[row_idx])
        if row_idx < len(raw_row_bboxes)
        else None
        for row_idx in range(row_count)
    ]
    columns = list(table.get("columns") or [])
    column_count = len(columns)
    raw_cell_bboxes = list(table.get("_cell_bboxes") or table.get("cell_bboxes") or [])
    cell_bboxes = []
    for row_idx in range(row_count):
        source_cells = (
            raw_cell_bboxes[row_idx]
            if row_idx < len(raw_cell_bboxes) and isinstance(raw_cell_bboxes[row_idx], list)
            else []
        )
        cell_bboxes.append([
            _valid_ocr_bbox(source_cells[col_idx])
            if col_idx < len(source_cells)
            else None
            for col_idx in range(column_count)
        ])
    if allow_fallback and image_path:
        needs_row_fallback = not row_bboxes or not any(row_bboxes)
        needs_cell_fallback = not cell_bboxes or not any(any(row_cells) for row_cells in cell_bboxes)
        needs_tight_row_fallback = any(_bbox_needs_grid_fallback(bbox) for bbox in row_bboxes)
        if needs_row_fallback or needs_cell_fallback or needs_tight_row_fallback:
            fallback_rows, fallback_cells = _fallback_table_grid_bboxes(
                image_path,
                row_count,
                column_count,
            )
            if needs_row_fallback or needs_tight_row_fallback:
                row_bboxes = fallback_rows
            if needs_cell_fallback or needs_tight_row_fallback:
                cell_bboxes = fallback_cells
    table["_row_bboxes"] = row_bboxes
    table["_cell_bboxes"] = cell_bboxes
    return table


def _clip_notify_text(value, limit):
    text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    return text[: max(0, limit - 1)]


def _play_app_notification_sound():
    try:
        import winsound

        winsound.PlaySound("SystemNotification", winsound.SND_ALIAS | winsound.SND_ASYNC)
    except Exception:
        try:
            import winsound

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass


def _format_elapsed(seconds):
    try:
        seconds = max(0.0, float(seconds))
    except Exception:
        seconds = 0.0
    if seconds < 60:
        return f"{seconds:.1f} giây"
    minutes = int(seconds // 60)
    remain = seconds - minutes * 60
    return f"{minutes} phút {remain:.0f} giây"


def _write_ocr_timing_log(file_name, action, started_at, ended_at, image_count, status, message):
    try:
        out = last_run_dir()
        out.mkdir(exist_ok=True)
        elapsed = max(0.0, (ended_at - started_at).total_seconds())
        payload = {
            "action": action,
            "status": status,
            "started_at": started_at.isoformat(timespec="seconds"),
            "ended_at": ended_at.isoformat(timespec="seconds"),
            "elapsed_seconds": round(elapsed, 3),
            "elapsed_text": _format_elapsed(elapsed),
            "image_count": image_count,
            "message": message,
        }
        (out / file_name).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _bring_app_to_front(app):
    try:
        root = getattr(app, "root", None)
        if root is None or not root.winfo_exists():
            return
        try:
            root.deiconify()
        except Exception:
            pass
        try:
            root.lift()
            root.focus_force()
        except Exception:
            pass
        try:
            import ctypes

            ctypes.windll.user32.SetForegroundWindow(root.winfo_id())
        except Exception:
            pass
    except Exception:
        pass


def _flash_app_taskbar(app):
    try:
        import ctypes
        from ctypes import wintypes

        root = getattr(app, "root", None)
        if root is None or not root.winfo_exists():
            return

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.UINT),
                ("hwnd", wintypes.HWND),
                ("dwFlags", wintypes.DWORD),
                ("uCount", wintypes.UINT),
                ("dwTimeout", wintypes.DWORD),
            ]

        info = FLASHWINFO(
            ctypes.sizeof(FLASHWINFO),
            root.winfo_id(),
            0x00000002 | 0x00000004,
            5,
            0,
        )
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except Exception:
        pass


def _show_windows_balloon_notification(app, title, message):
    try:
        import ctypes
        from ctypes import wintypes

        root = getattr(app, "root", None)
        if root is None or not root.winfo_exists():
            return False

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        class NOTIFYICONDATAW(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("hWnd", wintypes.HWND),
                ("uID", wintypes.UINT),
                ("uFlags", wintypes.UINT),
                ("uCallbackMessage", wintypes.UINT),
                ("hIcon", wintypes.HICON),
                ("szTip", wintypes.WCHAR * 128),
                ("dwState", wintypes.DWORD),
                ("dwStateMask", wintypes.DWORD),
                ("szInfo", wintypes.WCHAR * 256),
                ("uTimeoutOrVersion", wintypes.UINT),
                ("szInfoTitle", wintypes.WCHAR * 64),
                ("dwInfoFlags", wintypes.DWORD),
                ("guidItem", GUID),
                ("hBalloonIcon", wintypes.HICON),
            ]

        NIM_ADD = 0x00000000
        NIM_MODIFY = 0x00000001
        NIM_DELETE = 0x00000002
        NIF_MESSAGE = 0x00000001
        NIF_ICON = 0x00000002
        NIF_TIP = 0x00000004
        NIF_INFO = 0x00000010
        NIIF_USER = 0x00000004
        NIIF_LARGE_ICON = 0x00000020
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32

        icon_path = resource_path("assets", "gk_app_icon.ico")
        hicon = None
        if icon_path.exists():
            hicon = user32.LoadImageW(
                None,
                str(icon_path),
                IMAGE_ICON,
                0,
                0,
                LR_LOADFROMFILE | LR_DEFAULTSIZE,
            )

        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = wintypes.HWND(root.winfo_id())
        nid.uID = int(time.time() * 1000) & 0xFFFFFFFF
        nid.uFlags = NIF_MESSAGE | NIF_TIP
        nid.uCallbackMessage = 0x0400 + 88
        nid.szTip = _clip_notify_text("GK PilePro", 128)
        if hicon:
            nid.uFlags |= NIF_ICON
            nid.hIcon = hicon
            nid.hBalloonIcon = hicon

        if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
            if hicon:
                user32.DestroyIcon(hicon)
            return False

        nid.uFlags |= NIF_INFO
        nid.szInfoTitle = _clip_notify_text(title, 64)
        nid.szInfo = _clip_notify_text(message, 256)
        nid.uTimeoutOrVersion = 10000
        nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON
        shown = bool(shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid)))

        live_icons = getattr(app, "_notification_area_icons", None)
        if live_icons is None:
            live_icons = []
            app._notification_area_icons = live_icons
        live_icons.append((nid, hicon))

        def _cleanup_notification():
            try:
                shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
            except Exception:
                pass
            try:
                if hicon:
                    user32.DestroyIcon(hicon)
            except Exception:
                pass
            try:
                live_icons.remove((nid, hicon))
            except ValueError:
                pass

        root.after(12000, _cleanup_notification)
        return shown
    except Exception:
        return False


def _is_app_in_foreground(app):
    try:
        import ctypes
        from ctypes import wintypes

        root = getattr(app, "root", None)
        if root is None or not root.winfo_exists():
            return False

        try:
            if root.state() == "iconic":
                return False
        except Exception:
            pass

        user32 = ctypes.windll.user32
        foreground_hwnd = user32.GetForegroundWindow()
        if not foreground_hwnd:
            return False

        foreground_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(foreground_hwnd, ctypes.byref(foreground_pid))
        return int(foreground_pid.value) == os.getpid()
    except Exception:
        return False


def _show_ocr_done_notification(app, title, message):
    if _is_app_in_foreground(app):
        return

    now = time.monotonic()
    key = (str(title), str(message))
    last_key = getattr(app, "_last_ocr_notification_key", None)
    last_at = getattr(app, "_last_ocr_notification_at", 0.0)
    if last_key == key and now - float(last_at or 0.0) < 8.0:
        return
    app._last_ocr_notification_key = key
    app._last_ocr_notification_at = now
    _play_app_notification_sound()
    _show_windows_balloon_notification(app, title, message)


def run_gemini_phieu_coc(self):

    """

    Đọc phiếu cọc bằng AI.

    Workflow:

      1. BẮT BUỘC phải chọn Excel trước để biết cột cần điền.

      2. AI trả về 1 bảng duy nhất với cột Y HỆT Excel.

      3. Auto-map 1:1 vì tên cột đã khớp.

      4. Sẵn sàng điền vào Excel ngay.

    """

    image_paths = list(getattr(self, "image_paths", None) or ([] if not self.image_path else [self.image_path]))

    if not image_paths:

        messagebox.showwarning("Thiếu ảnh", "Bạn chưa chọn ảnh.")

        return

    api_key = self.api_key_var.get().strip()

    if not api_key:

        messagebox.showwarning("Thiếu khóa API", "Bạn chưa nhập khóa API.")

        return

    if not self.excel_headers:

        messagebox.showwarning(

            "Chưa chọn Excel",

            "Bạn cần chọn file Excel trước để tool biết cần đọc cột nào.\n\n"

            "Bước 1: Chọn Excel → bấm Đọc lại Excel\n"

            "Bước 2: Chọn ảnh phiếu cọc → bấm Đọc phiếu cọc"

        )

        return



    self.save_key()
    started_at = datetime.now()

    excel_col_names = [name for _, name in self.excel_headers]

    self._set_status(

        f"Đang đọc phiếu cọc... ({len(image_paths)} ảnh, {len(excel_col_names)} cột Excel: "

        + ", ".join(excel_col_names[:5])

        + ("..." if len(excel_col_names) > 5 else "") + ")",

        "warn",

    )

    self.root.update()



    try:

        all_tables = []
        raw_parts = []

        for idx, image_path in enumerate(image_paths, start=1):

            self._set_status(f"Đang đọc phiếu cọc {idx}/{len(image_paths)}...", "warn")
            self.root.update()

            tables_one, raw_one = call_gemini_phieu_coc(

                image_path, api_key, self.model_var.get().strip(),

                excel_columns=excel_col_names

            )

            for table in tables_one or []:
                if isinstance(table, dict):
                    ensure_table_image_metadata(table, idx - 1, image_path)
            all_tables.extend(tables_one or [])
            raw_parts.append(f"=== IMAGE {idx}/{len(image_paths)}: {Path(image_path).name} ===\n{raw_one}")

        tables = all_tables
        raw = "\n\n".join(raw_parts)

        def _normalize_phieu_coc_dates(tables_data):

            normalized_tables = []

            for table in tables_data or []:

                if not isinstance(table, dict):

                    normalized_tables.append(table)

                    continue

                cols = list(table.get("columns") or [])

                date_cols = {
                    idx for idx, name in enumerate(cols)
                    if any(tok in norm(name) for tok in ("ngay", "date"))
                }

                rows = []

                for row in table.get("rows") or []:

                    if not isinstance(row, list):

                        rows.append(row)

                        continue

                    new_row = list(row)

                    for idx in date_cols:

                        if idx < len(new_row):

                            new_row[idx] = normalize_vietnam_date(new_row[idx])

                    rows.append(new_row)

                new_table = dict(table)

                new_table["rows"] = rows

                normalized_tables.append(new_table)

            return normalized_tables

        tables = _normalize_phieu_coc_dates(tables)
        tables = merge_ocr_tables_for_continuous_read(tables)
        for table in tables:
            if isinstance(table, dict):
                table["_source_image_count"] = len(image_paths)



        out = last_run_dir()

        out.mkdir(exist_ok=True)

        (out / "phieu_coc_raw_response.txt").write_text(raw, encoding="utf-8")

        (out / "phieu_coc_tables.json").write_text(

            json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8"

        )



        self.tables = tables

        self.table_editor.set_tables(tables)



        # Tự chuyển sang bảng có nhiều cột nhất (bảng dữ liệu, không phải key-value)

        best_idx = 0

        best_ncols = 0

        for i, t in enumerate(tables):

            nc = len(t.get("columns", []))

            if nc > best_ncols:

                best_ncols = nc

                best_idx = i

        if best_idx != 0:

            try:

                self.table_editor.combo.current(best_idx)

                self.table_editor.switch_table()

            except Exception:

                pass



        # Auto-map: vì cột AI trả về đã đặt tên y hệt Excel

        # → auto_map_columns sẽ khớp 1:1, không cần user chỉnh

        self.build_mapping()



        total_rows = sum(len(t["rows"]) for t in tables)

        data_table = tables[best_idx] if tables else None

        data_cols = data_table.get("columns", []) if data_table else []

        data_rows = len(data_table.get("rows", [])) if data_table else 0



        self._set_status(

            f"Đã đọc phiếu cọc: {len(image_paths)} ảnh, {data_rows} dòng × {len(data_cols)} cột. "

            "Kiểm tra preview rồi bấm 'Điền tiếp vào Excel'.",

            "success",

        )
        ended_at = datetime.now()
        elapsed_text = _format_elapsed((ended_at - started_at).total_seconds())
        _write_ocr_timing_log(
            "ocr_timing_phieu_coc.json",
            "read_phieu_coc",
            started_at,
            ended_at,
            len(image_paths),
            "success",
            f"Đọc phiếu cọc xong: {data_rows} dòng, {len(data_cols)} cột.",
        )
        _show_ocr_done_notification(
            self,
            "Đọc phiếu cọc xong",
            f"Đã đọc xong {len(image_paths)} ảnh phiếu cọc.\n"
            f"Kết quả: {data_rows} dòng, {len(data_cols)} cột.\n"
            f"Thời gian: {elapsed_text}.",
        )



        # Hiển thị tóm tắt trong excel_info

        info_lines = ["KẾT QUẢ ĐỌC PHIẾU CỌC\n", "=" * 50 + "\n"]

        info_lines.append(f"Cột Excel ({len(excel_col_names)}): " + " | ".join(excel_col_names) + "\n\n")
        info_lines.append(f"Số ảnh đã đọc: {len(image_paths)}\n\n")

        for t in tables:

            title = t.get("title") or "Bảng"

            cols = t.get("columns", [])

            rows = t.get("rows", [])

            info_lines.append(f"{title} ({len(rows)} dòng):\n")

            if cols == ["Trường", "Giá trị"]:

                for r in rows:

                    if len(r) >= 2 and (r[0] or r[1]):

                        info_lines.append(f"  {r[0]}: {r[1]}\n")

            else:

                info_lines.append("  Cột: " + " | ".join(cols) + "\n")

                info_lines.append(f"  Số dòng: {len(rows)}\n")

            info_lines.append("\n")

        info_lines.append("Log: last_run_v12\\phieu_coc_tables.json\n")

        self.excel_info.delete("1.0", "end")

        self.excel_info.insert("1.0", "".join(info_lines))
        self._refresh_daily_summary_panel(tables)

        self._record_history(

            "ocr_done",

            file_path=image_paths[0],

            rows=total_rows,

            message=f"Đọc phiếu cọc: {len(tables)} bảng",

            extra={

                "ocr_type": "phieu_coc",

                "table_count": len(tables),

                "best_table_index": best_idx,

                "tables_data": tables,
                "image_count": len(image_paths),

                "tables": [

                    {

                        "title": t.get("title") or "Bảng",

                        "columns": len(t.get("columns", [])),

                        "rows": len(t.get("rows", [])),

                    }

                    for t in tables[:8]

                ],

            },

            workflow_id=self.current_workflow_id,

            workflow_label=self.current_workflow_label,

        )



    except Exception:

        out = last_run_dir()

        out.mkdir(exist_ok=True)

        (out / "last_error_phieu_coc.txt").write_text(traceback.format_exc(), encoding="utf-8")
        ended_at = datetime.now()
        elapsed_text = _format_elapsed((ended_at - started_at).total_seconds())
        _write_ocr_timing_log(
            "ocr_timing_phieu_coc.json",
            "read_phieu_coc",
            started_at,
            ended_at,
            len(image_paths),
            "error",
            "Lỗi đọc phiếu cọc. Xem last_run_v12/last_error_phieu_coc.txt",
        )
        _show_ocr_done_notification(
            self,
            "Lỗi đọc phiếu cọc",
            "Có lỗi khi đọc phiếu cọc.\n"
            f"Thời gian: {elapsed_text}.\n"
            "Xem last_run_v12\\last_error_phieu_coc.txt",
        )

        self._set_status("Lỗi đọc phiếu cọc.", "error")


def choose_image(self):

    paths = filedialog.askopenfilenames(
        title="Chọn ảnh OCR",
        filetypes=[
            ("Image", "*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.gif;*.tif;*.tiff"),
            ("All", "*.*"),
        ],
    )

    if not paths:

        return

    existing_paths = list(getattr(self, "image_paths", None) or [])
    self.current_workflow_id = self.current_workflow_id or new_workflow_id()
    self.current_doc_kind = None

    self._reset_current_preview_state()

    snapshot_paths = [self._save_history_image_snapshot(p, self.current_workflow_id) for p in paths]
    if existing_paths:
        self.append_image_paths(snapshot_paths, "Đã thêm ảnh: ")
    else:
        self.current_workflow_id = new_workflow_id()
        self.set_image_paths(snapshot_paths, "Đã chọn ảnh: ")

    self._record_history(

        "image_selected",

        file_path=snapshot_paths[0],

        message="Chọn ảnh để OCR",

        workflow_id=self.current_workflow_id,

        workflow_label=self.current_workflow_label,

        extra={"image_count": len(snapshot_paths)},
    )


def _preview_thumbnail_size(self):

    w = 260 if not (self.tiny_ui or self.micro_ui) else 220

    h = max(170, int(self.image_drop_h))

    return (w, h)


def _clear_preview_image(self, message=None):

    self._original_image = None
    self.tk_img = None
    if hasattr(self, "img_label") and self.img_label is not None:
        self.img_label.config(image="", text=message or "Kéo thả ảnh OCR vào đây\n\nHỗ trợ: .jpg, .png, .jpeg")
    if hasattr(self, "preview_counter_var") and self.preview_counter_var is not None:
        self.preview_counter_var.set("")
    try:
        if hasattr(self, "preview_prev_btn"):
            self.preview_prev_btn.configure(state="disabled")
        if hasattr(self, "preview_next_btn"):
            self.preview_next_btn.configure(state="disabled")
    except Exception:
        pass


def _load_preview_image(self, path):

    if not path:

        self._clear_preview_image()

        return False

    try:

        self._original_image = Image.open(path)
        im = self._original_image.copy()
        im.thumbnail(self._preview_thumbnail_size())
        self.tk_img = ImageTk.PhotoImage(im)
        self.img_label.config(image=self.tk_img, text="")
        return True

    except Exception as e:

        self._clear_preview_image("Không tải được ảnh\n\nHãy kiểm tra file ảnh.")
        self._set_status(f"Lỗi tải ảnh: {e}", "error")
        return False


def _update_preview_counter(self):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        self.preview_image_index = 0

        if hasattr(self, "preview_counter_var") and self.preview_counter_var is not None:
            self.preview_counter_var.set("")

        return

    idx = getattr(self, "preview_image_index", 0) or 0

    idx = max(0, min(int(idx), len(paths) - 1))

    self.preview_image_index = idx

    if hasattr(self, "preview_counter_var") and self.preview_counter_var is not None:
        self.preview_counter_var.set(f"Ảnh đang xem: {idx + 1}/{len(paths)}")

    try:
        if hasattr(self, "preview_prev_btn"):
            self.preview_prev_btn.configure(state=("normal" if idx > 0 else "disabled"))
        if hasattr(self, "preview_next_btn"):
            self.preview_next_btn.configure(state=("normal" if idx < len(paths) - 1 else "disabled"))
    except Exception:
        pass


def _show_preview_image_index(self, index):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        self._clear_preview_image()

        return

    idx = max(0, min(int(index), len(paths) - 1))

    self.preview_image_index = idx

    self.image_path = paths[idx]

    self._load_preview_image(self.image_path)

    self._update_preview_counter()


def move_preview_image(self, delta):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        return

    self._show_preview_image_index((getattr(self, "preview_image_index", 0) or 0) + delta)


def delete_current_preview_image(self, event=None):

    widget = getattr(event, "widget", None)

    try:

        if widget is not None and widget.winfo_class() in {"Entry", "Text", "TEntry", "TCombobox", "Spinbox"}:

            return None

    except Exception:

        pass

    if getattr(self, "_is_reading_table", False):

        try:

            self._set_status("Đang đọc bảng, chưa thể bỏ ảnh lúc này.", "warn")

        except Exception:

            pass

        return "break"

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        return None

    idx = getattr(self, "preview_image_index", 0) or 0

    idx = max(0, min(int(idx), len(paths) - 1))

    removed_path = paths.pop(idx)

    removed_name = Path(removed_path).name

    self.image_paths = paths

    if paths:

        self.image_path = paths[min(idx, len(paths) - 1)]

        first_name = Path(paths[0]).name

        self.current_workflow_label = (
            f"{first_name} (+{len(paths) - 1} ảnh)" if len(paths) > 1 else first_name
        )

        self._show_preview_image_index(min(idx, len(paths) - 1))

        self._set_status(f"Đã bỏ ảnh: {removed_name}. Còn {len(paths)} ảnh.", "success")

    else:

        self.image_path = None

        self.preview_image_index = 0

        self.current_workflow_label = ""

        self._clear_preview_image()

        self._set_status(f"Đã bỏ ảnh: {removed_name}. Chưa còn ảnh nào.", "warn")

    return "break"


def get_current_preview_image_path(self):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        return None

    idx = getattr(self, "preview_image_index", 0) or 0

    if idx < 0 or idx >= len(paths):

        idx = 0

    return paths[idx]


def open_current_preview_image(self, event=None):

    path = self.get_current_preview_image_path()

    if not path:

        return "break"

    self.open_image_viewer(path)

    return "break"


def _source_image_context_for_selection(self, context=None):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        return None, None

    editor = getattr(self, "table_editor", None)
    if context is None and editor is not None:
        context = editor.get_selected_context()

    table = editor.get_current_table() if editor is not None else None
    row_index = int((context or {}).get("row_index", 0))
    row_sources = list((table or {}).get("_row_source_indexes") or [])
    source_index = (
        row_sources[row_index]
        if 0 <= row_index < len(row_sources)
        else (table or {}).get("_source_image_index")
    )

    row_bboxes = list((table or {}).get("_row_bboxes") or [])
    bbox = (
        row_bboxes[row_index]
        if 0 <= row_index < len(row_bboxes)
        else None
    )

    try:
        index = max(0, min(int(source_index), len(paths) - 1))
    except (TypeError, ValueError):
        index = max(
            0,
            min(int(getattr(self, "preview_image_index", 0) or 0), len(paths) - 1),
        )
    return index, bbox


def show_source_image_for_selection(self, context=None, open_viewer=False):

    source_index, bbox = self._source_image_context_for_selection(context)

    if source_index is None:

        return

    self._show_preview_image_index(source_index)
    path = self.get_current_preview_image_path()
    if path:
        editor = getattr(self, "table_editor", None)
        table = editor.get_current_table() if editor is not None else None
        row_index = int((context or {}).get("row_index", 0))
        rows = list((table or {}).get("rows") or [])
        columns = list((table or {}).get("columns") or [])
        needs_fallback = bbox is None
        try:
            if bbox and len(bbox) == 4:
                y1 = float(bbox[1])
                y2 = float(bbox[3])
                x1 = float(bbox[0])
                x2 = float(bbox[2])
                needs_fallback = (
                    y1 < 120
                    or (y2 - y1) < 10
                    or (y2 - y1) > 120
                    or (x2 - x1) < 180
                )
        except Exception:
            needs_fallback = True
        if needs_fallback:
            fallback_rows, fallback_cells = _fallback_table_grid_bboxes(path, len(rows), len(columns))
            if 0 <= row_index < len(fallback_rows):
                bbox = fallback_rows[row_index]
            if isinstance(table, dict):
                table["_row_bboxes"] = fallback_rows
                table["_cell_bboxes"] = fallback_cells
    editor = getattr(self, "table_editor", None)
    if editor is not None and hasattr(editor, "show_source_row_crop"):
        editor.show_source_row_crop(path, bbox, context=context)

    if open_viewer:

        if path:
            self.open_image_viewer(path, focus_bbox=bbox)


def set_image_path(self, path, status_prefix="Đã chọn ảnh: "):

    self.set_image_paths([path], status_prefix=status_prefix)


def set_image_paths(self, paths, status_prefix="Đã chọn ảnh: "):

    clean_paths = [str(p) for p in (paths or []) if str(p or "").strip()]

    self.image_paths = clean_paths
    self.image_path = clean_paths[0] if clean_paths else None
    self.current_workflow_id = self.current_workflow_id or new_workflow_id()

    if clean_paths:
        first_name = Path(clean_paths[0]).name
        if len(clean_paths) > 1:
            self.current_workflow_label = f"{first_name} (+{len(clean_paths) - 1} ảnh)"
        else:
            self.current_workflow_label = first_name
    else:
        self.current_workflow_label = ""

    if clean_paths:
        self._show_preview_image_index(0)
    else:
        self._clear_preview_image()

    suffix = Path(clean_paths[0]).name if clean_paths else "không có ảnh"
    if len(clean_paths) > 1:
        suffix = f"{Path(clean_paths[0]).name} (+{len(clean_paths) - 1} ảnh)"

    self._set_status(status_prefix + suffix)


def append_image_paths(self, paths, status_prefix="Đã thêm ảnh: "):

    clean_paths = [str(p) for p in (paths or []) if str(p or "").strip()]

    if not clean_paths:

        return

    current_paths = list(getattr(self, "image_paths", None) or [])

    combined = current_paths + clean_paths

    self.set_image_paths(combined, status_prefix=status_prefix)


def open_image_viewer(self, path, focus_bbox=None):

    paths = list(getattr(self, "image_paths", None) or [])

    if not paths:

        return

    try:

        start_index = paths.index(path)

    except ValueError:

        start_index = 0

    state = getattr(self, "_viewer_state", None)

    if state and state.get("win") is not None:

        try:

            if state["win"].winfo_exists():

                state["index"] = start_index
                state["focus_bbox"] = focus_bbox

                state["win"].lift()

                if callable(getattr(self, "_viewer_set_index", None)):

                    self._viewer_set_index(start_index)

                return

        except Exception:

            pass

    win = tk.Toplevel(self.root)
    win.title(f"Xem ảnh - {Path(path).name}")
    win.configure(bg=UI_BG)
    win.transient(self.root)

    sw = max(640, int(self.root.winfo_screenwidth() * 0.8))
    sh = max(480, int(self.root.winfo_screenheight() * 0.8))
    win.geometry(f"{sw}x{sh}")

    outer = tk.Frame(win, bg=UI_BG)
    outer.pack(fill="both", expand=True, padx=12, pady=12)

    title_var = tk.StringVar(value=Path(path).name)
    tk.Label(
        outer,
        textvariable=title_var,
        bg=UI_BG,
        fg=UI_TEXT,
        font=ui_font(12, bold=True),
        anchor="w",
    ).pack(fill="x", anchor="w")

    toolbar = tk.Frame(outer, bg=UI_BG)
    toolbar.pack(fill="x", pady=(8, 0))
    zoom_var = tk.StringVar(value="100%")

    canvas_frame = tk.Frame(outer, bg=UI_BG)
    canvas_frame.pack(fill="both", expand=True, pady=(10, 0))

    canvas = tk.Canvas(canvas_frame, bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER)
    x_scroll = ttk.Scrollbar(canvas_frame, orient="horizontal", command=canvas.xview)
    y_scroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
    canvas.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")
    canvas_frame.columnconfigure(0, weight=1)
    canvas_frame.rowconfigure(0, weight=1)

    prev_btn = tk.Button(
        canvas,
        text="‹",
        bg="#ffffff",
        fg=UI_TEXT,
        activebackground="#e8f1ff",
        activeforeground=UI_TEXT,
        relief="flat",
        bd=0,
        font=ui_font(20, bold=True),
        cursor="hand2",
        width=2,
        height=1,
    )
    next_btn = tk.Button(
        canvas,
        text="›",
        bg="#ffffff",
        fg=UI_TEXT,
        activebackground="#e8f1ff",
        activeforeground=UI_TEXT,
        relief="flat",
        bd=0,
        font=ui_font(20, bold=True),
        cursor="hand2",
        width=2,
        height=1,
    )
    prev_btn_id = canvas.create_window(28, 0, window=prev_btn, anchor="center", state="hidden")
    next_btn_id = canvas.create_window(0, 0, window=next_btn, anchor="center", state="hidden")
    image_id = canvas.create_image(0, 0, anchor="center")
    focus_id = canvas.create_rectangle(
        0,
        0,
        0,
        0,
        outline="#ff2d20",
        width=4,
        state="hidden",
    )
    canvas.image = None

    state = {
        "win": win,
        "canvas": canvas,
        "image_id": image_id,
        "focus_id": focus_id,
        "title_var": title_var,
        "prev_btn_id": prev_btn_id,
        "next_btn_id": next_btn_id,
        "paths": paths,
        "index": start_index,
        "sw": sw,
        "sh": sh,
        "zoom": 1.0,
        "fit_zoom": 1.0,
        "original_image": None,
        "focus_bbox": focus_bbox,
    }
    self._viewer_state = state

    def _sync_image_position(_event=None):
        try:
            canvas.update_idletasks()
            cw = max(1, canvas.winfo_width())
            ch = max(1, canvas.winfo_height())
            current = getattr(canvas, "image", None)
            iw = max(1, current.width() if current else 1)
            ih = max(1, current.height() if current else 1)
            canvas.coords(image_id, max(cw / 2, iw / 2), max(ch / 2, ih / 2))
            canvas.coords(prev_btn_id, 28, ch / 2)
            canvas.coords(next_btn_id, max(28, cw - 28), ch / 2)
            canvas.configure(scrollregion=(0, 0, max(cw, iw), max(ch, ih)))
        except Exception:
            pass

    def _render_current_image():
        original = state.get("original_image")
        if original is None:
            return
        zoom = max(0.15, min(float(state.get("zoom") or 1.0), 5.0))
        state["zoom"] = zoom
        width = max(1, int(original.width * zoom))
        height = max(1, int(original.height * zoom))
        resized = original.resize((width, height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(resized)
        canvas.image = photo
        canvas.itemconfigure(image_id, image=photo)
        zoom_var.set(f"{zoom * 100:.0f}%")
        _sync_image_position()
        win.after_idle(_focus_selected_bbox)

    def _set_zoom(zoom):
        state["zoom"] = max(0.15, min(float(zoom), 5.0))
        _render_current_image()

    def _fit_image():
        original = state.get("original_image")
        if original is None:
            return
        canvas.update_idletasks()
        available_w = max(1, canvas.winfo_width() - 20)
        available_h = max(1, canvas.winfo_height() - 20)
        fit_zoom = min(
            available_w / max(1, original.width),
            available_h / max(1, original.height),
            1.0,
        )
        state["fit_zoom"] = max(0.15, fit_zoom)
        _set_zoom(state["fit_zoom"])

    def _focus_selected_bbox():
        bbox = state.get("focus_bbox")
        original = state.get("original_image")
        current = getattr(canvas, "image", None)
        if (
            original is None
            or current is None
            or not isinstance(bbox, (list, tuple))
            or len(bbox) != 4
        ):
            canvas.itemconfigure(focus_id, state="hidden")
            return
        try:
            x1, y1, x2, y2 = [float(value) for value in bbox]
        except (TypeError, ValueError):
            canvas.itemconfigure(focus_id, state="hidden")
            return

        canvas.update_idletasks()
        cw = max(1, canvas.winfo_width())
        ch = max(1, canvas.winfo_height())
        zoom = float(state.get("zoom") or 1.0)
        target_w = max(1.0, (x2 - x1) * original.width / 1000.0)
        target_h = max(1.0, (y2 - y1) * original.height / 1000.0)
        desired_zoom = min(
            5.0,
            max(
                float(state.get("fit_zoom") or 0.15),
                min(cw * 0.72 / target_w, ch * 0.42 / target_h),
            ),
        )
        if abs(desired_zoom - zoom) > 0.03:
            state["zoom"] = desired_zoom
            _render_current_image()
            return

        iw = current.width()
        ih = current.height()
        origin_x = max(cw / 2, iw / 2) - iw / 2
        origin_y = max(ch / 2, ih / 2) - ih / 2
        left = origin_x + x1 * iw / 1000.0
        top = origin_y + y1 * ih / 1000.0
        right = origin_x + x2 * iw / 1000.0
        bottom = origin_y + y2 * ih / 1000.0
        pad = 10
        canvas.coords(
            focus_id,
            left - pad,
            top - pad,
            right + pad,
            bottom + pad,
        )
        canvas.itemconfigure(focus_id, state="normal")
        canvas.tag_raise(focus_id)

        scrollregion = canvas.bbox("all")
        if scrollregion:
            total_w = max(cw, scrollregion[2] - scrollregion[0])
            total_h = max(ch, scrollregion[3] - scrollregion[1])
            center_x = (left + right) / 2
            center_y = (top + bottom) / 2
            if total_w > cw:
                canvas.xview_moveto(max(0.0, min(1.0, (center_x - cw / 2) / total_w)))
            if total_h > ch:
                canvas.yview_moveto(max(0.0, min(1.0, (center_y - ch / 2) / total_h)))

    def _load_current_image():
        current_paths = list(state.get("paths") or [])
        if not current_paths:
            return
        idx = max(0, min(int(state.get("index") or 0), len(current_paths) - 1))
        state["index"] = idx
        current_path = current_paths[idx]
        try:
            img = Image.open(current_path).convert("RGB")
        except Exception as e:
            messagebox.showerror("Xem ảnh", f"Không mở được ảnh:\n{e}")
            return
        state["original_image"] = img.copy()
        title_var.set(f"{Path(current_path).name} ({idx + 1}/{len(current_paths)})")
        self.preview_image_index = idx
        self.image_path = current_path
        self._update_preview_counter()
        win.after_idle(_fit_image)

    def _move(delta):
        current_paths = list(state.get("paths") or [])
        if not current_paths:
            return
        state["index"] = (int(state.get("index") or 0) + delta) % len(current_paths)
        _load_current_image()

    state["hover_hide_job"] = None

    def _viewer_pointer_inside():
        try:
            x, y = canvas.winfo_pointerxy()
            widget = canvas.winfo_containing(x, y)
            while widget is not None:
                if widget is canvas or widget is prev_btn or widget is next_btn:
                    return True
                widget = getattr(widget, "master", None)
        except Exception:
            pass
        return False

    def _hover_on(_event=None):
        try:
            job = state.get("hover_hide_job")
            if job:
                canvas.after_cancel(job)
                state["hover_hide_job"] = None
        except Exception:
            pass
        canvas.itemconfigure(prev_btn_id, state="normal")
        canvas.itemconfigure(next_btn_id, state="normal")

    def _hover_off(_event=None):
        def _hide_if_outside():
            state["hover_hide_job"] = None
            if not _viewer_pointer_inside():
                canvas.itemconfigure(prev_btn_id, state="hidden")
                canvas.itemconfigure(next_btn_id, state="hidden")

        try:
            job = state.get("hover_hide_job")
            if job:
                canvas.after_cancel(job)
            state["hover_hide_job"] = canvas.after(140, _hide_if_outside)
        except Exception:
            _hide_if_outside()

    prev_btn.configure(command=lambda: _move(-1))
    next_btn.configure(command=lambda: _move(1))
    canvas.bind("<Configure>", _sync_image_position)
    canvas.bind("<Enter>", _hover_on)
    canvas.bind("<Leave>", _hover_off)
    prev_btn.bind("<Enter>", _hover_on)
    prev_btn.bind("<Leave>", _hover_off)
    next_btn.bind("<Enter>", _hover_on)
    next_btn.bind("<Leave>", _hover_off)

    tk.Button(
        toolbar,
        text="−",
        command=lambda: _set_zoom((state.get("zoom") or 1.0) / 1.2),
        bg="#eef4ff",
        fg=UI_TEXT,
        relief="flat",
        width=4,
    ).pack(side="left")
    tk.Label(
        toolbar,
        textvariable=zoom_var,
        bg=UI_BG,
        fg=UI_TEXT,
        font=ui_font(10, bold=True),
        width=7,
    ).pack(side="left", padx=4)
    tk.Button(
        toolbar,
        text="+",
        command=lambda: _set_zoom((state.get("zoom") or 1.0) * 1.2),
        bg="#eef4ff",
        fg=UI_TEXT,
        relief="flat",
        width=4,
    ).pack(side="left")
    tk.Button(
        toolbar,
        text="Vừa màn hình",
        command=_fit_image,
        bg="#eef4ff",
        fg=UI_TEXT,
        relief="flat",
        padx=12,
    ).pack(side="left", padx=(8, 0))
    tk.Label(
        toolbar,
        text="Ctrl + lăn chuột để zoom",
        bg=UI_BG,
        fg=UI_MUTED,
        font=ui_font(9),
    ).pack(side="right")

    def _zoom_with_wheel(event):
        factor = 1.12 if event.delta > 0 else (1 / 1.12)
        _set_zoom((state.get("zoom") or 1.0) * factor)
        return "break"

    canvas.bind("<Control-MouseWheel>", _zoom_with_wheel)

    def _viewer_set_index(idx):
        state["index"] = idx
        _load_current_image()

    self._viewer_set_index = _viewer_set_index
    self._sync_viewer_layout = _sync_image_position

    _load_current_image()
    _hover_on()

    def _close(_event=None):
        try:
            win.grab_release()
        except Exception:
            pass
        try:
            if getattr(self, "_viewer_state", None) and self._viewer_state.get("win") is win:
                self._viewer_state = None
        except Exception:
            pass
        win.destroy()
        return "break"

    win.bind("<Escape>", _close)
    tk.Button(outer, text="Đóng", command=_close, bg="#eef4ff", fg=UI_TEXT, relief="flat", padx=16, pady=6).pack(anchor="e", pady=(10, 0))


def _reset_current_preview_state(self):
    self.tables = []
    self.current_doc_kind = None
    try:
        if hasattr(self, "table_editor") and self.table_editor is not None:
            self.table_editor.set_tables([])
    except Exception:
        pass
    try:
        if hasattr(self, "mapping_editor") and self.mapping_editor is not None:
            self.mapping_editor.clear()
    except Exception:
        pass
    try:
        if hasattr(self, "excel_info") and self.excel_info is not None:
            self.excel_info.delete("1.0", "end")
    except Exception:
        pass
    try:
        self._clear_preview_image()
    except Exception:
        pass
    try:
        self._set_daily_summary_text([])
    except Exception:
        pass
    self.history_selected_entry = None
    try:
        self._render_history_detail(None)
    except Exception:
        pass


def _save_history_image_snapshot(self, source_path, workflow_id=None):
    try:
        source = Path(str(source_path))
        if not source.exists():
            return str(source_path)
        out = last_run_dir() / "history_images"
        out.mkdir(parents=True, exist_ok=True)
        wf = str(workflow_id or self.current_workflow_id or new_workflow_id()).strip() or new_workflow_id()
        suffix = source.suffix.lower() if source.suffix else ".png"
        safe_name = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in source.stem)[:48] or "image"
        target = out / f"{wf}_{safe_name}{suffix}"
        shutil.copy2(str(source), str(target))
        return str(target)
    except Exception:
        return str(source_path)


def paste_image_from_clipboard(self, event=None):
    try:

        from PIL import ImageGrab
        import time

        clip = None
        for attempt in range(5):
            clip = ImageGrab.grabclipboard()
            if clip is not None:
                break
            if attempt < 4:
                time.sleep(0.15)

    except Exception:

        messagebox.showwarning("Clipboard", "Không đọc được clipboard ảnh trên máy này.")

        if self._clipboard_text_widget_has_focus():

            return None

        return "break"



    image = None

    source_path = None

    if isinstance(clip, Image.Image):

        image = clip.convert("RGB")

    elif isinstance(clip, list):

        for item in clip:

            try:

                p = Path(item)

                if p.exists() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".jfif"}:

                    source_path = p

                    break

            except Exception:

                pass



    if source_path:

        self.current_workflow_id = new_workflow_id()

        self.current_workflow_label = Path(source_path).name
        self.current_doc_kind = None

        self._reset_current_preview_state()
        snapshot_path = self._save_history_image_snapshot(source_path, self.current_workflow_id)
        if getattr(self, "image_paths", None):
            self.append_image_paths([snapshot_path], "Đã dán ảnh từ clipboard: ")
        else:
            self.set_image_path(snapshot_path, "Đã dán ảnh từ clipboard: ")

        self._record_history(

            "clipboard_paste",

            file_path=snapshot_path,

            message="Dán ảnh từ clipboard",

            workflow_id=self.current_workflow_id,

            workflow_label=self.current_workflow_label,

        )

        return "break"



    if image is None:

        try:

            text = self.root.clipboard_get()

        except Exception:

            text = ""

        if str(text or "").strip():

            if self._clipboard_text_widget_has_focus():

                return None

            self._set_status("Clipboard đang là text. Hãy copy ảnh bằng Ctrl+C rồi bấm Ctrl+V để dán ảnh.", "warn")

        else:

            messagebox.showinfo(

                "Clipboard",

                "Clipboard chưa có ảnh hoặc Windows chưa trả dữ liệu ảnh đúng định dạng. "
                "Hãy copy đúng file/ảnh bằng Ctrl+C rồi thử lại. Nếu vẫn lỗi, dùng nút Chọn ảnh / Tải lên."

            )

        return "break"



    try:

        out = last_run_dir()

        out.mkdir(exist_ok=True)

        paste_path = out / f"{new_workflow_id()}_clipboard_paste.png"

        image.save(paste_path)

        self.current_workflow_id = new_workflow_id()

        self.current_workflow_label = paste_path.name
        self.current_doc_kind = None

        self._reset_current_preview_state()
        snapshot_path = self._save_history_image_snapshot(paste_path, self.current_workflow_id)
        if getattr(self, "image_paths", None):
            self.append_image_paths([snapshot_path], "Đã dán ảnh từ clipboard: ")
        else:
            self.set_image_path(snapshot_path, "Đã dán ảnh từ clipboard: ")

        self._record_history(

            "clipboard_paste",

            file_path=snapshot_path,

            message="Dán ảnh từ clipboard",

            workflow_id=self.current_workflow_id,

            workflow_label=self.current_workflow_label,

        )

    except Exception:

        out = last_run_dir()

        out.mkdir(exist_ok=True)

        (out / "last_error_clipboard.txt").write_text(traceback.format_exc(), encoding="utf-8")

        messagebox.showerror("Clipboard", "Không lưu được ảnh clipboard. Xem last_run_v12/last_error_clipboard.txt")

    return "break"


def run_gemini(self):

    image_paths = list(getattr(self, "image_paths", None) or ([] if not self.image_path else [self.image_path]))

    if not image_paths:

        messagebox.showwarning("Thiếu ảnh", "Bạn chưa chọn ảnh.")

        return

    api_key = self.api_key_var.get().strip()

    if not api_key:

        messagebox.showwarning("Thiếu khóa API", "Bạn chưa nhập khóa API.")

        return



    self.save_key()

    self._set_status(f"Đang đọc ảnh... ({len(image_paths)} ảnh)", "warn")

    self.root.update()



    try:

        all_tables = []
        raw_parts = []

        for idx, image_path in enumerate(image_paths, start=1):

            self._set_status(f"Đang đọc ảnh {idx}/{len(image_paths)}...", "warn")
            self.root.update()

            tables_one, raw_one = call_gemini(image_path, api_key, self.model_var.get().strip())

            for table in tables_one or []:
                ensure_table_image_metadata(table, idx - 1, image_path)

            tables_one = postprocess_to_hop_coc_d1_d2(tables_one)

            for table in tables_one or []:
                ensure_table_image_metadata(table, idx - 1, image_path)

            all_tables.extend(tables_one or [])
            raw_parts.append(f"=== IMAGE {idx}/{len(image_paths)}: {Path(image_path).name} ===\n{raw_one}")

        tables = merge_ocr_tables_for_continuous_read(all_tables)
        for table in tables:
            if isinstance(table, dict):
                table["_source_image_count"] = len(image_paths)
        raw = "\n\n".join(raw_parts)

        out = last_run_dir()

        out.mkdir(exist_ok=True)

        (out / "ai_raw_response.txt").write_text(raw, encoding="utf-8")

        (out / "ai_tables.json").write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")



        # V20: giữ nguyên cấu trúc bảng mà ảnh trả về, không ép mẫu cố định.

        self.tables = tables

        self.table_editor.set_tables(tables)
        self._refresh_daily_summary_panel(tables)

        self.build_mapping()
        self.current_doc_kind = "bang_khoi_luong"



        total_rows = sum(len(t["rows"]) for t in tables)

        self._set_status(f"Đọc xong: {len(image_paths)} ảnh, {len(tables)} bảng, {total_rows} dòng. Đã giữ cấu trúc đúng theo ảnh.", "success")
        _show_ocr_done_notification(
            self,
            "Đọc bảng xong",
            f"Đã đọc xong {len(image_paths)} ảnh.\n"
            f"Kết quả: {len(tables)} bảng, {total_rows} dòng.",
        )

        self._record_history(

            "ocr_done",

            file_path=image_paths[0],

            rows=total_rows,

            message=f"Đọc xong {len(tables)} bảng",

            extra={

                "ocr_type": "bang_khoi_luong",

                "table_count": len(tables),

                "tables_data": tables,
                "image_count": len(image_paths),

                "tables": [

                    {

                        "title": t.get("title") or "Bảng",

                        "columns": len(t.get("columns", [])),

                        "rows": len(t.get("rows", [])),

                    }

                    for t in tables[:8]

                ],

            },

            workflow_id=self.current_workflow_id,

            workflow_label=self.current_workflow_label,

        )

    except Exception:

        out = last_run_dir()

        out.mkdir(exist_ok=True)

        (out / "last_error.txt").write_text(traceback.format_exc(), encoding="utf-8")

        self._set_status("Lỗi đọc ảnh", "error")

        self._record_history(

            "ocr_error",

            status="error",

            file_path=image_paths[0],

            message="Lỗi khi OCR ảnh",

            workflow_id=self.current_workflow_id,

            workflow_label=self.current_workflow_label,

        )


def install_ocr_ui(app_cls):
    app_cls.run_gemini_phieu_coc = run_gemini_phieu_coc
    app_cls.choose_image = choose_image
    app_cls._preview_thumbnail_size = _preview_thumbnail_size
    app_cls._clear_preview_image = _clear_preview_image
    app_cls._load_preview_image = _load_preview_image
    app_cls._update_preview_counter = _update_preview_counter
    app_cls._show_preview_image_index = _show_preview_image_index
    app_cls.move_preview_image = move_preview_image
    app_cls.delete_current_preview_image = delete_current_preview_image
    app_cls.get_current_preview_image_path = get_current_preview_image_path
    app_cls.open_current_preview_image = open_current_preview_image
    app_cls._source_image_context_for_selection = _source_image_context_for_selection
    app_cls.show_source_image_for_selection = show_source_image_for_selection
    app_cls.set_image_path = set_image_path
    app_cls.set_image_paths = set_image_paths
    app_cls.append_image_paths = append_image_paths
    app_cls.open_image_viewer = open_image_viewer
    app_cls._reset_current_preview_state = _reset_current_preview_state
    app_cls._save_history_image_snapshot = _save_history_image_snapshot
    app_cls.paste_image_from_clipboard = paste_image_from_clipboard
    app_cls.run_gemini = run_gemini
