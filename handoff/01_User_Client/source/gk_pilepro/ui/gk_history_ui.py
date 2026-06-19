"""History rendering views for the main application window."""

import copy
import difflib
import json
import math
import os
import re
import tkinter as tk
import unicodedata
from datetime import datetime
from pathlib import Path
from tkinter import ttk

from PIL import Image, ImageTk

from gk_pilepro.gk_core import last_run_dir, load_history_entries
from gk_pilepro.gk_excel import normalize_vietnam_date
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


def _history_status_style(self, entry):

    status = str(entry.get("status") or "").lower()

    action = str(entry.get("action") or "").lower()

    if status == "error":

        return "#fee2e2", "#dc2626", "#991b1b"

    if action == "export_excel":

        return "#dcfce7", "#16a34a", "#15803d"

    if action == "read_excel":

        return "#dbeafe", "#2563eb", "#1d4ed8"

    return "#f3f4f6", "#6b7280", "#4b5563"


def _history_ocr_haystack(self, entry):

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}

    parts = [

        str(entry.get("date") or ""),

        str(entry.get("time") or ""),

        str(entry.get("timestamp") or ""),

        str(entry.get("workflow_id") or ""),

        str(entry.get("workflow_label") or ""),

        str(entry.get("file_name") or ""),

        str(entry.get("file_path") or ""),

        str(entry.get("message") or ""),

        str(extra.get("ocr_type") or ""),

        str(extra.get("table_count") or ""),

        str(extra.get("best_table_index") or ""),

    ]

    for item in extra.get("tables") or []:

        if isinstance(item, dict):

            parts.extend(

                [

                    str(item.get("title") or ""),

                    str(item.get("columns") or ""),

                    str(item.get("rows") or ""),

                ]

            )

        else:

            parts.append(str(item))

    return " | ".join(parts).lower()


def _history_is_khoi_luong_entry(self, entry):

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    ocr_type = str(extra.get("ocr_type") or "").strip().lower()
    action = str(entry.get("action") or "").strip().lower()

    if ocr_type == "bang_khoi_luong":
        return True

    if ocr_type == "phieu_coc":
        return False

    haystack = self._history_ocr_haystack(entry)
    return any(
        token in haystack
        for token in (
            "bang tong hop khoi luong ep coc",
            "tong hop khoi luong",
            "khoi luong",
            "stt | ngay | ten coc | loai coc",
            "d1 | d2 | d3 | d4",
        )
    )


def _history_entry_kind(self, entry):

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    ocr_type = str(extra.get("ocr_type") or "").strip().lower()
    action = str(entry.get("action") or "").strip().lower()

    if ocr_type == "bang_khoi_luong":
        return "bang_khoi_luong"
    if ocr_type == "phieu_coc":
        return "phieu_coc"
    if action in {"phieu_coc_read", "phieu_coc_error"}:
        return "phieu_coc"
    if action in {"ocr_done", "ocr_error", "export_excel", "export_error"} and self._history_is_khoi_luong_entry(entry):
        return "bang_khoi_luong"
    return "unknown"


def _history_entry_kind_label(self, entry):

    kind = self._history_entry_kind(entry)
    if kind == "bang_khoi_luong":
        return "Bảng khối lượng"
    if kind == "phieu_coc":
        return "Phiếu cọc"
    return "OCR"


def _history_mode_matches_entry(self, entry, mode):

    mode = str(mode or "all").strip().lower()
    kind = self._history_entry_kind(entry)

    if mode in {"", "all", "ca hai", "cả hai"}:
        return kind in {"bang_khoi_luong", "phieu_coc"}
    if mode in {"bang_khoi_luong", "khoi_luong", "khối lượng"}:
        return kind == "bang_khoi_luong"
    if mode in {"phieu_coc", "phiếu cọc"}:
        return kind == "phieu_coc"
    return True


def _history_set_detail_text(self, text):

    widget = getattr(self, "history_detail_text", None)

    if widget is None:
        host = getattr(self, "history_detail_host", None)
        if host is None:
            return
        self._history_clear_detail_host()
        tk.Label(
            host,
            text=text or "",
            bg=UI_SURFACE,
            fg=UI_TEXT,
            justify="left",
            anchor="nw",
            font=("Consolas", 9),
            wraplength=360,
        ).pack(fill="both", expand=True, anchor="nw")
        return

    try:

        widget.configure(state="normal")

        widget.delete("1.0", "end")

        widget.insert("1.0", text or "")

        widget.configure(state="disabled")

        widget.see("1.0")

    except Exception:

        pass


def _history_format_cell(self, value):

    if value is None:

        return ""

    if isinstance(value, (dict, list, tuple)):

        try:

            return json.dumps(value, ensure_ascii=False)

        except Exception:

            return str(value)

    return str(value)


def _history_display_date(self, value):

    text = normalize_vietnam_date(value)

    return text or "Không rõ"


def _history_format_row(self, row):

    if isinstance(row, dict):

        return " | ".join(f"{k}: {self._history_format_cell(v)}" for k, v in row.items())

    if isinstance(row, (list, tuple)):

        return " | ".join(self._history_format_cell(v) for v in row)

    return self._history_format_cell(row)


def _history_make_box_line(self, text, width=66):

    content = str(text or "")[: max(0, width - 2)]

    return "│ " + content.ljust(max(0, width - 2)) + " │"


def _history_norm_col(self, value):

    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def _history_make_row_map(self, table):

    columns = table.get("columns") or []
    rows = table.get("rows") or []
    mapped = []
    for idx, row in enumerate(rows, start=1):
        if isinstance(row, dict):
            row_map = {self._history_norm_col(k): v for k, v in row.items()}
        elif isinstance(row, (list, tuple)):
            row_map = {}
            for c_idx, col in enumerate(columns):
                key = self._history_norm_col(col)
                if c_idx < len(row):
                    row_map[key] = row[c_idx]
        else:
            row_map = {"value": row}
        row_map.setdefault("stt", idx)
        mapped.append(row_map)
    return mapped


def _history_pick_field(self, row_map, *needles, default=""):

    for key, value in row_map.items():
        norm_key = self._history_norm_col(key)
        if any(n in norm_key for n in needles):
            if value not in (None, ""):
                return value
    return default


def _history_render_metric_box(self, title, rows, width=66):

    lines = []
    lines.append("┌" + "─" * width + "┐")
    lines.append(self._history_make_box_line(title, width))
    lines.append("├" + "─" * width + "┤")
    for line in rows:
        lines.append(self._history_make_box_line(line, width))
    lines.append("└" + "─" * width + "┘")
    return "\n".join(lines)


def _history_pretty_column_label(self, label):

    norm = self._history_norm_col(label)
    mapping = {
        "stt": "STT",
        "ngay xuat": "Ngày xuất",
        "so phieu": "Số phiếu",
        "tong so m coc d500": "Tổng số m cọc D500",
        "tong so m coc": "Tổng số m cọc",
        "tong so m": "Tổng số m",
        "ghi chu": "Ghi chú",
        "hop dong": "Hợp đồng",
        "mui d300": "Mũi D300",
        "xe van chuyen": "Xe vận chuyển",
    }
    for key, pretty in mapping.items():
        if key == norm:
            return pretty
    return str(label or "")


def _history_score_table(self, table):

    columns = table.get("columns") or []
    if not columns:
        return 0

    norm_cols = [self._history_norm_col(c) for c in columns]
    joined = " | ".join(norm_cols)

    score = 0

    primary_hits = [
        ("stt", 8),
        ("ngay", 8),
        ("ten coc", 8),
        ("loai coc", 8),
        ("d1", 4),
        ("d2", 4),
        ("d3", 3),
        ("d4", 3),
        ("chieu dai", 4),
        ("tong", 3),
        ("ghi chu", 2),
        ("hop dong", 2),
    ]

    for needle, weight in primary_hits:
        if needle in joined:
            score += weight

    score += min(len(columns), 20)
    score += min(len(table.get("rows") or []), 10)

    return score


def _history_pick_main_table(self, entry, tables_data):

    if not tables_data:
        return None

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    preferred_title = str(extra.get("filled_table_title") or extra.get("source_table_title") or "").strip().lower()
    best_table_index = extra.get("best_table_index")

    if preferred_title:
        for table in tables_data:
            if not isinstance(table, dict):
                continue
            title = str(table.get("title") or "").strip().lower()
            if title and preferred_title in title:
                return table

    try:
        if best_table_index not in (None, ""):
            idx = int(best_table_index)
            if 0 <= idx < len(tables_data) and isinstance(tables_data[idx], dict):
                return tables_data[idx]
    except Exception:
        pass

    scored = []
    for idx, table in enumerate(tables_data):
        if isinstance(table, dict):
            scored.append((self._history_score_table(table), idx, table))

    if not scored:
        return None

    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return scored[0][2]


def _history_clear_detail_host(self):

    host = getattr(self, "history_detail_host", None)
    if host is None:
        return
    for child in host.winfo_children():
        child.destroy()


def _history_build_detail_grid(self, table):

    host = getattr(self, "history_detail_host", None)
    if host is None:
        return

    self._history_clear_detail_host()

    columns = table.get("columns") or []
    rows = table.get("rows") or []
    row_maps = self._history_make_row_map(table)

    display_columns = [self._history_pretty_column_label(c) for c in columns] if columns else []
    if not display_columns and row_maps:
        display_columns = [self._history_pretty_column_label(k) for k in row_maps[0].keys()]

    if not display_columns:
        tk.Label(host, text="Không có cột dữ liệu để hiển thị.", bg=UI_SURFACE, fg=UI_MUTED, font=ui_font(11)).pack(anchor="w")
        return

    top = tk.Frame(host, bg=UI_SURFACE)
    top.pack(fill="x", pady=(0, 8))

    tk.Label(top, text=self._history_pretty_column_label(table.get("title") or "Bảng OCR"), bg=UI_SURFACE, fg=UI_TEXT, font=ui_font(11, bold=True), anchor="w").pack(anchor="w")

    table_shell = tk.Frame(host, bg=UI_SURFACE)
    table_shell.pack(fill="both", expand=True)

    vsb = ttk.Scrollbar(table_shell, orient="vertical")
    hsb = ttk.Scrollbar(table_shell, orient="horizontal")
    tree = ttk.Treeview(
        table_shell,
        columns=[f"c{i}" for i in range(len(display_columns))],
        show="headings",
        style="Preview.Treeview",
        yscrollcommand=vsb.set,
        xscrollcommand=hsb.set,
    )
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)

    for idx, col_name in enumerate(display_columns):
        col_id = f"c{idx}"
        width = 90
        norm_col = self._history_norm_col(columns[idx] if idx < len(columns) else col_name)
        if "ngay xuat" in norm_col:
            width = 100
        elif any(k in norm_col for k in ("ghi chu", "hop dong")):
            width = 110
        elif any(k in norm_col for k in ("tong so", "mui", "xe van chuyen")):
            width = 120
        elif norm_col == "stt":
            width = 50
        tree.heading(col_id, text=col_name)
        tree.column(col_id, width=width, anchor="center", stretch=False)

    if not row_maps:
        for row in rows[:10]:
            if isinstance(row, (list, tuple)):
                values = [self._history_format_cell(v) for v in row[:len(display_columns)]]
            elif isinstance(row, dict):
                values = [self._history_format_cell(v) for v in row.values()]
            else:
                values = [self._history_format_cell(row)]
            values += [""] * max(0, len(display_columns) - len(values))
            tree.insert("", "end", values=values[:len(display_columns)])
    else:
        for idx, row_map in enumerate(row_maps, start=1):
            values = []
            for col_idx, col in enumerate(columns):
                norm = self._history_norm_col(col)
                value = row_map.get(norm, "")
                if (value in (None, "")) and norm == "stt":
                    value = idx
                values.append(self._history_format_cell(value))
            values += [""] * max(0, len(display_columns) - len(values))
            tree.insert("", "end", values=values[:len(display_columns)])

    tree.tag_configure("odd", background="#ffffff")
    tree.tag_configure("even", background="#f8fbff")
    for i, iid in enumerate(tree.get_children()):
        tree.item(iid, tags=("even" if i % 2 else "odd",))


def _history_render_table_preview(self, entry, table, width=66):

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    table_title = str(table.get("title") or "Bang OCR").strip()
    columns = table.get("columns") or []
    rows = table.get("rows") or []
    row_maps = self._history_make_row_map(table)
    lines = []

    def _field_value(row_map, *needles, default=""):
        value = self._history_pick_field(row_map, *needles, default=default)
        return self._history_format_cell(value) if value not in (None, "") else default

    use_structured = False
    target_fields = {
        "ngay xuat",
        "so phieu",
        "tong so m coc",
        "tong so m coc d500",
        "tong so m",
        "ghi chu",
        "hop dong",
    }
    for row_map in row_maps[:3]:
        if any(
            any(target in self._history_norm_col(k) for target in target_fields)
            for k in row_map.keys()
        ):
            use_structured = True
            break

    lines.append("┌" + "─" * width + "┐")
    lines.append(self._history_make_box_line("KHUNG HÔM NAY", width))
    lines.append(self._history_make_box_line(f"Ngày: {self._history_display_date(entry.get('date'))}", width))
    lines.append(self._history_make_box_line(f"Loại: {extra.get('ocr_type') or entry.get('action') or 'ocr'}", width))
    lines.append(self._history_make_box_line(f"Bảng: {table_title}", width))
    lines.append(self._history_make_box_line(f"Cột: {len(columns)}  Dòng: {len(rows)}", width))
    lines.append("├" + "─" * width + "┤")

    if use_structured:
        shown = 0
        for idx, row_map in enumerate(row_maps[:6], start=1):
            stt = _field_value(row_map, "stt", default=str(idx))
            ngay_xuat = _field_value(row_map, "ngay xuat", default="")
            tong_so = _field_value(row_map, "tong so m coc", "tong so m coc d500", "tong so m", default="")
            ghi_chu = _field_value(row_map, "ghi chu", "note", default="")
            hop_dong = _field_value(row_map, "hop dong", "contract", default="")
            so_phieu = _field_value(row_map, "so phieu", default="")
            if so_phieu:
                lines.append(self._history_make_box_line(f"Số phiếu: {so_phieu}", width))
            lines.append(self._history_make_box_line(f"STT: {stt}", width))
            if ngay_xuat:
                lines.append(self._history_make_box_line(f"Ngày xuất: {ngay_xuat}", width))
            if tong_so:
                lines.append(self._history_make_box_line(f"Tổng số m cọc: {tong_so}", width))
            if ghi_chu:
                lines.append(self._history_make_box_line(f"Ghi chú: {ghi_chu}", width))
            if hop_dong:
                lines.append(self._history_make_box_line(f"Hợp đồng: {hop_dong}", width))
            if idx < min(len(row_maps), 6):
                lines.append(self._history_make_box_line("─" * min(width - 4, 30), width))
            shown += 1
        if len(row_maps) > shown:
            lines.append(self._history_make_box_line(f"... còn {len(row_maps) - shown} dòng nữa", width))
    else:
        preview_rows = rows[:6]
        if columns:
            header_cols = [self._history_format_cell(c) for c in columns[:8]]
            lines.append(self._history_make_box_line("CỘT: " + " | ".join(header_cols), width))

        if preview_rows:
            lines.append(self._history_make_box_line("DỮ LIỆU:", width))
            for idx, row in enumerate(preview_rows, start=1):
                if isinstance(row, dict):
                    items = list(row.items())[:8]
                    row_text = " | ".join(f"{k}: {self._history_format_cell(v)}" for k, v in items)
                elif isinstance(row, (list, tuple)):
                    row_text = " | ".join(self._history_format_cell(v) for v in row[:8])
                else:
                    row_text = self._history_format_cell(row)
                lines.append(self._history_make_box_line(f"{idx}. {row_text}", width))
        else:
            lines.append(self._history_make_box_line("Không có dòng dữ liệu.", width))

        if len(rows) > len(preview_rows):
            lines.append(self._history_make_box_line(f"... còn {len(rows) - len(preview_rows)} dòng nữa", width))

    lines.append("└" + "─" * width + "┘")
    return "\n".join(lines)


def _history_load_tables_for_entry(self, entry):

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    action = str(entry.get("action") or "").strip().lower()

    if action in {"export_excel", "export_error"}:
        source_excel_name = str(extra.get("source_excel_name") or "").strip()
        output_name = str(extra.get("output_name") or "").strip()
        source_table_title = str(extra.get("filled_table_title") or extra.get("source_table_title") or "Dữ liệu đã điền").strip()
        cols = extra.get("filled_table_columns")
        rows = extra.get("filled_rows_data")
        if isinstance(cols, list) and isinstance(rows, list) and rows:
            normalized_rows = []
            for row in rows[:100]:
                if isinstance(row, (list, tuple)):
                    normalized_rows.append(list(row))
                elif isinstance(row, dict):
                    normalized_rows.append(list(row.values()))
                else:
                    normalized_rows.append([row])
            return [
                {
                    "title": source_table_title,
                    "columns": cols,
                    "rows": normalized_rows,
                }
            ]
        return [
            {
                "title": "Xuất Excel khối lượng",
                "columns": ["Excel gốc", "Excel xuất", "Sheet", "Dòng bắt đầu", "Số dòng", "Cột bảng"],
                "rows": [[
                    source_excel_name,
                    output_name,
                    entry.get("sheet") or "",
                    extra.get("start_fill_row") or "",
                    entry.get("rows") or "",
                    extra.get("table_columns") or "",
                ]],
            }
        ]

    tables_data = extra.get("tables_data")
    if isinstance(tables_data, list) and tables_data:
        return tables_data

    if action in {"read_excel", "scan_workbook", "workbook_loaded"}:
        last_run = last_run_dir()
        profile_file = last_run / "current_workbook_profiles.json"
        if profile_file.exists():
            try:
                profiles = json.loads(profile_file.read_text(encoding="utf-8"))
                if isinstance(profiles, list) and profiles:
                    rows = []
                    for p in profiles:
                        if not isinstance(p, dict):
                            continue
                        chain = p.get("selected_chain") or {}
                        if isinstance(chain, dict):
                            chain_text = f"{chain.get('from_stt') or ''} → {chain.get('to_stt') or ''}".strip()
                        else:
                            chain_text = ""
                        rows.append([
                            p.get("sheet") or "",
                            p.get("header_row") or "",
                            p.get("total_row") or "",
                            p.get("stt_col") or "",
                            chain_text,
                            len(p.get("sum_columns") or []),
                        ])
                    if rows:
                        return [
                            {
                                "title": "Tổng quan workbook khối lượng",
                                "columns": ["Sheet", "Header", "TỔNG", "STT", "Chuỗi STT", "Cột SUM"],
                                "rows": rows,
                            }
                        ]
            except Exception:
                pass

    tables = extra.get("tables")
    if isinstance(tables, list) and tables and all(isinstance(t, dict) and ("rows" in t or "columns" in t) for t in tables):
        fallback = []
        for table in tables:
            if not isinstance(table, dict):
                continue
            fallback.append(
                {
                    "title": table.get("title") or "Bang OCR",
                    "columns": table.get("columns") if isinstance(table.get("columns"), list) else [],
                    "rows": table.get("rows") if isinstance(table.get("rows"), list) else [],
                }
            )
        if fallback:
            return fallback

    ocr_type = str(extra.get("ocr_type") or "").strip().lower()
    last_run = last_run_dir()
    candidate_files = []
    if ocr_type == "phieu_coc":
        candidate_files.extend([last_run / "phieu_coc_tables.json", last_run / "phieu_coc_raw_response.txt"])
    elif ocr_type == "bang_khoi_luong":
        candidate_files.extend([last_run / "ai_tables.json", last_run / "ai_raw_response.txt"])
    else:
        candidate_files.extend([
            last_run / "phieu_coc_tables.json",
            last_run / "ai_tables.json",
        ])

    for candidate in candidate_files:
        if not candidate.exists():
            continue
        try:
            if candidate.suffix.lower() == ".json":
                loaded = json.loads(candidate.read_text(encoding="utf-8"))
                if isinstance(loaded, list):
                    normalized = []
                    for table in loaded:
                        if isinstance(table, dict):
                            normalized.append(
                                {
                                    "title": table.get("title") or "Bang OCR",
                                    "columns": table.get("columns") if isinstance(table.get("columns"), list) else [],
                                    "rows": table.get("rows") if isinstance(table.get("rows"), list) else [],
                                }
                            )
                    if normalized:
                        return normalized
        except Exception:
            pass

    return []


def _render_history_detail(self, entry=None):

    host = getattr(self, "history_detail_host", None)
    if host is None and not getattr(self, "history_detail_text", None):
        return

    if host is not None:
        self._history_clear_detail_host()

    if not entry:
        if host is not None:
            tk.Label(
                host,
                text="Chọn một mục OCR ở bên trái để xem lại dữ liệu đã đọc.\n\nDouble-click vào một card OCR để mở chi tiết.",
                bg=UI_SURFACE,
                fg=UI_MUTED,
                justify="left",
                anchor="nw",
                font=ui_font(11),
                wraplength=360,
            ).pack(fill="both", expand=True, anchor="nw")
        else:
            self._history_set_detail_text(
                "Chọn một mục OCR ở bên trái để xem lại dữ liệu đã đọc.\n\n"
                "Double-click vào một card OCR để mở chi tiết."
            )

        return

    extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
    tables_data = self._history_load_tables_for_entry(entry)

    title = str(entry.get("workflow_label") or entry.get("file_name") or entry.get("file_path") or "OCR").strip()
    ocr_type = str(extra.get("ocr_type") or "").strip()
    action = str(entry.get("action") or "").strip()
    status = str(entry.get("status") or "").strip()
    rows_value = entry.get("rows")
    table_count = extra.get("table_count")
    best_table_index = extra.get("best_table_index")

    header_rows = [
        f"Tên: {title}",
        f"Thời gian: {entry.get('timestamp') or entry.get('date') or ''} {entry.get('time') or ''}".strip(),
        f"Loại: {ocr_type or action}",
        f"Trạng thái: {status}",
    ]
    if rows_value not in (None, ""):
        header_rows.append(f"Số dòng: {rows_value}")
    if table_count not in (None, ""):
        header_rows.append(f"Số bảng: {table_count}")
    if best_table_index not in (None, ""):
        header_rows.append(f"Bảng chính: {best_table_index}")
    if entry.get("message"):
        header_rows.append(f"Thông điệp: {entry.get('message')}")
    if extra.get("source_excel_name"):
        header_rows.append(f"Excel nguồn: {extra.get('source_excel_name')}")
    if extra.get("output_name"):
        header_rows.append(f"Excel xuất: {extra.get('output_name')}")
    if extra.get("source_image_path"):
        header_rows.append(f"Ảnh OCR: {Path(str(extra.get('source_image_path'))).name}")
    if extra.get("source_table_title"):
        header_rows.append(f"Bảng nguồn: {extra.get('source_table_title')}")
    if entry.get("file_path"):
        header_rows.append(f"File: {entry.get('file_path')}")

    if host is not None:
        summary_box = tk.Frame(host, bg="#f8fbff", highlightthickness=1, highlightbackground="#d7e3f2", padx=10, pady=8)
        summary_box.pack(fill="x", pady=(0, 8))
        for line in header_rows:
            tk.Label(summary_box, text=line, bg="#f8fbff", fg=UI_TEXT, font=ui_font(10), anchor="w", justify="left").pack(anchor="w")
    else:
        self._history_set_detail_text("\n".join(header_rows))

    chosen_table = self._history_pick_main_table(entry, tables_data)

    if chosen_table:
        if host is not None:
            self._history_build_detail_grid(chosen_table)
        else:
            lines = [self._history_render_table_preview(entry, chosen_table)]
            self._history_set_detail_text("\n".join(lines))
    else:
        summary_tables = extra.get("tables") or []
        fallback_lines = []
        if summary_tables:
            fallback_lines.append("Bảng đã đọc:")
            for idx, table in enumerate(summary_tables, start=1):
                if not isinstance(table, dict):
                    continue
                table_title = str(table.get("title") or f"Bảng {idx}").strip()
                fallback_lines.append(
                    f"[{idx}] {table_title} - {table.get('rows') or 0} dòng - {table.get('columns') or 0} cột"
                )
        else:
            fallback_lines.append("Không có dữ liệu OCR chi tiết trong record này.")
        if host is not None:
            tk.Label(host, text="\n".join(fallback_lines), bg=UI_SURFACE, fg=UI_MUTED, font=ui_font(10), justify="left", anchor="nw", wraplength=360).pack(fill="x", anchor="nw")
        else:
            self._history_set_detail_text("\n".join(fallback_lines))


def _render_history_view(self):

    inner = getattr(self, "history_inner", None)

    if inner is None:

        return

    for child in inner.winfo_children():

        child.destroy()

    entries = [e for e in load_history_entries() if isinstance(e, dict)]

    keep_actions = {"image_selected", "clipboard_paste", "ocr_done", "ocr_error", "export_excel", "export_error", "phieu_coc_read", "phieu_coc_error"}

    entries = [e for e in entries if str(e.get("action") or "").lower() in keep_actions]

    entries.sort(key=lambda e: str(e.get("timestamp") or ""), reverse=True)

    grouped = {}

    for entry in entries:

        day = str(entry.get("date") or "")[:10] or "Không rõ"

        file_path = str(entry.get("file_path") or "")

        group_key = (day, file_path)

        bucket = grouped.setdefault(

            group_key,

            {

                "date": day,

                "file_path": file_path,

                "file_name": str(entry.get("file_name") or Path(file_path).name or "Không rõ"),

                "entries": [],

                "has_read": False,

                "has_export": False,

                "has_error": False,

                "latest_ts": str(entry.get("timestamp") or ""),

            },

        )

        bucket["entries"].append(entry)

        if str(entry.get("action") or "").lower() == "read_excel":

            bucket["has_read"] = True

        if str(entry.get("action") or "").lower() == "export_excel":

            bucket["has_export"] = True

        if str(entry.get("status") or "").lower() == "error":

            bucket["has_error"] = True

        ts = str(entry.get("timestamp") or "")

        if ts and ts > bucket["latest_ts"]:

            bucket["latest_ts"] = ts

    grouped = {k: v for k, v in grouped.items() if v["has_ocr"] or v["has_export"] or v["has_error"]}

    if not grouped:

        empty = tk.Frame(inner, bg=UI_SURFACE, padx=18, pady=26, highlightthickness=1, highlightbackground="#e7edf6")

        empty.pack(fill="x")

        tk.Label(empty, text="Chưa có lịch sử.", bg=UI_SURFACE, fg=UI_TEXT, font=ui_font(11, bold=True)).pack(anchor="center")
        tk.Label(empty, text="Sau khi đọc Excel và xuất dữ liệu, các mục sẽ xuất hiện ở đây theo ngày.", bg=UI_SURFACE, fg=UI_MUTED, font=ui_font(10)).pack(anchor="center", pady=(4, 0))
        self.history_summary_var.set("0 mục")
        return

    grouped_items = sorted(grouped.values(), key=lambda g: g.get("latest_ts") or "", reverse=True)

    total_events = len(entries)

    total_files = len(grouped_items)

    self.history_summary_var.set(f"{total_files} file - {total_events} sự kiện")

    current_day = None

    current_day_box = None

    for group in grouped_items:

        if group["date"] != current_day:

            current_day = group["date"]
            current_day_display = self._history_display_date(current_day)

            day_box = tk.Frame(inner, bg=UI_SURFACE, padx=0, pady=0)

            day_box.pack(fill="x", pady=(0, 12))

            current_day_box = day_box

            day_head = tk.Frame(day_box, bg=UI_SURFACE)

            day_head.pack(fill="x", pady=(0, 8))

            tk.Label(day_head, text=current_day_display, bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 11, "bold")).pack(side="left")

            tk.Label(

                day_head,

                text=f"{sum(1 for g in grouped_items if g['date'] == current_day)} file",

                bg="#eef4ff",

                fg=UI_PRIMARY,

                font=ui_font(10, bold=True),

                padx=10,

                pady=5,

            ).pack(side="right")

        status_bg, status_border, accent = self._history_status_style(

            {"status": "error" if group["has_error"] else "success", "action": "export_excel" if group["has_export"] else ("read_excel" if group["has_read"] else "")}

        )

        row = tk.Frame(

            current_day_box,

            bg=status_bg,

            highlightthickness=1,

            highlightbackground=status_border,

            padx=12,

            pady=10,

        )

        row.pack(fill="x", pady=(0, 8))

        accent_bar = tk.Frame(row, bg=accent, width=6)

        accent_bar.pack(side="left", fill="y", padx=(0, 12))

        accent_bar.pack_propagate(False)

        thumb_box = tk.Frame(row, bg=status_bg, width=96, height=120)
        thumb_box.pack(side="left", padx=(0, 12))
        thumb_box.pack_propagate(False)

        primary_entry = next((candidate for candidate in group["entries"] if str(candidate.get("action") or "").lower() in {"ocr_done", "phieu_coc_read"}), None)
        if primary_entry is None:
            primary_entry = group["entries"][0] if group["entries"] else None
        thumb_source = str(primary_entry.get("file_path") or "").strip() if primary_entry else ""

        if thumb_source and Path(thumb_source).exists():
            try:
                img = Image.open(thumb_source).convert("RGBA")
                img.thumbnail((88, 112))
                tk_img = ImageTk.PhotoImage(img)
                self.history_image_refs = getattr(self, "history_image_refs", [])
                self.history_image_refs.append(tk_img)
                tk.Label(thumb_box, image=tk_img, bg=status_bg).pack(fill="both", expand=True)
            except Exception:
                tk.Label(thumb_box, text="OCR", bg="#e5e7eb", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(fill="both", expand=True)
        else:
            tk.Label(thumb_box, text="OCR", bg="#e5e7eb", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(fill="both", expand=True)

        body = tk.Frame(row, bg=status_bg)

        body.pack(side="left", fill="both", expand=True)

        head = tk.Frame(body, bg=status_bg)

        head.pack(fill="x")

        tk.Label(head, text=group["file_name"], bg=status_bg, fg=UI_TEXT, font=ui_font(11, bold=True), anchor="w").pack(side="left")

        badges = tk.Frame(head, bg=status_bg)
        badges.pack(side="right")

        if group["has_read"]:
            tk.Label(badges, text="Đã đọc", bg="#dbeafe", fg="#1d4ed8", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))

        if group["has_export"]:
            tk.Label(badges, text="Đã xuất", bg="#dcfce7", fg="#15803d", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))

        if group["has_error"]:
            tk.Label(badges, text="Lỗi", bg="#fee2e2", fg="#b91c1c", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left")

        tk.Label(body, text=group["file_path"] or "Không rõ đường dẫn", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(anchor="w", pady=(3, 0))

        summary_parts = []

        seen_rows = set()

        for e in group["entries"]:

            action = str(e.get("action") or "").strip()

            label = "Đọc" if action == "read_excel" else "Xuất" if action == "export_excel" else action or "Khác"

            time_text = str(e.get("time") or "")[:8]

            parts = [time_text, label]

            if e.get("sheet"):

                parts.append(f"Sheet: {e['sheet']}")

            if e.get("rows") not in (None, ""):

                parts.append(f"{e['rows']} dòng")

            if e.get("message"):

                parts.append(str(e["message"]))

            row_text = " - ".join(parts)

            dedup_key = (action, str(e.get("sheet") or ""), str(e.get("rows") or ""), str(e.get("message") or ""))

            if dedup_key in seen_rows:

                continue

            seen_rows.add(dedup_key)

            summary_parts.append(row_text)

        first_ocr_entry = next((e for e in group["entries"] if str(e.get("action") or "").lower() in {"ocr_done", "phieu_coc_read"}), None)

        if first_ocr_entry and isinstance(first_ocr_entry.get("extra"), dict):

            tables = first_ocr_entry["extra"].get("tables") or []

            for t in tables[:4]:

                title = str(t.get("title") or "Bảng").strip()

                cols = t.get("columns")

                rows = t.get("rows")

                summary_parts.append(f"{title} - {rows or 0} dòng, {cols or 0} cột")

        details = tk.Frame(body, bg=status_bg)

        details.pack(fill="x", pady=(8, 0))

        for idx, text in enumerate(summary_parts[:4]):
            tk.Label(details, text=f"• {text}", bg=status_bg, fg=UI_TEXT, font=ui_font(10), anchor="w", justify="left", wraplength=900).pack(anchor="w", pady=(0, 2))

        if len(summary_parts) > 4:
            tk.Label(details, text=f"• ... và {len(summary_parts) - 4} mục khác", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(anchor="w", pady=(0, 2))

        def open_history_file(_e=None, target_path=group["file_path"]):

            if target_path and Path(target_path).exists():

                try:

                    self._load_excel_file(target_path)

                    self.show_excel_page()

                except Exception:

                    pass

            return "break"

        for widget in (row, accent_bar, body, head, badges, details):

            widget.bind("<Double-Button-1>", open_history_file)

        for child in body.winfo_children():

            child.bind("<Double-Button-1>", open_history_file)


def _render_history_workflow_view(self):

    inner = getattr(self, "history_inner", None)

    if inner is None:

        return

    for child in inner.winfo_children():

        child.destroy()

    entries = [e for e in load_history_entries() if isinstance(e, dict)]

    entries.sort(key=lambda e: str(e.get("timestamp") or ""), reverse=True)

    grouped = {}

    for entry in entries:

        day = str(entry.get("date") or "")[:10] or "Không rõ"

        wf_id = str(entry.get("workflow_id") or "").strip() or str(entry.get("timestamp") or "")

        key = (day, wf_id)

        bucket = grouped.setdefault(

            key,

            {

                "date": day,

                "workflow_id": wf_id,

                "workflow_label": str(entry.get("workflow_label") or entry.get("file_name") or entry.get("file_path") or "Workflow").strip(),

                "entries": [],

                "has_image": False,

                "has_export": False,

                "has_ocr": False,

                "has_error": False,

                "ocr_type": "",

                "latest_ts": str(entry.get("timestamp") or ""),

            },

        )

        if not bucket["workflow_label"]:

            bucket["workflow_label"] = str(entry.get("file_name") or entry.get("file_path") or "Workflow").strip()

        bucket["entries"].append(entry)

        action = str(entry.get("action") or "").lower()

        if action in {"image_selected", "clipboard_paste"}:

            bucket["has_image"] = True

        if action in {"ocr_done", "phieu_coc_read"}:

            bucket["has_ocr"] = True

        if action == "export_excel":

            bucket["has_export"] = True

        if str(entry.get("status") or "").lower() == "error":

            bucket["has_error"] = True

        extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}

        if not bucket["ocr_type"]:

            bucket["ocr_type"] = str(extra.get("ocr_type") or "").strip()

        ts = str(entry.get("timestamp") or "")

        if ts and ts > bucket["latest_ts"]:

            bucket["latest_ts"] = ts

    if not grouped:

        empty = tk.Frame(inner, bg=UI_SURFACE, padx=18, pady=26, highlightthickness=1, highlightbackground="#e7edf6")

        empty.pack(fill="x")

        tk.Label(empty, text="Chưa có lịch sử quy trình.", bg=UI_SURFACE, fg=UI_TEXT, font=ui_font(11, bold=True)).pack(anchor="center")
        tk.Label(empty, text="Mỗi workflow sẽ lưu các bước như chọn ảnh, OCR, xem trước và xuất Excel.", bg=UI_SURFACE, fg=UI_MUTED, font=ui_font(10)).pack(anchor="center", pady=(4, 0))
        self.history_summary_var.set("0 OCR")
        return

    grouped_items = sorted(grouped.values(), key=lambda g: g.get("latest_ts") or "", reverse=True)

    self.history_summary_var.set(f"{len(grouped_items)} OCR - {len(entries)} bước")

    current_day = None

    current_day_box = None

    for group in grouped_items:

        if group["date"] != current_day:

            current_day = group["date"]
            current_day_display = self._history_display_date(current_day)

            current_day_box = tk.Frame(inner, bg=UI_SURFACE)

            current_day_box.pack(fill="x", pady=(0, 12))

            day_head = tk.Frame(current_day_box, bg=UI_SURFACE)

            day_head.pack(fill="x", pady=(0, 8))

            tk.Label(day_head, text=current_day_display, bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 11, "bold")).pack(side="left")

            tk.Label(day_head, text=f"{sum(1 for g in grouped_items if g['date'] == current_day)} workflow", bg="#eef4ff", fg=UI_PRIMARY, font=ui_font(10, bold=True), padx=10, pady=5).pack(side="right")

        status_bg, status_border, accent = self._history_status_style(
            {"status": "error" if group["has_error"] else "success", "action": "export_excel" if group["has_export"] else ("ocr_done" if group["has_ocr"] else "")}
        )

        row = tk.Frame(current_day_box, bg=status_bg, highlightthickness=1, highlightbackground=status_border, padx=12, pady=10)
        row.pack(fill="x", pady=(0, 8))

        accent_bar = tk.Frame(row, bg=accent, width=6)
        accent_bar.pack(side="left", fill="y", padx=(0, 12))
        accent_bar.pack_propagate(False)

        body = tk.Frame(row, bg=status_bg)
        body.pack(side="left", fill="both", expand=True)

        head = tk.Frame(body, bg=status_bg)

        head.pack(fill="x")

        title_text = group["workflow_label"]

        ocr_type = str(group.get("ocr_type") or "").strip().lower()
        if ocr_type == "bang_khoi_luong":
            title_text = f"Bảng khối lượng: {title_text}"
        elif ocr_type == "phieu_coc":
            title_text = f"Phiếu cọc: {title_text}"
        elif group["has_ocr"]:
            title_text = f"OCR: {title_text}"

        tk.Label(head, text=title_text, bg=status_bg, fg=UI_TEXT, font=ui_font(11, bold=True), anchor="w").pack(side="left")

        badges = tk.Frame(head, bg=status_bg)
        badges.pack(side="right")

        if group["has_image"]:
            tk.Label(badges, text="Có ảnh", bg="#ede9fe", fg="#6d28d9", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))

        if group["has_ocr"]:
            tk.Label(badges, text="Đã đọc", bg="#dbeafe", fg="#1d4ed8", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))

        if group["has_export"]:
            tk.Label(badges, text="Đã xuất", bg="#dcfce7", fg="#15803d", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))

        if group["has_error"]:
            tk.Label(badges, text="Lỗi", bg="#fee2e2", fg="#b91c1c", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left")

        info_line = tk.Frame(body, bg=status_bg)
        info_line.pack(fill="x", pady=(3, 0))
        tk.Label(info_line, text=f"ID: {group['workflow_id']}", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(side="left")

        summary_parts = []

        label_map = {

            "image_selected": "Chọn ảnh",

            "clipboard_paste": "Dán ảnh",

            "ocr_done": "Đọc OCR",

            "export_excel": "Xuất Excel",

            "ocr_error": "Lỗi OCR",

            "export_error": "Lỗi xuất",

            "phieu_coc_read": "Đọc phiếu cọc",

            "phieu_coc_error": "Lỗi phiếu cọc",

        }

        for e in group["entries"]:

            action = str(e.get("action") or "").strip()

            label = label_map.get(action, action or "Khác")

            time_text = str(e.get("time") or "")[:8]

            parts = [time_text, label]

            if e.get("sheet"):

                parts.append(f"Sheet: {e['sheet']}")

            if e.get("rows") not in (None, ""):

                parts.append(f"{e['rows']} dòng")

            if e.get("message"):

                parts.append(str(e["message"]))

            summary_parts.append(" - ".join(parts))

        details = tk.Frame(body, bg=status_bg)

        details.pack(fill="x", pady=(8, 0))

        for text in summary_parts[:5]:
            tk.Label(details, text=f"• {text}", bg=status_bg, fg=UI_TEXT, font=ui_font(10), anchor="w", justify="left", wraplength=900).pack(anchor="w", pady=(0, 2))

        if len(summary_parts) > 5:
            tk.Label(details, text=f"• ... và {len(summary_parts) - 5} bước khác", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(anchor="w", pady=(0, 2))

        def open_history_item(_e=None, target_entry=group["entries"][0] if group["entries"] else None):

            try:

                if not target_entry:

                    return "break"

                path = str(target_entry.get("file_path") or "").strip()

                if not path or not Path(path).exists():

                    return "break"

                p = Path(path)

                if p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".jfif"}:

                    self.set_image_path(p, "Đã mở lại từ lịch sử: ")

                    self.show_home_page()

                else:

                    self._load_excel_file(path)

                    self.show_excel_page()

            except Exception:

                pass

            return "break"

        for widget in (row, accent_bar, body, head, badges, details):

            widget.bind("<Double-Button-1>", open_history_item)

        for child in body.winfo_children():

            child.bind("<Double-Button-1>", open_history_item)

    self._bind_history_mousewheel_recursive()


def _render_history_ocr_view(self):

    inner = getattr(self, "history_inner", None)

    if inner is None:

        return

    for child in inner.winfo_children():

        child.destroy()

    entries = [e for e in load_history_entries() if isinstance(e, dict)]
    entries = [
        e for e in entries
        if str(e.get("action") or "").lower() in {"ocr_done", "ocr_error", "export_excel", "export_error", "phieu_coc_read", "phieu_coc_error"}
    ]
    entries.sort(key=lambda e: str(e.get("timestamp") or ""), reverse=True)
    query = str(getattr(self, "history_filter_var", tk.StringVar(value="")).get()).strip().lower()
    mode = str(getattr(self, "history_kind_var", tk.StringVar(value="Cả hai")).get()).strip()

    visible_entries = [e for e in entries if self._history_mode_matches_entry(e, mode)]

    grouped = {}
    for entry in visible_entries:
        day = str(entry.get("date") or "")[:10] or "Không rõ"
        wf_id = str(entry.get("workflow_id") or "").strip() or str(entry.get("timestamp") or "")
        kind = self._history_entry_kind(entry)
        key = (day, wf_id, kind)
        bucket = grouped.setdefault(
            key,
            {
                "date": day,
                "workflow_id": wf_id,
                "workflow_label": str(entry.get("workflow_label") or entry.get("file_name") or entry.get("file_path") or "OCR").strip(),
                "kind": kind,
                "entries": [],
                "has_error": False,
                "ocr_type": "",
                "latest_ts": str(entry.get("timestamp") or ""),
            },
        )
        bucket["entries"].append(entry)
        if str(entry.get("status") or "").lower() == "error":
            bucket["has_error"] = True
        extra = entry.get("extra") if isinstance(entry.get("extra"), dict) else {}
        if not bucket["ocr_type"]:
            bucket["ocr_type"] = str(extra.get("ocr_type") or "").strip()
        ts = str(entry.get("timestamp") or "")
        if ts and ts > bucket["latest_ts"]:
            bucket["latest_ts"] = ts

    grouped_items = sorted(grouped.values(), key=lambda g: g.get("latest_ts") or "", reverse=True)

    if query:
        filtered_items = []
        for group in grouped_items:
            haystack = " | ".join(
                [
                    str(group.get("date") or ""),
                    str(group.get("workflow_id") or ""),
                    str(group.get("workflow_label") or ""),
                    str(group.get("ocr_type") or ""),
                    *[self._history_ocr_haystack(entry) for entry in group["entries"]],
                ]
            ).lower()
            if query in haystack:
                filtered_items.append(group)
        grouped_items = filtered_items

    if not grouped_items:
        mode_text = str(mode or "Cả hai").strip()
        if mode_text.lower() in {"cả hai", "ca hai", "all"}:
            empty_title = "Chưa có lịch sử OCR."
            empty_sub = "Hãy dùng Đọc bảng hoặc Đọc phiếu cọc để tạo dữ liệu lịch sử."
            summary_text = "0 mục"
        elif self._history_mode_matches_entry({"extra": {"ocr_type": "bang_khoi_luong"}, "action": "ocr_done"}, mode_text):
            empty_title = "Chưa có lịch sử bảng khối lượng."
            empty_sub = "Hãy dùng nút Đọc bảng để OCR ảnh khối lượng, rồi mở lịch sử lại."
            summary_text = "0 khối lượng"
        else:
            empty_title = "Chưa có lịch sử phiếu cọc."
            empty_sub = "Hãy dùng nút Đọc phiếu cọc để OCR ảnh phiếu cọc, rồi mở lịch sử lại."
            summary_text = "0 phiếu cọc"
        empty = tk.Frame(inner, bg=UI_SURFACE, padx=18, pady=26, highlightthickness=1, highlightbackground="#e7edf6")
        empty.pack(fill="x")
        tk.Label(empty, text=empty_title, bg=UI_SURFACE, fg=UI_TEXT, font=ui_font(11, bold=True)).pack(anchor="center")
        tk.Label(empty, text=empty_sub, bg=UI_SURFACE, fg=UI_MUTED, font=ui_font(10)).pack(anchor="center", pady=(4, 0))
        self.history_summary_var.set(summary_text)
        self.history_selected_entry = None
        self._render_history_detail(None)
        return

    self.history_summary_var.set(f"{len(grouped_items)} mục - {len(visible_entries)} bước")
    self.history_image_refs = []

    def bind_double_click_tree(widget, handler):
        try:
            widget.bind("<Double-Button-1>", handler)
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                bind_double_click_tree(child, handler)
        except Exception:
            pass

    current_day = None
    current_day_box = None

    for group in grouped_items:
        if group["date"] != current_day:
            current_day = group["date"]
            current_day_display = self._history_display_date(current_day)
            current_day_box = tk.Frame(inner, bg=UI_SURFACE)
            current_day_box.pack(fill="x", pady=(0, 12))
            day_head = tk.Frame(current_day_box, bg=UI_SURFACE)
            day_head.pack(fill="x", pady=(0, 8))
            tk.Label(day_head, text=current_day_display, bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 11, "bold")).pack(side="left")
            tk.Label(day_head, text=f"{sum(1 for g in grouped_items if g['date'] == current_day)} OCR", bg="#eef4ff", fg=UI_PRIMARY, font=ui_font(10, bold=True), padx=10, pady=5).pack(side="right")

        status_bg, status_border, accent = self._history_status_style({"status": "error" if group["has_error"] else "success", "action": "ocr_done"})

        row = tk.Frame(current_day_box, bg=status_bg, highlightthickness=1, highlightbackground=status_border, padx=12, pady=10)
        row.pack(fill="x", pady=(0, 8))

        accent_bar = tk.Frame(row, bg=accent, width=6)
        accent_bar.pack(side="left", fill="y", padx=(0, 12))
        accent_bar.pack_propagate(False)

        thumb_box = tk.Frame(row, bg=status_bg, width=96, height=120)
        thumb_box.pack(side="left", padx=(0, 12))
        thumb_box.pack_propagate(False)

        primary_entry = next((candidate for candidate in group["entries"] if str(candidate.get("action") or "").lower() == "export_excel"), None)
        if primary_entry is None:
            primary_entry = next((candidate for candidate in group["entries"] if str(candidate.get("action") or "").lower() in {"ocr_done", "phieu_coc_read"}), None)
        if primary_entry is None:
            primary_entry = group["entries"][0] if group["entries"] else None

        thumb_source = ""
        image_candidates = []
        for candidate_entry in group["entries"]:
            if not isinstance(candidate_entry, dict):
                continue
            candidate_extra = candidate_entry.get("extra") if isinstance(candidate_entry.get("extra"), dict) else {}
            image_candidates.extend([
                candidate_extra.get("source_image_path"),
                candidate_entry.get("file_path"),
            ])
        if primary_entry:
            extra = primary_entry.get("extra") if isinstance(primary_entry.get("extra"), dict) else {}
            image_candidates = [
                extra.get("source_image_path"),
                primary_entry.get("file_path"),
            ] + image_candidates
        for candidate_path in image_candidates:
            candidate_path = str(candidate_path or "").strip()
            if not candidate_path:
                continue
            candidate_p = Path(candidate_path)
            if candidate_p.exists() and candidate_p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".jfif"}:
                thumb_source = str(candidate_p)
                break
        primary_action = str(primary_entry.get("action") or "").lower() if primary_entry else ""
        if primary_action == "export_excel" and not thumb_source:
            export_extra = primary_entry.get("extra") if isinstance(primary_entry.get("extra"), dict) else {}
            export_name = str(export_extra.get("output_name") or export_extra.get("source_excel_name") or "EXCEL").strip()
            tk.Label(thumb_box, text=f"EXCEL\n{export_name}", bg="#e5e7eb", fg=UI_TEXT, font=ui_font(11, bold=True), justify="center", wraplength=82).pack(fill="both", expand=True)
        elif thumb_source and Path(thumb_source).exists():
            try:
                img = Image.open(thumb_source).convert("RGBA")
                img.thumbnail((88, 112))
                tk_img = ImageTk.PhotoImage(img)
                self.history_image_refs.append(tk_img)
                tk.Label(thumb_box, image=tk_img, bg=status_bg).pack(fill="both", expand=True)
            except Exception:
                tk.Label(thumb_box, text="OCR", bg="#e5e7eb", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(fill="both", expand=True)
        else:
            tk.Label(thumb_box, text="OCR", bg="#e5e7eb", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(fill="both", expand=True)

        body = tk.Frame(row, bg=status_bg)
        body.pack(side="left", fill="both", expand=True)

        head = tk.Frame(body, bg=status_bg)
        head.pack(fill="x")

        kind = str(group.get("kind") or "").strip().lower()
        if kind == "bang_khoi_luong":
            title_text = f"Bảng khối lượng: {group['workflow_label']}"
        elif kind == "phieu_coc":
            title_text = f"Phiếu cọc: {group['workflow_label']}"
        else:
            title_text = f"OCR: {group['workflow_label']}"
        tk.Label(head, text=title_text, bg=status_bg, fg=UI_TEXT, font=ui_font(11, bold=True), anchor="w").pack(side="left")

        badges = tk.Frame(head, bg=status_bg)
        badges.pack(side="right")
        kind = str(group.get("kind") or "").strip().lower()
        if kind == "bang_khoi_luong":
            tk.Label(badges, text="Khối lượng", bg="#e0f2fe", fg="#0369a1", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))
        elif kind == "phieu_coc":
            tk.Label(badges, text="Phiếu cọc", bg="#ede9fe", fg="#6d28d9", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))
        tk.Label(badges, text="Đã đọc", bg="#dbeafe", fg="#1d4ed8", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left", padx=(0, 6))
        if group["has_error"]:
            tk.Label(badges, text="Lỗi", bg="#fee2e2", fg="#b91c1c", font=ui_font(10, bold=True), padx=8, pady=4).pack(side="left")

        info_line = tk.Frame(body, bg=status_bg)
        info_line.pack(fill="x", pady=(3, 0))
        tk.Label(info_line, text=f"ID: {group['workflow_id']}", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(side="left")

        summary_parts = []
        label_map = {
            "ocr_done": "Đọc OCR",
            "ocr_error": "Lỗi OCR",
            "export_excel": "Xuất Excel",
            "export_error": "Lỗi xuất",
            "phieu_coc_read": "Đọc phiếu cọc",
            "phieu_coc_error": "Lỗi phiếu cọc",
        }

        for e in group["entries"]:
            action = str(e.get("action") or "").strip()
            if action not in label_map:
                continue
            label = label_map[action]
            time_text = str(e.get("time") or "")[:8]
            parts = [time_text, label]
            if e.get("rows") not in (None, ""):
                parts.append(f"{e['rows']} dòng")
            if action == "export_excel" and e.get("file_path"):
                parts.append(f"File: {Path(str(e.get('file_path'))).name}")
                extra = e.get("extra") if isinstance(e.get("extra"), dict) else {}
                if extra.get("source_excel_name"):
                    parts.append(f"Gốc: {extra.get('source_excel_name')}")
                if extra.get("output_name"):
                    parts.append(f"Xuất: {extra.get('output_name')}")
            if e.get("message"):
                parts.append(str(e["message"]))
            extra = e.get("extra") if isinstance(e.get("extra"), dict) else {}
            if extra.get("table_count"):
                parts.append(f"{extra['table_count']} bảng")
            summary_parts.append(" - ".join(parts))

        details = tk.Frame(body, bg=status_bg)
        details.pack(fill="x", pady=(8, 0))

        for text in summary_parts[:5]:
            tk.Label(details, text=f"• {text}", bg=status_bg, fg=UI_TEXT, font=ui_font(10), anchor="w", justify="left", wraplength=900).pack(anchor="w", pady=(0, 2))

        if len(summary_parts) > 5:
            tk.Label(details, text=f"• ... và {len(summary_parts) - 5} bước khác", bg=status_bg, fg=UI_MUTED, font=ui_font(10), anchor="w").pack(anchor="w", pady=(0, 2))

        def open_history_item(_e=None, target_entry=primary_entry):
            try:
                if target_entry:
                    self.history_selected_entry = target_entry
                    self._render_history_detail(target_entry)
            except Exception:
                pass
            return "break"

        bind_double_click_tree(row, open_history_item)

    if getattr(self, "history_selected_entry", None) is None:
        self.history_selected_entry = next(
            (
                e
                for group in grouped_items
                for e in group["entries"]
                if str(e.get("action") or "").lower() in {"ocr_done", "phieu_coc_read"}
            ),
            grouped_items[0]["entries"][0] if grouped_items and grouped_items[0]["entries"] else None,
        )

    self._bind_history_mousewheel_recursive()
    self._render_history_detail(self.history_selected_entry)


def _sync_history_view(self):

    inner = getattr(self, "history_inner", None)

    if inner is None:

        return

    self._render_history_ocr_view()


def show_history_page(self, event=None):

    return self.show_page("history")


def _history_widget_contains(self, widget):
    targets = {
        getattr(self, "history_canvas", None),
        getattr(self, "history_inner", None),
        getattr(self, "history_list_panel", None),
    }
    while widget is not None:
        if widget in targets:
            return True
        widget = getattr(widget, "master", None)
    return False


def _on_history_mousewheel(self, event):
    canvas = getattr(self, "history_canvas", None)
    if canvas is None:
        return
    if not self._history_widget_contains(getattr(event, "widget", None)):
        return
    try:
        delta = getattr(event, "delta", 0) or 0
        if delta:
            canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            return
        num = getattr(event, "num", None)
        if num == 4:
            canvas.yview_scroll(-1, "units")
        elif num == 5:
            canvas.yview_scroll(1, "units")
    except Exception:
        pass


def _bind_history_mousewheel(self, event=None):
    try:
        for widget in (
            getattr(self, "history_canvas", None),
            getattr(self, "history_inner", None),
            getattr(self, "history_list_panel", None),
        ):
            if widget is None:
                continue
            widget.bind("<MouseWheel>", self._on_history_mousewheel, add="+")
            widget.bind("<Button-4>", self._on_history_mousewheel, add="+")
            widget.bind("<Button-5>", self._on_history_mousewheel, add="+")
    except Exception:
        pass


def _unbind_history_mousewheel(self, event=None):
    try:
        for widget in (
            getattr(self, "history_canvas", None),
            getattr(self, "history_inner", None),
            getattr(self, "history_list_panel", None),
        ):
            if widget is None:
                continue
            widget.unbind("<MouseWheel>")
            widget.unbind("<Button-4>")
            widget.unbind("<Button-5>")
    except Exception:
        pass


def _bind_history_mousewheel_recursive(self, widget=None):
    widget = widget or getattr(self, "history_inner", None)
    if widget is None:
        return
    try:
        widget.bind("<MouseWheel>", self._on_history_mousewheel, add="+")
        widget.bind("<Button-4>", self._on_history_mousewheel, add="+")
        widget.bind("<Button-5>", self._on_history_mousewheel, add="+")
    except Exception:
        pass
    try:
        for child in widget.winfo_children():
            self._bind_history_mousewheel_recursive(child)
    except Exception:
        pass


def install_history_ui(app_cls):
    app_cls._history_status_style = _history_status_style
    app_cls._history_ocr_haystack = _history_ocr_haystack
    app_cls._history_is_khoi_luong_entry = _history_is_khoi_luong_entry
    app_cls._history_entry_kind = _history_entry_kind
    app_cls._history_entry_kind_label = _history_entry_kind_label
    app_cls._history_mode_matches_entry = _history_mode_matches_entry
    app_cls._history_set_detail_text = _history_set_detail_text
    app_cls._history_format_cell = _history_format_cell
    app_cls._history_display_date = _history_display_date
    app_cls._history_format_row = _history_format_row
    app_cls._history_make_box_line = _history_make_box_line
    app_cls._history_norm_col = _history_norm_col
    app_cls._history_make_row_map = _history_make_row_map
    app_cls._history_pick_field = _history_pick_field
    app_cls._history_render_metric_box = _history_render_metric_box
    app_cls._history_pretty_column_label = _history_pretty_column_label
    app_cls._history_score_table = _history_score_table
    app_cls._history_pick_main_table = _history_pick_main_table
    app_cls._history_clear_detail_host = _history_clear_detail_host
    app_cls._history_build_detail_grid = _history_build_detail_grid
    app_cls._history_render_table_preview = _history_render_table_preview
    app_cls._history_load_tables_for_entry = _history_load_tables_for_entry
    app_cls._render_history_detail = _render_history_detail
    app_cls._render_history_view = _render_history_view
    app_cls._render_history_workflow_view = _render_history_workflow_view
    app_cls._render_history_ocr_view = _render_history_ocr_view
    app_cls._sync_history_view = _sync_history_view
    app_cls.show_history_page = show_history_page
    app_cls._history_widget_contains = _history_widget_contains
    app_cls._on_history_mousewheel = _on_history_mousewheel
    app_cls._bind_history_mousewheel = _bind_history_mousewheel
    app_cls._unbind_history_mousewheel = _unbind_history_mousewheel
    app_cls._bind_history_mousewheel_recursive = _bind_history_mousewheel_recursive
