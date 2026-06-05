"""OCR workflow, image preview, viewer, and clipboard actions for the main window."""

import json
import math
import os
import subprocess
import shutil
import tempfile
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from gk_pilepro.gk_core import last_run_dir, new_workflow_id, write_role_error_log
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

        messagebox.showerror("Lỗi đọc phiếu cọc", "Có lỗi. Xem last_run_v12/last_error_phieu_coc.txt")

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


def open_image_viewer(self, path):

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

                state["win"].lift()

                state["win"].focus_force()

                if callable(getattr(self, "_viewer_set_index", None)):

                    self._viewer_set_index(start_index)

                return

        except Exception:

            pass

    win = tk.Toplevel(self.root)
    win.title(f"Xem ảnh - {Path(path).name}")
    win.configure(bg=UI_BG)
    win.transient(self.root)

    try:
        win.grab_set()
    except Exception:
        pass

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
    canvas.image = None

    state = {
        "win": win,
        "canvas": canvas,
        "image_id": image_id,
        "title_var": title_var,
        "prev_btn_id": prev_btn_id,
        "next_btn_id": next_btn_id,
        "paths": paths,
        "index": start_index,
        "sw": sw,
        "sh": sh,
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

    def _load_current_image():
        current_paths = list(state.get("paths") or [])
        if not current_paths:
            return
        idx = max(0, min(int(state.get("index") or 0), len(current_paths) - 1))
        state["index"] = idx
        current_path = current_paths[idx]
        try:
            img = Image.open(current_path)
        except Exception as e:
            messagebox.showerror("Xem ảnh", f"Không mở được ảnh:\n{e}")
            return
        img_copy = img.copy()
        max_w = max(800, int(state["sw"]) - 120)
        max_h = max(600, int(state["sh"]) - 180)
        img_copy.thumbnail((max_w, max_h))
        photo = ImageTk.PhotoImage(img_copy)
        canvas.image = photo
        canvas.itemconfigure(image_id, image=photo)
        title_var.set(f"{Path(current_path).name} ({idx + 1}/{len(current_paths)})")
        self.preview_image_index = idx
        self.image_path = current_path
        self._update_preview_counter()
        _sync_image_position()

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

            tables_one = postprocess_to_hop_coc_d1_d2(tables_one)

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

        messagebox.showerror("Lỗi đọc ảnh", "Có lỗi. Xem last_run_v12/last_error.txt")

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
    app_cls.get_current_preview_image_path = get_current_preview_image_path
    app_cls.open_current_preview_image = open_current_preview_image
    app_cls.set_image_path = set_image_path
    app_cls.set_image_paths = set_image_paths
    app_cls.append_image_paths = append_image_paths
    app_cls.open_image_viewer = open_image_viewer
    app_cls._reset_current_preview_state = _reset_current_preview_state
    app_cls._save_history_image_snapshot = _save_history_image_snapshot
    app_cls.paste_image_from_clipboard = paste_image_from_clipboard
    app_cls.run_gemini = run_gemini
