"""Admin log and approval panels for the main application window."""

import re
import threading
import tkinter as tk
import unicodedata
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from gk_pilepro.gk_core import (
    delete_admin_approved_machine,
    fetch_presence_error_logs,
    format_machine_last_seen,
    import_local_approval_to_admin_list,
    is_admin_build,
    is_machine_active_recently,
    list_backup_files,
    load_admin_approved_machines,
    parse_machine_datetime,
    presence_server_url_from_env,
    remember_admin_approved_machine,
    restore_backup_file,
    resolve_presence_error_log,
    resolve_presence_server_url,
    resource_path,
    sync_presence_machines_to_admin_list,
)
from gk_pilepro.ui.gk_ui import (
    UI_BG,
    UI_BORDER,
    UI_ERROR,
    UI_MUTED,
    UI_PRIMARY,
    UI_PRIMARY_ACTIVE,
    UI_SUCCESS,
    UI_SURFACE,
    UI_SURFACE_2,
    UI_TEXT,
    RoundedMappingEntry,
    ui_button,
    ui_font,
)


def norm(text):
    value = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", value).strip().lower()


SIDEBAR_BG = "#053f32"
SIDEBAR_CARD = "#0b5a45"
SIDEBAR_ACTIVE = "#0f8d6d"
SIDEBAR_TEXT = "#d8f3e8"
SIDEBAR_MUTED = "#9fcec0"
MAIN_BG = "#edf7f8"


def _admin_ui_icon(self, filename, size=22):
    try:
        return self._ui_icon(filename, size)
    except Exception:
        return None


def _admin_img_label(parent, image, bg, **kwargs):
    if image is None:
        return tk.Label(parent, bg=bg, **kwargs)
    return tk.Label(parent, image=image, bg=bg, **kwargs)


def _admin_log_badge_loop(self):
    if not is_admin_build():
        return

    def _worker():
        rows = fetch_presence_error_logs(presence_server_url_from_env(), limit=300, timeout=2, unresolved_only=True)

        def _apply():
            try:
                badge_canvas = getattr(self, "admin_log_notify_canvas", None)
                badge_oval = getattr(self, "admin_log_badge_oval", None)
                badge_text = getattr(self, "admin_log_badge_text", None)
                if badge_canvas is not None and badge_canvas.winfo_exists() and badge_oval is not None and badge_text is not None:
                    count = len(rows)
                    if count:
                        count_text = "99+" if count > 99 else str(count)
                        badge_canvas.itemconfigure(badge_oval, state="normal")
                        badge_canvas.itemconfigure(badge_text, text=count_text, state="normal")
                    else:
                        badge_canvas.itemconfigure(badge_oval, state="hidden")
                        badge_canvas.itemconfigure(badge_text, state="hidden")
            except Exception:
                pass
            try:
                self.root.after(5000, self._admin_log_badge_loop)
            except Exception:
                pass

        try:
            self.root.after(0, _apply)
        except Exception:
            pass

    threading.Thread(target=_worker, daemon=True).start()


def open_admin_log_panel(self):
    if not is_admin_build():
        return
    if getattr(self, "admin_log_panel", None) is not None:
        try:
            if self.admin_log_panel.winfo_exists():
                self.admin_log_panel.lift()
                self.admin_log_panel.focus_force()
                return
        except Exception:
            pass

    win = tk.Toplevel(self.root)
    win.transient(self.root)
    win.title("Thông báo log")
    win.configure(bg="#f3f6fb")
    self.admin_log_panel = win
    win.bind("<Destroy>", lambda e: setattr(self, "admin_log_panel", None) if e.widget is win else None, add="+")
    self._fit_dialog_to_screen(win, 920, 620, min_w=820, min_h=520, max_ratio=0.84, lock_size=False)

    body = tk.Frame(win, bg="#f3f6fb", padx=18, pady=18)
    body.pack(fill="both", expand=True)
    header = tk.Frame(body, bg="#f3f6fb")
    header.pack(fill="x", pady=(0, 12))
    title_box = tk.Frame(header, bg="#f3f6fb")
    title_box.pack(side="left", fill="x", expand=True)
    tk.Label(title_box, text="Thông báo log lỗi", bg="#f3f6fb", fg=UI_TEXT, font=ui_font(16, bold=True)).pack(anchor="w")
    summary_var = tk.StringVar(value="Đang tải...")
    tk.Label(title_box, textvariable=summary_var, bg="#f3f6fb", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(2, 0))

    table_card = tk.Frame(body, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3")
    table_card.pack(fill="both", expand=True)
    table_frame = tk.Frame(table_card, bg="#ffffff", padx=12, pady=12)
    table_frame.pack(fill="both", expand=True)
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)

    style = ttk.Style()
    style.configure("Log.Treeview", background="#ffffff", fieldbackground="#ffffff", rowheight=34, borderwidth=0, font=ui_font(10))
    style.configure("Log.Treeview.Heading", font=ui_font(10, bold=True), background="#eff6ff", foreground="#0f172a", relief="flat", borderwidth=0)
    log_tree = ttk.Treeview(
        table_frame,
        style="Log.Treeview",
        columns=("state", "time", "user", "windows", "computer", "message"),
        show="headings",
        height=9,
    )
    for col, text, width, anchor in (
        ("state", "TT", 58, "center"),
        ("time", "Thời gian", 158, "center"),
        ("user", "Tên người", 130, "center"),
        ("windows", "User Windows", 126, "center"),
        ("computer", "Tên máy", 140, "center"),
        ("message", "Nội dung", 300, "w"),
    ):
        log_tree.heading(col, text=text)
        log_tree.column(col, width=width, anchor=anchor, stretch=(col == "message"))
    log_tree.tag_configure("new", foreground=UI_ERROR, background="#fff7f7")
    log_tree.tag_configure("done", foreground="#64748b", background="#f8fafc")
    scroll = ttk.Scrollbar(table_frame, orient="vertical", command=log_tree.yview)
    log_tree.configure(yscrollcommand=scroll.set)
    log_tree.grid(row=0, column=0, sticky="nsew")
    scroll.grid(row=0, column=1, sticky="ns")

    rows_cache = {"rows": []}

    def selected_row():
        selected = log_tree.selection()
        if not selected:
            return None
        sid = str(selected[0])
        for item in rows_cache["rows"]:
            if str(item.get("id") or "") == sid:
                return item
        return None

    def detail_text(row):
        log_text = str(row.get('log_text', ''))
        import re
        blocks = re.split(r"={10,}", log_text)
        error_blocks = []
        for b in blocks:
            b_strip = b.strip()
            if not b_strip:
                continue
            if "NoneType: None" in b_strip and "Traceback" not in b_strip and "error:" not in b_strip.lower() and "Exception" not in b_strip:
                continue
            error_blocks.append(b_strip)
            
        if error_blocks:
            display_log = "\n" + ("=" * 80) + "\n".join(["\n" + b + "\n" + ("=" * 80) for b in error_blocks])
        else:
            display_log = "Hiện tại không có lỗi."

        return (
            f"Thời gian: {row.get('created_at', '')}\n"
            f"Mã máy: {row.get('machine_code', '')}\n"
            f"Tên người: {row.get('user_name', '')}\n"
            f"User Windows: {row.get('windows_user', '')}\n"
            f"Tên máy Windows: {row.get('computer_name', '')}\n"
            f"Vai trò: {row.get('role', '')}\n"
            f"Trạng thái: {'Đã hoàn thành' if row.get('resolved_at') else 'Chưa xử lý'}\n\n"
            f"Kiểm tra log: {'Hợp lệ' if row.get('hash_ok') else 'Có dấu hiệu bị sửa hoặc log cũ chưa có hash'}\n"
            f"Mã hash: {row.get('log_hash', '')}\n\n"
            f"Nội dung log:\n{display_log}"
        )

    def show_detail():
        row = selected_row()
        if not row:
            messagebox.showinfo("Thông báo log", "Chọn một dòng log trước.")
            return
        top = tk.Toplevel(win)
        top.transient(win)
        top.title("Chi tiết log lỗi")
        top.configure(bg="#f3f6fb")
        self._fit_dialog_to_screen(top, 860, 600, min_w=780, min_h=500, max_ratio=0.82, lock_size=False)
        wrap = tk.Frame(top, bg="#f3f6fb", padx=18, pady=18)
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Chi tiết log lỗi", bg="#f3f6fb", fg=UI_TEXT, font=ui_font(16, bold=True)).pack(anchor="w")
        tk.Label(wrap, text=f"{row.get('user_name', '')} - {row.get('computer_name', '')}", bg="#f3f6fb", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(2, 12))
        text = tk.Text(wrap, wrap="word", bg="#0f172a", fg="#e5edf7", insertbackground="#e5edf7", relief="flat", bd=0, font=("Consolas", 10), padx=12, pady=10)
        text.pack(fill="both", expand=True)
        text.insert("1.0", detail_text(row))
        text.configure(state="disabled")
        actions = tk.Frame(wrap, bg="#f3f6fb")
        actions.pack(fill="x", pady=(12, 0))
        ui_button(actions, "Copy log", lambda: (self.root.clipboard_clear(), self.root.clipboard_append(detail_text(row))), width=10, variant="soft").pack(side="left")
        ui_button(actions, "Đóng", top.destroy, width=8, variant="default").pack(side="right")
        self._center_dialog_on_screen(top)

    def copy_log():
        row = selected_row()
        if not row:
            messagebox.showinfo("Thông báo log", "Chọn một dòng log trước.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(detail_text(row))

    def mark_done():
        row = selected_row()
        if not row:
            messagebox.showinfo("Thông báo log", "Chọn một dòng log trước.")
            return
        if row.get("resolved_at"):
            return
        if resolve_presence_error_log(presence_server_url_from_env(), row.get("id"), timeout=3):
            self._set_status("Đã hoàn thành log lỗi.", "success")
            refresh_logs()
            self._admin_log_badge_loop()
        else:
            messagebox.showerror("Thông báo log", "Không đánh dấu hoàn thành được log này.")

    def refresh_logs():
        summary_var.set("Đang tải...")

        def _worker():
            rows = fetch_presence_error_logs(presence_server_url_from_env(), limit=200, timeout=3)

            def _apply():
                try:
                    rows_cache["rows"] = rows
                    log_tree.delete(*log_tree.get_children())
                    unresolved = 0
                    for row in rows:
                        done = bool(row.get("resolved_at"))
                        if not done:
                            unresolved += 1
                        iid = str(row.get("id") or "")
                        if not iid:
                            continue
                        log_tree.insert(
                            "",
                            "end",
                            iid=iid,
                            values=(
                                "Mới" if not done else "Xong",
                                row.get("created_at", ""),
                                row.get("user_name", ""),
                                row.get("windows_user", ""),
                                row.get("computer_name", ""),
                                str(row.get("message") or row.get("log_text") or "")[:120],
                            ),
                            tags=("done" if done else "new",),
                        )
                    summary_var.set(f"{unresolved} thông báo chưa xử lý / {len(rows)} log gần nhất.")
                except Exception:
                    summary_var.set("Không tải được log.")

            try:
                self.root.after(0, _apply)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    log_tree.bind("<Double-1>", lambda _e: show_detail())
    actions = tk.Frame(body, bg="#f3f6fb")
    actions.pack(fill="x", pady=(12, 0))
    ui_button(actions, "Tải lại", refresh_logs, width=9, variant="soft").pack(side="left")
    ui_button(actions, "Xem chi tiết", show_detail, width=12, variant="primary").pack(side="left", padx=(8, 0))
    ui_button(actions, "Hoàn thành", mark_done, width=11, variant="success").pack(side="left", padx=(8, 0))
    ui_button(actions, "Copy log", copy_log, width=10, variant="default").pack(side="left", padx=(8, 0))
    ui_button(actions, "Đóng", win.destroy, width=8, variant="default").pack(side="right")

    refresh_logs()
    self._center_dialog_on_screen(win)


def open_admin_backup_panel(self):
    if not is_admin_build():
        return
    win = tk.Toplevel(self.root)
    win.transient(self.root)
    win.title("Backup Excel")
    win.configure(bg="#f3f6fb")
    self._fit_dialog_to_screen(win, 960, 640, min_w=840, min_h=600, max_ratio=0.9, lock_size=False)

    body = tk.Frame(win, bg="#f3f6fb", padx=20, pady=18)
    body.pack(fill="both", expand=True)
    tk.Label(body, text="Backup Excel", bg="#f3f6fb", fg=UI_TEXT, font=ui_font(16, bold=True)).pack(anchor="w")
    summary_var = tk.StringVar(value="Chọn backup rồi chọn file Excel cần khôi phục.")
    tk.Label(body, textvariable=summary_var, bg="#f3f6fb", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(2, 10))

    table_card = tk.Frame(body, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3")
    table_card.pack(fill="x", expand=False)
    table_card.configure(height=390)
    table_card.pack_propagate(False)
    table_frame = tk.Frame(table_card, bg="#ffffff", padx=12, pady=12)
    table_frame.pack(fill="both", expand=True)
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)

    tree = ttk.Treeview(table_frame, columns=("time", "name", "size", "path"), show="headings", height=9)
    for col, text, width, anchor in (
        ("time", "Thời gian", 160, "center"),
        ("name", "Tên backup", 330, "w"),
        ("size", "Dung lượng", 110, "center"),
        ("path", "Vị trí lưu", 250, "w"),
    ):
        tree.heading(col, text=text)
        tree.column(col, width=width, anchor=anchor, stretch=(col in ("name", "path")))
    scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    scroll.grid(row=0, column=1, sticky="ns")
    rows_cache = {"rows": []}

    detail_card = tk.Frame(body, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3")
    detail_card.pack(fill="x", pady=(10, 0))
    detail_card.configure(height=82)
    detail_card.pack_propagate(False)
    detail_inner = tk.Frame(detail_card, bg="#ffffff", padx=12, pady=10)
    detail_inner.pack(fill="x")
    detail_var = tk.StringVar(value="Chưa chọn backup.")
    tk.Label(detail_inner, text="Chi tiết backup", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
    tk.Label(detail_inner, textvariable=detail_var, bg="#ffffff", fg=UI_MUTED, font=ui_font(9), justify="left", wraplength=860).pack(anchor="w", pady=(4, 0))

    def format_size(size):
        try:
            value = int(size or 0)
        except Exception:
            return str(size or "")
        if value >= 1024 * 1024:
            return f"{value / (1024 * 1024):.1f} MB"
        if value >= 1024:
            return f"{value / 1024:.1f} KB"
        return f"{value} B"

    def short_path(path):
        path = str(path or "")
        if len(path) <= 52:
            return path
        parts = re.split(r"[\\/]+", path)
        if len(parts) >= 3:
            return f"{parts[0]}\\...\\{parts[-2]}\\{parts[-1]}"
        return f"...{path[-49:]}"

    def selected_backup():
        selected = tree.selection()
        if not selected:
            return None
        idx = int(selected[0])
        rows = rows_cache.get("rows") or []
        return rows[idx] if 0 <= idx < len(rows) else None

    def refresh():
        rows = list_backup_files("excel", limit=300)
        rows_cache["rows"] = rows
        tree.delete(*tree.get_children())
        for idx, row in enumerate(rows):
            tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(row.get("modified_at", ""), row.get("name", ""), format_size(row.get("size")), short_path(row.get("path"))),
            )
        summary_var.set(f"{len(rows)} backup Excel gần nhất." if rows else "Chưa có backup Excel.")
        detail_var.set("Chưa chọn backup.")

    def update_detail(*_args):
        row = selected_backup()
        if not row:
            detail_var.set("Chưa chọn backup.")
            return
        detail_var.set(
            f"Tên: {row.get('name', '')}\n"
            f"Dung lượng: {format_size(row.get('size'))}\n"
            f"Đường dẫn: {row.get('path', '')}"
        )

    def restore_selected():
        row = selected_backup()
        if not row:
            messagebox.showinfo("Backup Excel", "Chọn một backup trước.")
            return
        target = filedialog.askopenfilename(title="Chọn file Excel cần khôi phục", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not target:
            return
        if not messagebox.askyesno("Khôi phục backup", f"Khôi phục backup này vào file Excel đã chọn?\n\nBackup: {row.get('name')}\nFile đích: {target}\n\nApp sẽ tạo thêm một backup trước khi ghi đè."):
            return
        try:
            pre = restore_backup_file(row.get("path"), target)
            self._set_status("Đã khôi phục backup Excel.", "success")
            messagebox.showinfo("Backup Excel", f"Đã khôi phục backup.\nBackup trước khi ghi đè: {pre or 'Không có'}")
        except Exception as exc:
            messagebox.showerror("Backup Excel", f"Không khôi phục được backup:\n{exc}")

    def show_detail():
        row = selected_backup()
        if not row:
            messagebox.showinfo("Backup Excel", "Chọn một backup trước.")
            return
        top = tk.Toplevel(win)
        top.title("Chi tiết backup")
        top.transient(win)
        top.configure(bg="#f3f6fb")
        self._fit_dialog_to_screen(top, 660, 420, min_w=500, min_h=300, max_ratio=0.8, lock_size=False)
        wrap = tk.Frame(top, bg="#f3f6fb", padx=18, pady=18)
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Chi tiết backup", bg="#f3f6fb", fg=UI_TEXT, font=ui_font(16, bold=True)).pack(anchor="w")
        tk.Label(wrap, text=f"{row.get('name', '')}", bg="#f3f6fb", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(2, 12))
        text = tk.Text(wrap, wrap="word", bg="#0f172a", fg="#e5edf7", insertbackground="#e5edf7", relief="flat", bd=0, font=("Consolas", 10), padx=12, pady=10)
        text.pack(fill="both", expand=True)
        content = (
            f"Tên backup: {row.get('name', '')}\n"
            f"Thời gian tạo: {row.get('modified_at', '')}\n"
            f"Dung lượng: {format_size(row.get('size'))}\n\n"
            f"Vị trí lưu đầy đủ:\n{row.get('path', '')}"
        )
        text.insert("1.0", content)
        text.configure(state="disabled")
        actions_top = tk.Frame(wrap, bg="#f3f6fb")
        actions_top.pack(fill="x", pady=(12, 0))
        ui_button(actions_top, "Copy đường dẫn", lambda: (self.root.clipboard_clear(), self.root.clipboard_append(row.get('path', ''))), width=16, variant="soft").pack(side="left")
        ui_button(actions_top, "Đóng", top.destroy, width=8, variant="default").pack(side="right")
        self._center_dialog_on_screen(top)

    actions = tk.Frame(body, bg="#f3f6fb")
    actions.pack(fill="x", pady=(12, 0))
    ui_button(actions, "Tải lại", refresh, width=9, variant="soft").pack(side="left")
    ui_button(actions, "Xem chi tiết", show_detail, width=12, variant="primary").pack(side="left", padx=(8, 0))
    ui_button(actions, "Khôi phục", restore_selected, width=11, variant="success").pack(side="left", padx=(8, 0))
    ui_button(actions, "Đóng", win.destroy, width=8, variant="default").pack(side="right")
    tree.bind("<<TreeviewSelect>>", update_detail)
    refresh()
    self._center_dialog_on_screen(win)


def open_admin_approval_panel(self):
    if self.admin_approval_panel is not None:
        try:
            self.admin_approval_panel.lift()
            self.admin_approval_panel.focus_set()
        except Exception:
            pass
        return

    panel = tk.Frame(self.root, bg=MAIN_BG)
    panel.place(relx=0, rely=0, relwidth=1, relheight=1)
    self.admin_approval_panel = panel

    refresh_job = {"id": None}

    def close_panel(*args):
        try:
            if refresh_job["id"] is not None:
                try:
                    panel.after_cancel(refresh_job["id"])
                except Exception:
                    pass
            panel.destroy()
        finally:
            self.admin_approval_panel = None

    sidebar_w = 232
    sidebar = tk.Frame(panel, width=sidebar_w, bg=SIDEBAR_BG, highlightthickness=1, highlightbackground=SIDEBAR_CARD)
    sidebar.pack(side="left", fill="y")
    sidebar.pack_propagate(False)

    brand = tk.Frame(sidebar, bg=SIDEBAR_BG)
    brand.pack(fill="x", pady=(0, 16))
    try:
        if hasattr(self, 'app_logo_img') and self.app_logo_img is not None:
            tk.Label(brand, image=self.app_logo_img, bg=SIDEBAR_BG).pack(anchor="center", pady=(14, 0))
    except Exception:
        pass

    nav_items = [
        ("home", "04_sidebar_home.png", "Trang chủ", True),
        ("excel", "05_sidebar_excel.png", "Excel", False),
        ("history", "06_sidebar_history.png", "Lịch sử", False),
        ("mapping", "07_sidebar_mapping.png", "Mẫu mapping", False),
        ("settings", "08_sidebar_settings.png", "Cài đặt", False),
        ("help", "09_sidebar_help.png", "Trợ giúp", False),
        ("about", "10_sidebar_info.png", "Giới thiệu", False),
    ]

    for page_id, icon, text, active in nav_items:
        bg = SIDEBAR_ACTIVE if active else SIDEBAR_BG
        fg = "#ffffff" if active else SIDEBAR_TEXT
        border = SIDEBAR_ACTIVE if active else SIDEBAR_CARD
        row = tk.Frame(sidebar, bg=bg, padx=0, pady=0, highlightthickness=1, highlightbackground=border, cursor="hand2")
        row.pack(fill="x", padx=16, pady=2)
        inner = tk.Frame(row, bg=bg, padx=12, pady=7, cursor="hand2")
        inner.pack(fill="both", expand=True)
        icon_img = _admin_ui_icon(self, icon, 20)
        _admin_img_label(inner, icon_img, bg, width=24, cursor="hand2").pack(side="left", padx=(0, 9))
        lbl = tk.Label(inner, text=text, font=ui_font(11, bold=active), bg=bg, fg=fg, anchor="w", cursor="hand2")
        lbl.pack(side="left", fill="x", expand=True)
        for widget in (row, inner, lbl):
            widget.bind("<Button-1>", lambda e: close_panel())

    # Sidebar bottom section - keep the admin block fixed and visually identical.
    member_info = tk.Frame(
        sidebar,
        bg=SIDEBAR_CARD,
        padx=10,
        pady=12,
        highlightthickness=1,
        highlightbackground="#19765d",
    )
    member_info.pack(side="bottom", fill="x", padx=12, pady=(0, 14))
    tk.Label(member_info, text="Quản trị viên", font=ui_font(10, bold=True), bg=SIDEBAR_CARD, fg="#ffffff").pack(anchor="center")
    tk.Label(member_info, text="Admin", font=ui_font(9), bg=SIDEBAR_CARD, fg=SIDEBAR_MUTED).pack(anchor="center", pady=(4, 10))
    ui_button(member_info, "Duyệt máy", lambda: None, width=14, variant="warn").pack(anchor="center")
    ui_button(member_info, "Backup Excel", self.open_admin_backup_panel, width=14, variant="soft").pack(anchor="center", pady=(8, 0))

    # Stats box in sidebar (quick glance)
    stat_total_var = tk.StringVar(value="Đang tải...")
    stat_time_var = tk.StringVar(value="Cập nhật: --:--:--")
    stats_box = tk.Frame(sidebar, bg=SIDEBAR_CARD, highlightthickness=1, highlightbackground="#19765d")
    stats_box.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
    tk.Label(stats_box, text="Kết nối server ●", bg=SIDEBAR_CARD, fg="#ffffff", font=ui_font(9, bold=True), anchor="w").pack(padx=10, pady=(8, 0), anchor="w")
    tk.Label(stats_box, textvariable=stat_total_var, bg=SIDEBAR_CARD, fg="#30d083", font=ui_font(9, bold=True), anchor="w").pack(padx=10, pady=(4, 0), anchor="w")
    tk.Label(stats_box, textvariable=stat_time_var, bg=SIDEBAR_CARD, fg=SIDEBAR_MUTED, font=ui_font(8), anchor="w").pack(padx=10, pady=(2, 8), anchor="w")
    try:
        side_decor = self._ui_asset_image("assets/gk_sidebar_decoration.png", (210, 160), alpha=0.9)
        if side_decor is not None:
            tk.Label(sidebar, image=side_decor, bg=SIDEBAR_BG, bd=0).pack(side="bottom", fill="x", pady=(8, 0))
    except Exception:
        pass

    stat_active_var = tk.StringVar(value="0")
    stat_pending_var = tk.StringVar(value="0")
    stat_blocked_var = tk.StringVar(value="0")
    stat_total_num_var = tk.StringVar(value="0")

    # Main content mirrors the 24-inch admin approval workspace.
    main_content = tk.Frame(panel, bg=MAIN_BG)
    main_content.pack(side="left", fill="both", expand=True, padx=(14, 16), pady=12)

    header_frame = tk.Frame(main_content, bg="#f8fcff", padx=14, pady=12, highlightthickness=1, highlightbackground="#dbe8f1")
    header_frame.pack(fill="x", pady=(0, 10))
    
    title_row = tk.Frame(header_frame, bg="#f8fcff")
    title_row.pack(fill="x")
    icon_box = tk.Frame(title_row, bg="#eefaf4", width=46, height=46)
    icon_box.pack(side="left", padx=(0, 12))
    icon_box.pack_propagate(False)
    avatar_title = _admin_ui_icon(self, "12_header_avatar.png", 24)
    _admin_img_label(icon_box, avatar_title, "#eefaf4", text="A", fg=UI_PRIMARY, font=ui_font(16, bold=True)).pack(expand=True)
    tk.Label(title_row, text="Duyệt máy thành viên", bg="#f8fcff", fg="#0f2f2a", font=ui_font(16, bold=True)).pack(side="left")
    profile = tk.Frame(title_row, bg="#f8fcff")
    profile.pack(side="right")
    bell = _admin_ui_icon(self, "11_header_notification.png", 18)
    _admin_img_label(profile, bell, "#f8fcff").pack(side="left", padx=(0, 16))
    avatar = _admin_ui_icon(self, "12_header_avatar.png", 36)
    _admin_img_label(profile, avatar, "#f8fcff", text="A", fg="#ffffff", font=ui_font(12, bold=True)).pack(side="left", padx=(0, 10))
    user_box = tk.Frame(profile, bg="#f8fcff")
    user_box.pack(side="left")
    tk.Label(user_box, text="Quản trị viên", bg="#f8fcff", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w")
    tk.Label(user_box, text="Admin", bg="#f8fcff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
    drop = _admin_ui_icon(self, "13_header_dropdown.png", 14)
    _admin_img_label(profile, drop, "#f8fcff").pack(side="left", padx=(10, 0))
    
    tk.Label(header_frame, text="Nhập mã máy do thành viên cung cấp để tạo mã duyệt cho họ.", bg="#f8fcff", fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(2, 0), padx=(58, 0))

    approval_body = tk.Frame(main_content, bg=MAIN_BG)
    approval_body.pack(fill="both", expand=True)

    footer = tk.Frame(main_content, bg="#f8fcff", padx=12, pady=7, highlightthickness=1, highlightbackground="#dbe8f1", highlightcolor="#dbe8f1")
    footer.pack(fill="x", side="bottom", pady=(8, 0))
    try:
        footer_server_var = getattr(self, "footer_server_var", tk.StringVar(value=self._footer_server_state()))
        footer_status_var = getattr(self, "footer_status_var", tk.StringVar(value=getattr(self, "_last_status_text", "Sẵn sàng")))
        footer_time_var = getattr(self, "footer_time_var", tk.StringVar(value=datetime.now().strftime("%H:%M:%S")))
        footer_date_var = getattr(self, "footer_date_var", tk.StringVar(value=datetime.now().strftime("%d/%m/%Y")))
    except Exception:
        footer_server_var = tk.StringVar(value="Đang kiểm tra")
        footer_status_var = tk.StringVar(value="Sẵn sàng")
        footer_time_var = tk.StringVar(value=datetime.now().strftime("%H:%M:%S"))
        footer_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))

    def footer_item(label, variable=None, text=None, dot=False):
        item = tk.Frame(footer, bg="#f8fcff")
        item.pack(side="left", padx=(0, 20))
        if dot:
            tk.Label(item, text="●", bg="#f8fcff", fg=UI_SUCCESS, font=ui_font(9, bold=True)).pack(side="left", padx=(0, 5))
        tk.Label(item, text=f"{label}:", bg="#f8fcff", fg=UI_MUTED, font=ui_font(9, bold=True)).pack(side="left")
        if variable is not None:
            tk.Label(item, textvariable=variable, bg="#f8fcff", fg=UI_TEXT, font=ui_font(9)).pack(side="left", padx=(4, 0))
        else:
            tk.Label(item, text=text or "", bg="#f8fcff", fg=UI_TEXT, font=ui_font(9)).pack(side="left", padx=(4, 0))

    footer_item("Phiên bản", text="1.0.0")
    footer_item("Máy chủ", variable=footer_server_var, dot=True)
    footer_item("Trạng thái", variable=footer_status_var, dot=True)
    footer_item("Thời gian", variable=footer_time_var)
    footer_item("Ngày", variable=footer_date_var)
    try:
        admin_decor = self._ui_asset_image("assets/gk_footer_decoration.png", (260, 96), alpha=0.62)
        if admin_decor is not None:
            tk.Label(footer, image=admin_decor, bg="#f8fcff", bd=0).pack(side="right", padx=(12, 0))
    except Exception:
        pass

    left_area = tk.Frame(approval_body, bg=MAIN_BG)
    left_area.pack(side="left", fill="both", expand=True)

    overview_panel = tk.Frame(
        approval_body,
        width=282,
        bg="#ffffff",
        highlightthickness=1,
        highlightbackground="#d7e5e0",
        highlightcolor="#d7e5e0",
        padx=16,
        pady=16,
    )
    overview_panel.pack(side="right", fill="y", padx=(14, 0))
    overview_panel.pack_propagate(False)

    tk.Label(overview_panel, text="Tổng quan hệ thống", bg="#ffffff", fg=UI_TEXT, font=ui_font(12, bold=True)).pack(anchor="w", pady=(0, 14))

    def make_stat_tile(parent, icon, label, var, color, bg):
        tile = tk.Frame(parent, bg="#ffffff")
        tile.pack(fill="x", pady=(0, 10))
        badge = tk.Canvas(tile, bg="#ffffff", width=46, height=46, bd=0, highlightthickness=0)
        badge.pack(side="left", padx=(0, 12))
        badge.create_round_rect = lambda x1, y1, x2, y2, radius=10, **kwargs: badge.create_polygon(
            [
                x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
                x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
                x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
            ],
            smooth=True,
            splinesteps=24,
            **kwargs,
        )
        badge.create_round_rect(1, 1, 45, 45, radius=12, fill=bg, outline=bg)
        badge.create_text(23, 23, text=icon, fill=color, font=ui_font(15, bold=True))
        text_box = tk.Frame(tile, bg="#ffffff")
        text_box.pack(side="left", fill="x", expand=True)
        tk.Label(text_box, text=label, bg="#ffffff", fg=UI_TEXT, font=ui_font(10)).pack(anchor="w")
        tk.Label(text_box, textvariable=var, bg="#ffffff", fg=color, font=ui_font(17, bold=True)).pack(anchor="w")
        sep = tk.Frame(parent, bg="#eef2f7", height=1)
        sep.pack(fill="x", pady=(0, 10))

    make_stat_tile(overview_panel, "▣", "Máy đã duyệt", stat_total_num_var, UI_PRIMARY, "#eaf2ff")
    make_stat_tile(overview_panel, "⌾", "Đang hoạt động", stat_active_var, UI_SUCCESS, "#dcfce7")
    make_stat_tile(overview_panel, "▢", "Chờ duyệt", stat_pending_var, "#f59e0b", "#fff7ed")
    make_stat_tile(overview_panel, "⊘", "Đã chặn", stat_blocked_var, UI_ERROR, "#fee2e2")
    tk.Label(
        overview_panel,
        textvariable=stat_time_var,
        bg="#ffffff",
        fg=UI_TEXT,
        font=ui_font(8),
        anchor="w",
        justify="left",
        wraplength=248,
    ).pack(side="bottom", fill="x")

    def admin_button(parent, text, command, width=12, variant="default"):
        palette = {
            "default": ("#ffffff", UI_TEXT, "#e5edf7", "#f8fbff"),
            "primary": (UI_PRIMARY, "#ffffff", UI_PRIMARY, UI_PRIMARY_ACTIVE),
            "soft": ("#f8fbff", UI_PRIMARY, "#bcd2ee", "#eef6ff"),
            "warn": ("#fff7ed", "#ea580c", "#f59e0b", "#fff0ce"),
            "danger": ("#fff5f5", UI_ERROR, "#fecaca", "#fee2e2"),
        }
        bg, fg, border, active_bg = palette.get(variant, palette["default"])
        pixel_w = max(76, int(width * 9.8) + 20)
        pixel_h = 34
        canvas = tk.Canvas(parent, width=pixel_w, height=pixel_h, bg=parent.cget("bg"), bd=0, highlightthickness=0, cursor="hand2")

        def draw(fill):
            canvas.delete("all")
            r = 14
            canvas.create_polygon(
                [
                    1 + r, 1, pixel_w - 1 - r, 1, pixel_w - 1, 1, pixel_w - 1, 1 + r,
                    pixel_w - 1, pixel_h - 1 - r, pixel_w - 1, pixel_h - 1,
                    pixel_w - 1 - r, pixel_h - 1, 1 + r, pixel_h - 1, 1, pixel_h - 1,
                    1, pixel_h - 1 - r, 1, 1 + r, 1, 1,
                ],
                smooth=True,
                splinesteps=24,
                fill=fill,
                outline=border,
            )
            canvas.create_text(pixel_w // 2, pixel_h // 2, text=text, fill=fg, font=ui_font(10, bold=variant in {"primary", "danger"}))

        def click(_event=None):
            try:
                command()
            except TypeError:
                command(_event)
            return "break"

        canvas.bind("<Button-1>", click)
        canvas.bind("<Enter>", lambda _e: draw(active_bg))
        canvas.bind("<Leave>", lambda _e: draw(bg))
        draw(bg)
        return canvas

    form_box = tk.Frame(left_area, bg="#ffffff", highlightthickness=1, highlightbackground="#d7e5e0", highlightcolor="#d7e5e0")
    form_box.pack(fill="x", pady=(0, 10))
    
    form_title = tk.Frame(form_box, bg="#ffffff")
    form_title.pack(fill="x", padx=16, pady=(12, 8))
    mini_icon = tk.Frame(form_title, bg="#eef6ff", width=28, height=28)
    mini_icon.pack(side="left", padx=(0, 10))
    mini_icon.pack_propagate(False)
    tk.Label(mini_icon, text="🔑", bg="#eef6ff", fg=UI_PRIMARY, font=ui_font(12)).pack(expand=True)
    tk.Label(form_title, text="Tạo mã duyệt mới", bg="#ffffff", fg=UI_PRIMARY, font=ui_font(12, bold=True)).pack(side="left")

    form_grid = tk.Frame(form_box, bg="#ffffff")
    form_grid.pack(fill="x", padx=16)

    col1 = tk.Frame(form_grid, bg="#ffffff")
    col1.pack(side="left", fill="x", expand=True, padx=(0, 16))
    tk.Label(col1, text="Mã máy", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
    machine_var = tk.StringVar()
    machine_shell = RoundedMappingEntry(col1, textvariable=machine_var, bg_color="#f8fafc", border_color=UI_BORDER, width=360, height=30, radius=8, font=ui_font(11))
    machine_shell.pack(fill="x", pady=(5, 9))
    machine_entry = machine_shell.entry
    machine_entry.configure(fg=UI_TEXT)
    machine_entry.insert(0, "⊙  Nhập mã máy do thành viên cung cấp")
    machine_entry.bind("<FocusIn>", lambda e: machine_entry.delete(0, 'end') if machine_var.get().startswith("⊙") else None)
    machine_entry.bind("<FocusOut>", lambda e: machine_entry.insert(0, "⊙  Nhập mã máy do thành viên cung cấp") if not machine_var.get() else None)

    col2 = tk.Frame(form_grid, bg="#ffffff")
    col2.pack(side="left", fill="x", expand=True)
    tk.Label(col2, text="Tên người (tuỳ chọn)", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
    name_var = tk.StringVar()
    name_shell = RoundedMappingEntry(col2, textvariable=name_var, bg_color="#f8fafc", border_color=UI_BORDER, width=360, height=30, radius=8, font=ui_font(11))
    name_shell.pack(fill="x", pady=(5, 9))
    name_entry = name_shell.entry
    name_entry.configure(fg=UI_TEXT)
    name_entry.insert(0, "👤  Nhập tên người sử dụng (không bắt buộc)")
    name_entry.bind("<FocusIn>", lambda e: name_entry.delete(0, 'end') if name_var.get().startswith("👤") else None)
    name_entry.bind("<FocusOut>", lambda e: name_entry.insert(0, "👤  Nhập tên người sử dụng (không bắt buộc)") if not name_var.get() else None)

    tk.Label(form_box, text="Mã duyệt", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w", padx=16)
    approval_var = tk.StringVar()
    approval_shell = RoundedMappingEntry(form_box, textvariable=approval_var, bg_color="#f8fafc", border_color="#E5EAF3", width=740, height=30, radius=8, font=ui_font(11))
    approval_shell.pack(fill="x", padx=16, pady=(5, 10))
    approval_entry = approval_shell.entry
    approval_entry.configure(fg=UI_TEXT)
    approval_entry.insert(0, "Mã duyệt sẽ được tạo tự động")
    approval_entry.configure(state="readonly")

    last_generated_code = {"value": ""}

    actions = tk.Frame(form_box, bg="#ffffff")
    actions.pack(fill="x", padx=16, pady=(0, 12))

    list_box = tk.Frame(left_area, bg="#ffffff", highlightthickness=1, highlightbackground="#d7e5e0", highlightcolor="#d7e5e0")
    list_box.pack(fill="both", expand=True)

    list_header = tk.Frame(list_box, bg="#ffffff")
    list_header.pack(fill="x", padx=16, pady=(12, 8))

    list_title_box = tk.Frame(list_header, bg="#ffffff")
    list_title_box.pack(side="left")

    tk.Label(list_title_box, text="▣  Danh sách máy đã duyệt", bg="#ffffff", fg="#0f172a", font=ui_font(12, bold=True)).pack(anchor="w")
    summary_var = tk.StringVar(value="Đang tải...")
    tk.Label(list_title_box, textvariable=summary_var, bg="#ffffff", fg=UI_TEXT, font=ui_font(9)).pack(anchor="w", pady=(2, 0))

    search_bar = tk.Frame(list_box, bg="#ffffff")
    search_bar.pack(fill="x", padx=16, pady=(0, 8))
    search_var = tk.StringVar()
    search_shell = RoundedMappingEntry(search_bar, textvariable=search_var, bg_color="#f8fafc", border_color=UI_BORDER, width=520, height=30, radius=8, font=ui_font(10))
    search_shell.pack(side="left", fill="x", expand=True)
    search_entry = search_shell.entry
    search_entry.configure(fg=UI_TEXT)
    search_entry.insert(0, "🔍  Tìm kiếm theo mã máy, tên người hoặc mã duyệt...")
    search_entry.bind("<FocusIn>", lambda e: search_entry.delete(0, 'end') if search_var.get().startswith("🔍") else None)
    search_entry.bind("<FocusOut>", lambda e: search_entry.insert(0, "🔍  Tìm kiếm theo mã máy, tên người hoặc mã duyệt...") if not search_var.get() else None)

    form_inputs = {machine_entry, name_entry, approval_entry, search_entry}

    list_frame = tk.Frame(list_box, bg="#ffffff")
    list_frame.pack(fill="both", expand=True, padx=16)

    style = ttk.Style()
    style.configure("Custom.Treeview", background="#ffffff", fieldbackground="#ffffff", rowheight=40, borderwidth=0, font=ui_font(10))
    style.configure("Custom.Treeview.Heading", font=ui_font(10, bold=True), background="#eff6ff", foreground="#0f172a", relief="flat", borderwidth=0)
    style.map("Custom.Treeview.Heading", background=[("active", "#e2e8f0")])
    style.layout("Custom.Treeview", [('Custom.Treeview.treearea', {'sticky': 'nswe'})])

    approved_tree = ttk.Treeview(list_frame, style="Custom.Treeview", columns=("select", "machine", "user", "code", "time", "status", "action"), show="headings", height=4)
    approved_tree.heading("select", text="☐")
    approved_tree.heading("machine", text="  Mã máy")
    approved_tree.heading("user", text="Tên người")
    approved_tree.heading("code", text="Mã duyệt")
    approved_tree.heading("time", text="Thời gian duyệt")
    approved_tree.heading("status", text="Trạng thái")
    approved_tree.heading("action", text="Thao tác")

    approved_tree.column("select", width=44, minwidth=44, anchor="center", stretch=False)
    approved_tree.column("machine", width=235, anchor="center")
    approved_tree.column("user", width=140, anchor="center")
    approved_tree.column("code", width=190, anchor="center")
    approved_tree.column("time", width=180, anchor="center")
    approved_tree.column("status", width=150, anchor="center")
    approved_tree.column("action", width=92, anchor="center")

    approved_tree.tag_configure("online", foreground=UI_SUCCESS, background="#ffffff")
    approved_tree.tag_configure("away", foreground="#f59e0b", background="#ffffff")
    approved_tree.tag_configure("old", foreground=UI_ERROR, background="#ffffff")
    approved_tree.tag_configure("stripe", background="#f8fafc")

    approved_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=approved_tree.yview)
    approved_tree.configure(yscrollcommand=approved_scroll.set)
    approved_tree.grid(row=0, column=0, sticky="nsew")
    approved_scroll.grid(row=0, column=1, sticky="ns")
    list_frame.rowconfigure(0, weight=1)
    list_frame.columnconfigure(0, weight=1)

    list_actions = tk.Frame(list_box, bg="#ffffff")
    list_actions.pack(fill="x", padx=16, pady=(8, 10))

    def filter_rows(rows, query):
        query = str(query or "").strip()
        if not query or query.startswith("Tìm kiếm") or query.startswith(" Tìm kiếm") or query.startswith("🔍"):
            return list(rows)
        query_norm = norm(query)
        filtered = []
        for row in rows:
            haystack = " ".join([str(row.get("machine_code", "")), str(row.get("user_name", "")), str(row.get("approval_code", "")), str(row.get("approved_at", "")), str(row.get("last_seen_at", ""))])
            if query_norm in norm(haystack):
                filtered.append(row)
        return filtered

    def _is_editing_approval_form():
        try:
            focused = self.root.focus_get()
        except Exception:
            focused = None
        if focused in form_inputs:
            return True
        try:
            return bool(focused and any(str(focused).startswith(str(widget)) for widget in form_inputs))
        except Exception:
            return False

    def refresh_list(select_machine=None, fill_selection=True, auto=False):
        if auto and _is_editing_approval_form():
            try:
                if getattr(self, "admin_approval_panel", None) is not None and self.admin_approval_panel.winfo_exists():
                    if refresh_job["id"] is not None:
                        try:
                            self.admin_approval_panel.after_cancel(refresh_job["id"])
                        except Exception:
                            pass
                    refresh_job["id"] = self.admin_approval_panel.after(1500, lambda: refresh_list(select_machine=select_machine, fill_selection=False, auto=True))
            except Exception:
                pass
            return

        # Capture search text on main thread (ignore placeholder)
        raw_q = search_var.get()
        query = "" if raw_q.startswith("🔍") or raw_q.startswith("⊙") or raw_q.startswith(" Tìm kiếm") else raw_q

        def _data_worker():
            try:
                import_local_approval_to_admin_list()
                srv = self._get_presence_machine_cache()
                sync_presence_machines_to_admin_list(srv)
                all_r = load_admin_approved_machines()
                filt = filter_rows(all_r, query)
                smap = {}
                for item in srv:
                    if not isinstance(item, dict):
                        continue
                    mk = str(item.get("machine_code") or "").strip().upper()
                    if mk:
                        smap[mk] = item
                # Schedule UI update on main thread
                try:
                    if getattr(self, "admin_approval_panel", None) is not None and self.admin_approval_panel.winfo_exists():
                        self.admin_approval_panel.after(0, lambda: _apply_ui(all_r, filt, smap, select_machine, fill_selection))
                except Exception:
                    pass
            except Exception:
                pass

        def _apply_ui(all_rows, rows, server_map, sel_machine, do_fill):
            try:
                existing_iids = set(approved_tree.get_children())
                new_iids = set()

                for i, row in enumerate(rows):
                    machine = row.get("machine_code", "")
                    if not machine:
                        continue
                    new_iids.add(machine)
                    live = server_map.get(str(machine).strip().upper(), {})
                    last_seen = live.get("last_seen_at") or row.get("last_seen_at") or row.get("approved_at") or ""
                    status_text = format_machine_last_seen(last_seen)
                    if is_machine_active_recently(last_seen):
                        status_tag = "online"
                        status_text = "● Đang hoạt động"
                    else:
                        dt = parse_machine_datetime(last_seen)
                        if dt is None:
                            status_tag = "old"
                            status_text = "● Không rõ"
                        else:
                            age_min = max(0, int((datetime.now() - dt).total_seconds() // 60))
                            if age_min < 180:
                                status_tag = "away"
                                status_text = "● Vừa rời"
                            else:
                                status_tag = "old"
                                status_text = "● Không hoạt động"
                    tags = (status_tag, "stripe" if i % 2 == 1 else "")
                    values = ("☐", "▣  " + machine, str(row.get("user_name", "") or "").strip(), row.get("approval_code", ""), row.get("approved_at", ""), status_text, "Thao tác")
                    if machine in existing_iids:
                        approved_tree.item(machine, values=values, tags=tags)
                    else:
                        approved_tree.insert("", "end", iid=machine, values=values, tags=tags)

                for old_iid in existing_iids - new_iids:
                    approved_tree.delete(old_iid)

                total_count = len(all_rows)
                online_count = sum(1 for r in all_rows if is_machine_active_recently(
                    (server_map.get(str(r.get("machine_code", "")).strip().upper(), {}) or {}).get("last_seen_at")
                    or r.get("last_seen_at") or r.get("approved_at") or ""
                ))
                summary_var.set(f"Đã duyệt: {total_count} máy | Đang hoạt động: {online_count} máy")
                stat_total_var.set(f"✓ Đã tải {total_count} dòng.")
                stat_total_num_var.set(str(total_count))
                stat_active_var.set(str(online_count))
                stat_pending_var.set("0")
                stat_blocked_var.set("0")
                stat_time_var.set(f"Cập nhật: {datetime.now().strftime('%H:%M:%S %d/%m/%Y')}")

                target = str(sel_machine or "").strip().upper()
                children = approved_tree.get_children()
                selected_id = target if target and approved_tree.exists(target) else (children[0] if children else "")
                if selected_id:
                    approved_tree.focus(selected_id)
                    approved_tree.see(selected_id)
                    if do_fill:
                        approved_tree.selection_set(selected_id)
                        fill_from_selected()
                    else:
                        approved_tree.selection_remove(approved_tree.selection())
            except Exception:
                pass

            # Schedule next auto-refresh
            try:
                if refresh_job["id"] is not None:
                    self.admin_approval_panel.after_cancel(refresh_job["id"])
            except Exception:
                pass
            try:
                if getattr(self, "admin_approval_panel", None) is not None and self.admin_approval_panel.winfo_exists():
                    refresh_job["id"] = self.admin_approval_panel.after(5000, lambda: refresh_list(select_machine=sel_machine, fill_selection=False, auto=True))
            except Exception:
                refresh_job["id"] = None

        threading.Thread(target=_data_worker, daemon=True).start()

    def search_rows():
        refresh_list()
        search_entry.focus_set()

    def clear_search():
        search_var.set("")
        search_entry.delete(0, 'end')
        search_entry.insert(0, "🔍  Tìm kiếm theo mã máy, tên người hoặc mã duyệt...")
        search_entry.configure(fg=UI_TEXT)
        refresh_list()

    def fill_from_selected(_event=None):
        selected = approved_tree.selection()
        if not selected:
            return
        values = approved_tree.item(selected[0], "values")
        if values:
            machine_text = re.sub(r"^[^\w]*", "", str(values[1] if len(values) > 1 else "")).replace("🖥", "").strip()
            machine_var.set(machine_text)
            machine_entry.configure(fg=UI_TEXT)
            name_var.set(values[2] if len(values) > 2 else "")
            name_entry.configure(fg=UI_TEXT)
            approval_entry.configure(state="normal")
            approval_var.set(values[3] if len(values) > 3 else "")
            approval_entry.configure(fg=UI_TEXT)
            approval_entry.configure(state="readonly")

    def clear_approval_form():
        approved_tree.selection_remove(approved_tree.selection())
        machine_var.set("")
        name_var.set("")
        approval_entry.configure(state="normal")
        approval_var.set("")
        approval_entry.configure(state="readonly")
        machine_entry.focus_set()

    approved_tree.bind("<<TreeviewSelect>>", fill_from_selected)
    search_entry.bind("<Return>", lambda _e: search_rows())

    def generate():
        machine = str(machine_var.get() or "").strip().upper()
        user_name = str(name_var.get() or "").strip()
        if machine.startswith("⊙") or "NHẬP MÃ MÁY" in machine:
            machine = ""
        if user_name.startswith("👤") or "Nhập tên người" in user_name:
            user_name = ""
        code = remember_admin_approved_machine(machine, user_name)
        if not code:
            messagebox.showwarning("Thiếu mã máy", "Bạn chưa nhập mã máy cần duyệt.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(code)
        last_generated_code["value"] = code
        approval_entry.configure(state="normal")
        approval_var.set(code)
        approval_entry.configure(state="readonly")
        refresh_list(select_machine=machine, fill_selection=False)
        machine_var.set("")
        name_var.set("")
        machine_entry.focus_set()

    def copy_code():
        code_to_copy = approval_var.get() or last_generated_code["value"]
        if not code_to_copy:
            generate()
            code_to_copy = approval_var.get() or last_generated_code["value"]
        if not code_to_copy:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(code_to_copy)

    def copy_selected_machine():
        selected = approved_tree.selection()
        if not selected:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(str(selected[0]))

    def copy_selected_code():
        selected = approved_tree.selection()
        if not selected:
            return
        values = approved_tree.item(selected[0], "values")
        code = str(values[3] if len(values) > 3 else "").strip()
        if not code:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(code)

    def delete_selected():
        selected = approved_tree.selection()
        if not selected:
            messagebox.showinfo("Chưa chọn máy", "Chọn một máy trong danh sách trước khi xóa.")
            return
        machine = str(selected[0])
        if not messagebox.askyesno("Xóa máy đã duyệt", f"Xóa máy này khỏi danh sách đã duyệt?\nMã duyệt cũ trên máy đó sẽ không dùng lại được.\n\n{machine}"):
            return
        delete_admin_approved_machine(machine)
        refresh_list(fill_selection=False)
        clear_approval_form()

    def open_list_menu(event):
        row_id = approved_tree.identify_row(event.y)
        if row_id:
            approved_tree.selection_set(row_id)
            approved_tree.focus(row_id)
            fill_from_selected()
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Copy mã máy", command=copy_selected_machine)
        menu.add_command(label="Copy mã duyệt", command=copy_selected_code)
        menu.add_separator()
        menu.add_command(label="Xóa máy này", command=delete_selected)
        menu.add_separator()
        menu.add_command(label="Tải lại danh sách", command=clear_search)
        menu.tk_popup(event.x_root, event.y_root)

    def handle_tree_click(event):
        region = approved_tree.identify("region", event.x, event.y)
        row_id = approved_tree.identify_row(event.y)
        col_id = approved_tree.identify_column(event.x)
        if row_id:
            approved_tree.selection_set(row_id)
            approved_tree.focus(row_id)
            fill_from_selected()
        if region == "cell" and col_id == "#7":
            open_list_menu(event)
            return "break"
        return None

    approved_tree.bind("<Delete>", lambda _e: delete_selected())
    approved_tree.bind("<Button-3>", open_list_menu)
    approved_tree.bind("<Button-1>", handle_tree_click)

    admin_button(actions, "+  Tạo mã duyệt", generate, width=14, variant="primary").pack(side="left", padx=(0, 10))
    admin_button(actions, "▣  Copy mã", copy_code, width=11, variant="soft").pack(side="left")
    admin_button(actions, "×  Đóng", close_panel, width=9).pack(side="right")

    admin_button(search_bar, "⌕  Tìm", search_rows, width=8, variant="soft").pack(side="left", padx=(10, 0))
    admin_button(search_bar, "▽  Bộ lọc", search_rows, width=9, variant="default").pack(side="left", padx=(8, 0))
    admin_button(search_bar, "🗑  Xóa máy", delete_selected, width=10, variant="danger").pack(side="left", padx=(8, 0))

    admin_button(list_actions, "🗑 Xóa máy đã chọn", delete_selected, width=15, variant="danger").pack(side="left")
    admin_button(list_actions, "↻ Tải lại danh sách", clear_search, width=15, variant="soft").pack(side="left", padx=(8, 0))

    log_box = tk.Frame(list_box, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3")
    # Log notifications live in the Admin sidebar now, not inside machine approval.

    log_header = tk.Frame(log_box, bg="#ffffff")
    log_header.pack(fill="x", padx=12, pady=(10, 6))
    tk.Label(log_header, text="⚠  Log lỗi User gửi", bg="#ffffff", fg=UI_ERROR, font=ui_font(11, bold=True)).pack(side="left")
    log_count_badge = tk.Label(
        log_header,
        text="0",
        bg=UI_ERROR,
        fg="#ffffff",
        font=ui_font(8, bold=True),
        padx=7,
        pady=2,
    )
    log_count_badge.pack(side="left", padx=(8, 0))
    log_summary_var = tk.StringVar(value="Chưa tải log.")
    tk.Label(log_header, textvariable=log_summary_var, bg="#ffffff", fg=UI_TEXT, font=ui_font(9)).pack(side="left", padx=(12, 0))

    log_frame = tk.Frame(log_box, bg="#ffffff")
    log_frame.pack(fill="x", padx=12)
    log_tree = ttk.Treeview(
        log_frame,
        style="Custom.Treeview",
        columns=("time", "user", "windows", "computer", "message"),
        show="headings",
        height=3,
    )
    log_tree.heading("time", text="Thời gian")
    log_tree.heading("user", text="Tên người")
    log_tree.heading("windows", text="User Windows")
    log_tree.heading("computer", text="Tên máy")
    log_tree.heading("message", text="Thông báo")
    log_tree.column("time", width=150, anchor="center", stretch=False)
    log_tree.column("user", width=130, anchor="center", stretch=False)
    log_tree.column("windows", width=130, anchor="center", stretch=False)
    log_tree.column("computer", width=130, anchor="center", stretch=False)
    log_tree.column("message", width=360, anchor="w")
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_tree.yview)
    log_tree.configure(yscrollcommand=log_scroll.set)
    log_tree.grid(row=0, column=0, sticky="ew")
    log_scroll.grid(row=0, column=1, sticky="ns")
    log_frame.columnconfigure(0, weight=1)
    log_cache = {"rows": []}

    def selected_log_row():
        selected = log_tree.selection()
        if not selected:
            return None
        try:
            idx = int(str(selected[0]))
            for row in log_cache["rows"]:
                if int(row.get("id") or -1) == idx:
                    return row
        except Exception:
            pass
        return None

    def show_selected_log():
        row = selected_log_row()
        if not row:
            messagebox.showinfo("Log lỗi", "Chọn một dòng log trước.")
            return
        detail = (
            f"Thời gian: {row.get('created_at', '')}\n"
            f"Mã máy: {row.get('machine_code', '')}\n"
            f"Tên người: {row.get('user_name', '')}\n"
            f"User Windows: {row.get('windows_user', '')}\n"
            f"Tên máy Windows: {row.get('computer_name', '')}\n"
            f"Vai trò: {row.get('role', '')}\n\n"
            f"Kiểm tra log: {'Hợp lệ' if row.get('hash_ok') else 'Có dấu hiệu bị sửa hoặc log cũ chưa có hash'}\n"
            f"Mã hash: {row.get('log_hash', '')}\n\n"
            f"Nội dung log:\n{row.get('log_text', '')}"
        )
        top = tk.Toplevel(self.root)
        top.title("Chi tiết log lỗi")
        top.configure(bg="#f3f6fb")
        self._fit_dialog_to_screen(top, 860, 600, min_w=780, min_h=500, max_ratio=0.82, lock_size=False)
        body = tk.Frame(top, bg="#f3f6fb", padx=18, pady=18)
        body.pack(fill="both", expand=True)

        header = tk.Frame(body, bg="#f3f6fb")
        header.pack(fill="x", pady=(0, 12))
        title_box = tk.Frame(header, bg="#f3f6fb")
        title_box.pack(side="left", fill="x", expand=True)
        tk.Label(title_box, text="Chi tiết log lỗi", bg="#f3f6fb", fg=UI_TEXT, font=ui_font(16, bold=True)).pack(anchor="w")
        tk.Label(
            title_box,
            text="Thông tin user gửi về server Admin",
            bg="#f3f6fb",
            fg=UI_MUTED,
            font=ui_font(10),
        ).pack(anchor="w", pady=(2, 0))

        meta_card = tk.Frame(body, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3", padx=14, pady=12)
        meta_card.pack(fill="x", pady=(0, 12))
        meta_card.grid_columnconfigure(0, weight=1)
        meta_card.grid_columnconfigure(1, weight=1)
        meta_card.grid_columnconfigure(2, weight=1)

        def meta_item(parent, row_idx, col_idx, label, value):
            cell = tk.Frame(parent, bg="#ffffff")
            cell.grid(row=row_idx, column=col_idx, sticky="ew", padx=(0 if col_idx == 0 else 14, 0), pady=(0 if row_idx == 0 else 10, 0))
            tk.Label(cell, text=label, bg="#ffffff", fg=UI_MUTED, font=ui_font(9)).pack(anchor="w")
            tk.Label(
                cell,
                text=str(value or "-"),
                bg="#ffffff",
                fg=UI_TEXT,
                font=ui_font(10, bold=True),
                anchor="w",
                justify="left",
                wraplength=230,
            ).pack(anchor="w", pady=(2, 0))

        meta_item(meta_card, 0, 0, "Thời gian", row.get("created_at", ""))
        meta_item(meta_card, 0, 1, "Tên người", row.get("user_name", ""))
        meta_item(meta_card, 0, 2, "Vai trò", row.get("role", ""))
        meta_item(meta_card, 1, 0, "Mã máy", row.get("machine_code", ""))
        meta_item(meta_card, 1, 1, "User Windows", row.get("windows_user", ""))
        meta_item(meta_card, 1, 2, "Tên máy Windows", row.get("computer_name", ""))

        log_card = tk.Frame(body, bg="#ffffff", highlightthickness=1, highlightbackground="#dbe6f3")
        log_card.pack(fill="both", expand=True)
        log_head = tk.Frame(log_card, bg="#ffffff", padx=14, pady=10)
        log_head.pack(fill="x")
        tk.Label(log_head, text="Nội dung log", bg="#ffffff", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(side="left")
        tk.Label(log_head, text=str(row.get("message") or ""), bg="#ffffff", fg=UI_MUTED, font=ui_font(9)).pack(side="right")
        text_wrap = tk.Frame(log_card, bg="#eef2f7", padx=1, pady=1)
        text_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        text = tk.Text(text_wrap, wrap="word", bg="#0f172a", fg="#e5edf7", insertbackground="#e5edf7", relief="flat", bd=0, font=("Consolas", 10), padx=12, pady=10)
        text.pack(fill="both", expand=True)
        text.insert("1.0", str(row.get("log_text", "") or "Không có nội dung log."))
        text.configure(state="disabled")
        actions_row = tk.Frame(body, bg="#f3f6fb")
        actions_row.pack(fill="x", pady=(12, 0))
        def copy_detail():
            self.root.clipboard_clear()
            self.root.clipboard_append(detail)
        admin_button(actions_row, "Copy log", copy_detail, width=10, variant="soft").pack(side="left")
        admin_button(actions_row, "Đóng", top.destroy, width=8).pack(side="right")
        self._center_dialog_on_screen(top)

    def copy_selected_log():
        row = selected_log_row()
        if not row:
            messagebox.showinfo("Log lỗi", "Chọn một dòng log trước.")
            return
        text = (
            f"{row.get('created_at', '')} | {row.get('machine_code', '')} | "
            f"{row.get('user_name', '')} | {row.get('windows_user', '')} | "
            f"{row.get('computer_name', '')}\n{row.get('log_text', '')}"
        )
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def refresh_error_logs():
        server_url = resolve_presence_server_url(getattr(self, "presence_server_var", tk.StringVar(value="")).get())
        log_summary_var.set("Đang tải log...")

        def _worker():
            rows = fetch_presence_error_logs(server_url, limit=80, timeout=3)

            def _apply():
                try:
                    log_cache["rows"] = rows
                    count_text = str(len(rows))
                    log_count_badge.configure(text=count_text, bg=UI_ERROR if rows else "#94a3b8")
                    log_tree.delete(*log_tree.get_children())
                    for row in rows[:80]:
                        iid = str(row.get("id") or "")
                        if not iid:
                            continue
                        log_tree.insert(
                            "",
                            "end",
                            iid=iid,
                            values=(
                                row.get("created_at", ""),
                                row.get("user_name", ""),
                                row.get("windows_user", ""),
                                row.get("computer_name", ""),
                                str(row.get("message") or row.get("log_text") or "")[:120],
                            ),
                        )
                    log_summary_var.set(f"{len(rows)} lần lỗi gần nhất." if rows else "Chưa có log user gửi.")
                except Exception:
                    log_summary_var.set("Không hiển thị được log.")

            try:
                self.root.after(0, _apply)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    log_tree.bind("<Double-1>", lambda _e: show_selected_log())
    log_actions = tk.Frame(log_box, bg="#ffffff")
    log_actions.pack(fill="x", padx=12, pady=(8, 10))
    admin_button(log_actions, "↻ Tải log", refresh_error_logs, width=10, variant="soft").pack(side="left")
    admin_button(log_actions, "Xem chi tiết", show_selected_log, width=12, variant="primary").pack(side="left", padx=(8, 0))
    admin_button(log_actions, "Copy log", copy_selected_log, width=10, variant="default").pack(side="left", padx=(8, 0))

    refresh_list(fill_selection=False)
    clear_approval_form()


def install_admin_ui(app_cls):
    app_cls._admin_log_badge_loop = _admin_log_badge_loop
    app_cls.open_admin_log_panel = open_admin_log_panel
    app_cls.open_admin_backup_panel = open_admin_backup_panel
    app_cls.open_admin_approval_panel = open_admin_approval_panel
