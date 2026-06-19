"""Settings, help, and about dialogs for the main application window."""

import re
import tkinter as tk
from tkinter import filedialog, messagebox

from gk_pilepro.gk_core import (
    DEFAULT_MODEL,
    DEFAULT_PRESENCE_SERVER_URL,
    current_app_user_name,
    current_user_role_labels,
    is_admin_build,
    load_env_values,
    resolve_presence_server_url,
    save_env,
)
from gk_pilepro.ui.gk_icons import get_windows_dpi, get_windows_work_area
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
    RoundedMappingEntry,
    scale_px,
    ui_button,
    ui_font,
)


def _rebuild_ui_after_settings_change(self, return_page="home"):
    try:
        for child in list(self.root.winfo_children()):
            child.destroy()
    except Exception:
        pass

    self.nav_widgets = {}
    self.home_page = None
    self.excel_page = None
    self.history_page = None
    self.mapping_page = None
    self.mapping_templates_inner = None
    self.admin_approval_panel = None
    self.settings_panel = None

    self._setup_responsive_metrics()
    self.build_ui()

    if return_page not in {"home", "excel", "history", "mapping"}:
        return_page = "home"
    try:
        self.show_page(return_page)
    except Exception:
        self.show_home_page()
    try:
        self._set_status("Đã áp dụng cấu hình hiển thị.", "success")
    except Exception:
        pass


def show_settings_dialog(self, event=None):
    return_page = getattr(self, "_dialog_return_page", None)
    if return_page not in {"home", "excel", "history", "mapping"}:
        return_page = getattr(self, "current_page", "home")
    settings_is_admin = is_admin_build()

    if getattr(self, "settings_panel", None) is not None:
        try:
            self.settings_panel.lift()
            self.settings_panel.focus_set()
        except Exception:
            pass
        return "break"

    win = tk.Toplevel(self.root)
    win.configure(bg="#1f2128")
    win.title("Cài đặt")
    win.resizable(False, False)
    try:
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass
    self.settings_panel = win

    win.bind("<Destroy>", lambda e: setattr(self, "settings_panel", None) if e.widget is win else None, add="+")

    screen_w = int(getattr(self, "screen_w", 0) or win.winfo_screenwidth() or 1366)
    screen_h = int(getattr(self, "screen_h", 0) or win.winfo_screenheight() or 768)
    if screen_w <= 1366 or screen_h <= 768:
        dialog_w, dialog_h = 760, 540
    elif screen_w <= 1600 or screen_h <= 900:
        dialog_w, dialog_h = 840, 580
    elif screen_w <= 1920 or screen_h <= 1080:
        dialog_w, dialog_h = 900, 620
    else:
        dialog_w, dialog_h = 960, 650
    dialog_w = min(dialog_w, int(screen_w * 0.90))
    dialog_h = min(dialog_h, int(screen_h * 0.88))
    self._fit_dialog_to_screen(win, dialog_w, dialog_h, min_w=dialog_w, min_h=dialog_h, max_ratio=0.90, lock_size=True)

    dark_bg = "#1f2128"
    panel_bg = "#272a33"
    card_bg = "#2f323d"
    card_border = "#444857"
    card_soft = "#383c49"
    accent = "#6f6bff"
    accent_soft = "#8f8aff"
    text_main = "#f4f6fb"
    text_sub = "#b4bccf"
    text_dim = "#8790a5"

    dialog_shell = tk.Frame(win, bg=dark_bg, highlightthickness=1, highlightbackground=card_border)
    dialog_shell.pack(fill="both", expand=True)

    root = tk.Frame(dialog_shell, bg=dark_bg, padx=14, pady=14)
    root.grid(row=0, column=0, sticky="nsew")
    dialog_shell.grid_rowconfigure(0, weight=1)
    dialog_shell.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(2, weight=1)

    title_row = tk.Frame(root, bg=dark_bg)
    title_row.grid(row=0, column=0, sticky="ew")
    tk.Label(title_row, text="Hiển thị", bg=dark_bg, fg=text_main, font=ui_font(11, bold=True)).pack(anchor="w")
    tk.Label(
        title_row,
        text="Chọn độ phân giải và cấu hình khởi động của ứng dụng.",
        bg=dark_bg,
        fg=text_sub,
        font=ui_font(10),
    ).pack(anchor="w", pady=(6, 0))

    nav_row = tk.Frame(root, bg=dark_bg)
    nav_row.grid(row=1, column=0, sticky="ew", pady=(10, 6))

    page_host = tk.Frame(root, bg=dark_bg)
    page_host.grid(row=2, column=0, sticky="nsew")
    root.grid_rowconfigure(2, weight=1)
    root.grid_columnconfigure(0, weight=1)

    footer = tk.Frame(root, bg=dark_bg)
    footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))

    def dark_button(parent, text, command, width=12, active=False, accent_button=False):
        bg = accent if accent_button else (card_soft if active else panel_bg)
        fg = "#ffffff"
        border = accent if accent_button or active else card_border
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=accent_soft if accent_button else "#3f4351",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=border,
            font=ui_font(10, bold=(active or accent_button)),
            padx=12,
            pady=7,
            width=width,
        )

    def entry_shell(parent, width=24):
        shell = tk.Frame(parent, bg=panel_bg, highlightthickness=1, highlightbackground=card_border)
        shell.pack_propagate(False)
        shell.config(width=width * 12, height=34)
        return shell

    top_tabs = {}
    display_page = tk.Frame(page_host, bg=dark_bg)
    api_page = tk.Frame(page_host, bg=dark_bg)

    def show_page(name):
        for child in page_host.winfo_children():
            child.pack_forget()
        target_page = display_page if name == "display" or not settings_is_admin else api_page
        target_page.pack(fill="both", expand=True)
        for key, btn in top_tabs.items():
            active = key == name
            btn.config(
                bg=accent if active else panel_bg,
                fg="#ffffff" if active else text_main,
                highlightbackground=accent if active else card_border,
                highlightcolor=accent if active else card_border,
                font=ui_font(11, bold=active),
            )

    tab_defs = [("display", "Hiển thị")]
    if settings_is_admin:
        tab_defs.append(("api", "API"))

    for key, label in tab_defs:
        btn = dark_button(nav_row, label, lambda k=key: show_page(k), width=11, active=(key == "display"))
        btn.pack(side="left", padx=(0, 8))
        top_tabs[key] = btn

    display_card = tk.Frame(display_page, bg=panel_bg, highlightthickness=1, highlightbackground=card_border)
    display_card.pack(fill="both", expand=True)

    left_panel = tk.Frame(display_card, bg=panel_bg, width=140, padx=12, pady=12)
    left_panel.pack(side="left", fill="y")
    left_panel.pack_propagate(False)
    tk.Label(left_panel, text="Độ phân giải", bg=panel_bg, fg=text_main, font=ui_font(11, bold=True)).pack(anchor="w")
    tk.Label(
        left_panel,
        text="Chọn preset phù hợp với màn hình để app hiển thị đầy đủ.",
        bg=panel_bg,
        fg=text_sub,
        font=ui_font(10),
        justify="left",
        wraplength=104,
    ).pack(anchor="w", pady=(8, 0))

    right_panel = tk.Frame(display_card, bg=panel_bg, padx=14, pady=12)
    right_panel.pack(side="left", fill="both", expand=True)

    mode_bar = tk.Frame(right_panel, bg=panel_bg)
    mode_bar.pack(fill="x")
    selected_summary_var = tk.StringVar(value="")
    selected_summary = tk.Label(
        right_panel,
        textvariable=selected_summary_var,
        bg=panel_bg,
        fg=text_sub,
        font=ui_font(10),
        anchor="w",
        justify="left",
        wraplength=460,
    )
    selected_summary.pack(fill="x", pady=(6, 0))
    mode_host = tk.Frame(right_panel, bg=panel_bg)
    mode_host.pack(fill="both", expand=True, pady=(10, 0))

    profile_catalog = {
        "Tự động": [],
        "Ngang": [
            ("1920 x 1080 (DPI 280)", "display_1920x1080"),
            ("1600 x 900 (DPI 240)", "display_1600x900"),
            ("1366 x 768 (DPI 240)", "display_1366x768"),
            ("1280 x 720 (DPI 240)", "display_1280x720"),
            ("960 x 540 (DPI 160)", "display_960x540"),
        ],
        "Dọc": [
            ("1080 x 1920 (DPI 280)", "display_1080x1920"),
            ("900 x 1600 (DPI 240)", "display_900x1600"),
            ("768 x 1366 (DPI 240)", "display_768x1366"),
            ("720 x 1280 (DPI 240)", "display_720x1280"),
            ("540 x 960 (DPI 160)", "display_540x960"),
        ],
        "Siêu rộng": [
            ("2560 x 1080 (DPI 240)", "display_2560x1080"),
            ("3440 x 1440 (DPI 280)", "display_3440x1440"),
            ("1920 x 800 (DPI 220)", "display_1920x800"),
        ],
    }

    profile_lookup = {}
    for mode_name, items in profile_catalog.items():
        for label, key in items:
            profile_lookup[key] = (mode_name, label)

    current_profile = str(getattr(self, "screen_profile_var", tk.StringVar(value="auto")).get() or "auto").strip().lower()
    custom_match = re.match(r"^custom:(\d+)x(\d+)(?:@(\d+))?$", current_profile)

    initial_mode = "Tự động"
    initial_choice = "auto"
    custom_width = str(self.screen_w)
    custom_height = str(self.screen_h)
    custom_dpi = "240"

    if current_profile == "auto":
        initial_mode = "Tự động"
    elif custom_match:
        initial_mode = "Tùy chỉnh"
        custom_width = custom_match.group(1)
        custom_height = custom_match.group(2)
        custom_dpi = custom_match.group(3) or "240"
    elif current_profile in profile_lookup:
        initial_mode = profile_lookup[current_profile][0]
        initial_choice = current_profile
    elif current_profile == "laptop_156":
        initial_mode = "Ngang"
        initial_choice = "display_1366x768"
    elif current_profile == "laptop_16":
        initial_mode = "Ngang"
        initial_choice = "display_1600x900"

    mode_var = tk.StringVar(value=initial_mode)
    choice_var = tk.StringVar(value=initial_choice)
    custom_width_var = tk.StringVar(value=custom_width)
    custom_height_var = tk.StringVar(value=custom_height)
    custom_dpi_var = tk.StringVar(value=custom_dpi)

    mode_buttons = {}
    mode_frames = {}
    option_refreshers = []

    def describe_selected():
        mode_name = mode_var.get()
        if mode_name == "Tự động":
            width = int(getattr(self, "screen_w", 0) or 0)
            height = int(getattr(self, "screen_h", 0) or 0)
            dpi = int(getattr(self, "screen_dpi", 96) or 96)
            return f"Đang chọn: Tự động theo màn hình thật {width} x {height} (DPI {dpi})"
        if mode_name == "Tùy chỉnh":
            width = str(custom_width_var.get()).strip() or "?"
            height = str(custom_height_var.get()).strip() or "?"
            dpi = str(custom_dpi_var.get()).strip() or "240"
            return f"Đang chọn: Tùy chỉnh {width} x {height} (DPI {dpi})"
        label = None
        for _mode_name, items in profile_catalog.items():
            for item_label, item_key in items:
                if item_key == choice_var.get():
                    label = item_label
                    break
            if label:
                break
        return f"Đang chọn: {label or 'Chưa chọn'}"

    def refresh_selected_summary(*_args):
        selected_summary_var.set(describe_selected())

    def set_mode(mode_name):
        mode_var.set(mode_name)
        for child in mode_host.winfo_children():
            child.pack_forget()
        for key, btn in mode_buttons.items():
            active = key == mode_name
            btn.config(
                bg=accent if active else card_soft,
                fg="#ffffff" if active else text_main,
                highlightbackground=accent if active else card_border,
                highlightcolor=accent if active else card_border,
                font=ui_font(11, bold=active),
            )
        mode_frames[mode_name].pack(fill="both", expand=True)
        refresh_selected_summary()

    for mode_name in ("Tự động", "Ngang", "Dọc", "Siêu rộng", "Tùy chỉnh"):
        btn = dark_button(mode_bar, mode_name, lambda m=mode_name: set_mode(m), width=10, active=(mode_name == initial_mode))
        btn.pack(side="left", padx=(0, 6))
        mode_buttons[mode_name] = btn

    auto_panel = tk.Frame(mode_host, bg=panel_bg)
    mode_frames["Tự động"] = auto_panel
    auto_card = tk.Frame(auto_panel, bg=card_soft, highlightthickness=1, highlightbackground=card_border, padx=14, pady=12)
    auto_card.pack(fill="x")
    tk.Label(auto_card, text="Tự động theo màn hình thật", bg=card_soft, fg=text_main, font=ui_font(11, bold=True)).pack(anchor="w")
    tk.Label(auto_card, text="Đây là chế độ mặc định khi chạy thật. App tự lấy vùng làm việc của màn hình hiện tại và tự scale giao diện.", bg=card_soft, fg=text_sub, font=ui_font(10), wraplength=460, justify="left").pack(anchor="w", pady=(4, 8))
    tk.Label(auto_card, text=f"Hiện tại: {self.screen_w} x {self.screen_h} (DPI {getattr(self, 'screen_dpi', 96)})", bg=card_soft, fg=text_main, font=ui_font(10, bold=True)).pack(anchor="w")

    def make_option_frame(parent, label, value, extra_text=None):
        row = tk.Frame(parent, bg=card_soft, highlightthickness=1, highlightbackground=card_border, cursor="hand2")
        row.pack(fill="x", pady=(0, 8))
        radio = tk.Radiobutton(
            row,
            text=label,
            variable=choice_var,
            value=value,
            bg=card_soft,
            fg=text_main,
            activebackground=card_soft,
            activeforeground=text_main,
            selectcolor=accent,
            indicatoron=True,
            bd=0,
            highlightthickness=0,
            font=ui_font(11),
            padx=12,
            pady=5,
            anchor="w",
        )
        radio.pack(fill="x")
        if extra_text:
            tk.Label(row, text=extra_text, bg=card_soft, fg=text_sub, font=ui_font(10), anchor="w").pack(fill="x", padx=38, pady=(0, 5))

        def select_row(_event=None, selected_value=value):
            choice_var.set(selected_value)
            refresh_selected_summary()
            return "break"

        for widget in (row, radio):
            widget.bind("<Button-1>", select_row)
            widget.bind("<Enter>", lambda _e, w=row: w.config(highlightbackground=accent))
            widget.bind("<Leave>", lambda _e, w=row: w.config(highlightbackground=accent if choice_var.get() == value else card_border))

        def update_row_state(*_args):
            row.config(highlightbackground=accent if choice_var.get() == value else card_border)
            radio.config(fg=text_main)

        choice_var.trace_add("write", update_row_state)
        option_refreshers.append(update_row_state)

    for mode_name, items in profile_catalog.items():
        panel = tk.Frame(mode_host, bg=panel_bg)
        mode_frames[mode_name] = panel
        option_box = tk.Frame(panel, bg=panel_bg)
        option_box.pack(fill="both", expand=True)
        for label, key in items:
            make_option_frame(option_box, label, key)

    custom_panel = tk.Frame(mode_host, bg=panel_bg)
    mode_frames["Tùy chỉnh"] = custom_panel
    custom_card = tk.Frame(custom_panel, bg=card_soft, highlightthickness=1, highlightbackground=card_border, padx=14, pady=14)
    custom_card.pack(fill="x")
    tk.Label(custom_card, text="Thiết lập thủ công", bg=card_soft, fg=text_main, font=ui_font(11, bold=True)).pack(anchor="w")
    tk.Label(custom_card, text="Nhập đúng kích thước màn hình để app ép bố cục theo giá trị này.", bg=card_soft, fg=text_sub, font=ui_font(10)).pack(anchor="w", pady=(3, 10))

    custom_grid = tk.Frame(custom_card, bg=card_soft)
    custom_grid.pack(fill="x")

    def make_field(parent, title, var, width=10):
        box = tk.Frame(parent, bg=card_soft)
        tk.Label(box, text=title, bg=card_soft, fg=text_sub, font=ui_font(10)).pack(anchor="w")
        shell = tk.Frame(box, bg=panel_bg, highlightthickness=1, highlightbackground=card_border)
        shell.pack(anchor="w", pady=(4, 0))
        entry = tk.Entry(
            shell,
            textvariable=var,
            bg=panel_bg,
            fg=text_main,
            insertbackground=text_main,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=ui_font(11),
            width=width,
        )
        entry.pack(padx=10, pady=7)
        return box

    make_field(custom_grid, "Rộng", custom_width_var, width=10).grid(row=0, column=0, sticky="w", padx=(0, 14))
    make_field(custom_grid, "Cao", custom_height_var, width=10).grid(row=0, column=1, sticky="w", padx=(0, 14))
    make_field(custom_grid, "DPI", custom_dpi_var, width=10).grid(row=0, column=2, sticky="w")
    tk.Label(custom_card, text="Ví dụ: 1600 x 900 để test màn 16 inch, 1366 x 768 để test màn 15.6 inch.", bg=card_soft, fg=text_dim, font=ui_font(10), wraplength=420, justify="left").pack(anchor="w", pady=(10, 0))

    custom_width_var.trace_add("write", refresh_selected_summary)
    custom_height_var.trace_add("write", refresh_selected_summary)
    custom_dpi_var.trace_add("write", refresh_selected_summary)
    choice_var.trace_add("write", refresh_selected_summary)
    for refresh_row in option_refreshers:
        refresh_row()
    refresh_selected_summary()

    set_mode(initial_mode)

    if settings_is_admin:
        api_card = tk.Frame(api_page, bg=panel_bg, highlightthickness=1, highlightbackground=card_border, padx=16, pady=16)
        api_card.pack(fill="both", expand=True)
        tk.Label(api_card, text="Cấu hình API", bg=panel_bg, fg=text_main, font=ui_font(12, bold=True)).pack(anchor="w")
        tk.Label(api_card, text="Thiết lập khóa Gemini và model dùng cho OCR / đọc bảng.", bg=panel_bg, fg=text_sub, font=ui_font(11)).pack(anchor="w", pady=(6, 14))

        key_box = tk.Frame(api_card, bg=panel_bg)
        key_box.pack(fill="x", pady=(0, 16))
        tk.Label(key_box, text="Khóa API", bg=panel_bg, fg=text_sub, font=ui_font(11, bold=True)).pack(anchor="w", pady=(0, 6))
        key_shell = tk.Frame(key_box, bg=card_soft, highlightthickness=1, highlightbackground=card_border)
        key_shell.pack(fill="x")
        api_entry = tk.Entry(
            key_shell,
            textvariable=self.api_key_var,
            bg=card_soft,
            fg=text_main,
            insertbackground=text_main,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=ui_font(11),
        )
        api_entry.pack(fill="x", padx=12, pady=9)

        model_box = tk.Frame(api_card, bg=panel_bg)
        model_box.pack(fill="x", pady=(0, 16))
        tk.Label(model_box, text="Mô hình", bg=panel_bg, fg=text_sub, font=ui_font(11, bold=True)).pack(anchor="w", pady=(0, 6))
        model_menu = tk.Menu(win, tearoff=0, bg="#ffffff", fg="#111827", activebackground=UI_PRIMARY, activeforeground="#ffffff", font=ui_font(11))
        model_values = [
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash-lite",
            "gemini-2.5-flash",
            "gemini-1.5-flash",
            "gemini-1.5-pro",
        ]
        model_button = tk.Button(
            model_box,
            text=self.model_var.get() or DEFAULT_MODEL,
            command=lambda: model_menu.tk_popup(model_button.winfo_rootx(), model_button.winfo_rooty() + model_button.winfo_height()),
            bg=card_soft,
            fg=text_main,
            activebackground="#3f4351",
            activeforeground=text_main,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=card_border,
            font=ui_font(11),
            anchor="w",
            padx=12,
            pady=8,
        )
        model_button.pack(fill="x")
        for value in model_values:
            model_menu.add_command(label=value, command=lambda v=value: (self.model_var.set(v), model_button.config(text=v)))

        presence_box = tk.Frame(api_card, bg=panel_bg)
        presence_box.pack(fill="x", pady=(0, 16))
        server_controls = {}
        server_header = tk.Frame(presence_box, bg=panel_bg)
        server_header.pack(fill="x", pady=(0, 6))
        tk.Label(server_header, text="Server trạng thái máy", bg=panel_bg, fg=text_sub, font=ui_font(11, bold=True)).pack(side="left")
        presence_shell = tk.Frame(presence_box, bg=card_soft, highlightthickness=1, highlightbackground=card_border)
        presence_shell.pack(fill="x")
        presence_entry = tk.Entry(
            presence_shell,
            textvariable=self.presence_server_var,
            bg=card_soft,
            fg=text_main,
            insertbackground=text_main,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=ui_font(11),
        )
        presence_entry.pack(fill="x", padx=12, pady=9)

        def set_server_button_state(btn, text, bg, hover, border):
            if not btn:
                return
            btn.bg_color = bg
            btn.hover_color = hover
            btn.border_color = border
            btn.fg_color = "#ffffff"
            btn.press_color = btn._press_color(hover)
            btn.config(text=text)

        def refresh_server_status_label():
            btn = server_controls.get("toggle")
            if self._presence_server_is_running():
                set_server_button_state(btn, "Tắt server", "#dc2626", "#b91c1c", "#dc2626")
            else:
                set_server_button_state(btn, "Bật server", "#16a34a", "#15803d", "#16a34a")

        def toggle_server():
            stopping_server = self._presence_server_is_running()
            if stopping_server:
                ok, msg = self._stop_presence_server()
                error_title = "Không tắt được server"
            else:
                server_url = resolve_presence_server_url(self.presence_server_var.get())
                if server_url:
                    self.presence_server_var.set(server_url)
                ok, msg = self._ensure_presence_server_for_url(self.presence_server_var.get())
                error_title = "Không khởi động được server"
            refresh_server_status_label()
            if ok:
                self._set_status(msg, "error" if stopping_server else "success")
            else:
                self._set_status(msg, "error")
            if not ok:
                messagebox.showerror(error_title, msg)

        server_controls["toggle"] = ui_button(server_header, "Bật server", toggle_server, width=10, variant="primary")
        server_controls["toggle"].pack(side="right")
        refresh_server_status_label()

        note = tk.Frame(api_card, bg=card_soft, highlightthickness=1, highlightbackground=card_border, padx=14, pady=12)
        note.pack(fill="x")
        tk.Label(note, text="Lưu ý", bg=card_soft, fg=text_main, font=ui_font(11, bold=True)).pack(anchor="w")
        tk.Label(note, text="Lưu để áp dụng ngay cấu hình hiển thị, API và server trong cửa sổ hiện tại.", bg=card_soft, fg=text_sub, font=ui_font(10), wraplength=620, justify="left").pack(anchor="w", pady=(4, 0))
    else:
        show_page("display")

    def return_to_previous_page():
        try:
            self.show_page(return_page)
        except Exception:
            try:
                self.show_home_page()
            except Exception:
                pass

    def resolve_selected_profile():
        mode_name = mode_var.get()
        if mode_name == "Tự động":
            return "auto"
        if mode_name == "Tùy chỉnh":
            try:
                width = max(320, int(str(custom_width_var.get()).strip()))
                height = max(240, int(str(custom_height_var.get()).strip()))
                dpi = max(72, int(str(custom_dpi_var.get()).strip() or "240"))
            except Exception:
                messagebox.showerror("Dữ liệu không hợp lệ", "Nhập đúng số cho rộng, cao và DPI.")
                return None
            return f"custom:{width}x{height}@{dpi}"
        selected = str(choice_var.get() or "").strip()
        if selected:
            return selected
        return "auto"

    def save():
        profile_key = resolve_selected_profile()
        if not profile_key:
            return
        server_url = resolve_presence_server_url(self.presence_server_var.get())
        self.screen_profile_var.set(profile_key)
        if settings_is_admin:
            self.presence_server_var.set(server_url or DEFAULT_PRESENCE_SERVER_URL)
            save_env(self.api_key_var.get(), self.model_var.get(), profile_key, self.presence_server_var.get())
            ok, msg = self._ensure_presence_server_for_url(self.presence_server_var.get())
            try:
                self._set_status(msg, "success" if ok else "error")
            except Exception:
                pass
            if not ok:
                messagebox.showerror("Không khởi động được server", msg)
                return
        else:
            save_env(self.api_key_var.get(), self.model_var.get(), profile_key)
        self._dialog_return_page = None
        win.destroy()
        self._rebuild_ui_after_settings_change(return_page)

    win.protocol("WM_DELETE_WINDOW", lambda: (return_to_previous_page(), setattr(self, "_dialog_return_page", None), win.destroy()))

    cancel_btn = dark_button(footer, "Hủy", lambda: (return_to_previous_page(), setattr(self, "_dialog_return_page", None), win.destroy()), width=12)
    cancel_btn.pack(side="right")
    save_btn = dark_button(footer, "Lưu & áp dụng ngay", save, width=21, accent_button=True)
    save_btn.pack(side="right", padx=(0, 10))

    show_page("display")
    try:
        self._center_dialog_on_screen(win)
        win.lift()
        win.focus_force()
        mode_buttons.get(initial_mode, win).focus_set()
    except Exception:
        pass
    self.root.wait_window(win)


def show_help_dialog(self, event=None):
    return_page = getattr(self, "_dialog_return_page", None)
    if return_page not in {"home", "excel", "history", "mapping"}:
        return_page = getattr(self, "current_page", "home")

    messagebox.showinfo(

        "Hướng dẫn sử dụng GK PilePro",

        "Quy trình sử dụng cơ bản:\n\n"

        "1. Click 'Cài đặt' ở thanh bên để chỉnh độ hiển thị của ứng dụng.\n"

        "2. Bấm 'Chọn Excel' để tải tệp dữ liệu cọc của bạn lên.\n"

        "3. Kéo thả ảnh hoặc chọn tệp ảnh phiếu cọc cần số hóa ở khung bên trái.\n"

        "4. Bấm 'Đọc bảng' hoặc 'Đọc phiếu cọc' tùy thuộc vào loại biểu mẫu của bạn.\n"

        "5. Kiểm tra kỹ dữ liệu trích xuất được ở khung preview, chỉnh sửa trực tiếp nếu có sai sót.\n"

        "6. Thực hiện ánh xạ (mapping) các cột ở khung bên phải.\n"

        "7. Bấm 'Điền vào Excel' để ghi trực tiếp dữ liệu vào tệp Excel của bạn."

    )
    try:
        self.show_page(return_page)
    except Exception:
        self.show_home_page()
    finally:
        self._dialog_return_page = None


def show_about_dialog(self, event=None):
    return_page = getattr(self, "_dialog_return_page", None)
    if return_page not in {"home", "excel", "history", "mapping"}:
        return_page = getattr(self, "current_page", "home")

    messagebox.showinfo(

        "Giới thiệu GK PilePro",

        "GK PilePro - Hệ thống số hóa nhật ký và phục hồi dữ liệu cọc bằng AI.\n\n"

        "Phiên bản: V2.3.1 (Chuẩn hóa dữ liệu)\n"

        "Bản quyền © 2026 GK PilePro Team. Bảo lưu mọi quyền.\n"

        "Ứng dụng được thiết kế giúp tự động hóa quá trình xử Khối Lượng và Phiếu Cọc xây dựng."

    )
    try:
        self.show_page(return_page)
    except Exception:
        self.show_home_page()
    finally:
        self._dialog_return_page = None


def install_settings_ui(app_cls):
    app_cls._rebuild_ui_after_settings_change = _rebuild_ui_after_settings_change
    app_cls.show_settings_dialog = show_settings_dialog
    app_cls.show_help_dialog = show_help_dialog
    app_cls.show_about_dialog = show_about_dialog
