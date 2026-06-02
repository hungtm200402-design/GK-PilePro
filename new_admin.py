    def open_admin_approval_panel(self):
        if self.admin_approval_panel is not None:
            try:
                self.admin_approval_panel.lift()
                self.admin_approval_panel.focus_set()
            except Exception:
                pass
            return

        panel = tk.Frame(self.root, bg=UI_BG)
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

        sidebar_w = getattr(self, "sidebar_w", scale_px(190)) if hasattr(self, 'sidebar_w') else 190
        sidebar = tk.Frame(panel, width=sidebar_w, bg="#f8fbff", highlightthickness=1, highlightbackground="#e7edf6")
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        brand = tk.Frame(sidebar, bg="#f8fbff")
        brand.pack(fill="x", pady=(0, 24))
        try:
            if hasattr(self, 'app_logo_img') and self.app_logo_img is not None:
                tk.Label(brand, image=self.app_logo_img, bg="#f8fbff").pack(anchor="center", pady=(14, 0))
        except Exception:
            pass

        nav_items = [
            ("home", "⌂", "Trang chủ", False),
            ("excel", "▦", "Excel", False),
            ("history", "◷", "Lịch sử", False),
            ("mapping", "▤", "Mẫu mapping", False),
            ("settings", "⚙", "Cài đặt", False),
            ("help", "?", "Trợ giúp", False),
            ("about", "i", "Giới thiệu", False),
        ]

        for page_id, icon, text, active in nav_items:
            bg = "#f8fbff"
            fg = "#667085"
            row = tk.Frame(sidebar, bg=bg, padx=0, pady=0, highlightthickness=1, highlightbackground=bg, cursor="hand2")
            row.pack(fill="x", pady=2)
            inner = tk.Frame(row, bg=bg, padx=8, pady=8, cursor="hand2")
            inner.pack(fill="both", expand=True)
            tk.Label(inner, text=icon, font=ui_font(12), bg=bg, fg=fg, width=2, anchor="center", cursor="hand2").pack(side="left", padx=(0, 6))
            lbl = tk.Label(inner, text=text, font=ui_font(10, bold=active), bg=bg, fg=fg, anchor="w", cursor="hand2")
            lbl.pack(side="left", fill="x", expand=True)
            for widget in (row, inner, lbl):
                widget.bind("<Button-1>", lambda e: close_panel())

        member_info = tk.Frame(sidebar, bg="#f8fbff")
        member_info.pack(side="bottom", fill="x", pady=(0, 24))
        tk.Label(member_info, text="Quản trị viên", font=ui_font(10, bold=True), bg="#f8fbff", fg=UI_TEXT).pack(anchor="center")
        tk.Label(member_info, text="Admin", font=ui_font(9), bg="#f8fbff", fg=UI_MUTED).pack(anchor="center", pady=(2, 10))
        active_btn = tk.Frame(member_info, bg="#fff7ed", highlightthickness=1, highlightbackground="#fed7aa")
        active_btn.pack(pady=4, padx=12, fill="x")
        tk.Label(active_btn, text="Duyệt máy ✓", font=ui_font(9, bold=True), bg="#fff7ed", fg="#ea580c").pack(pady=6)

        right_panel = tk.Frame(panel, bg=UI_BG, width=280)
        right_panel.pack(side="right", fill="y", padx=(10, 20), pady=20)
        right_panel.pack_propagate(False)

        right_card = tk.Frame(right_panel, bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER)
        right_card.pack(fill="x")
        
        tk.Label(right_card, text="Tổng quan hệ thống", bg="#ffffff", fg=UI_TEXT, font=ui_font(11, bold=True)).pack(anchor="w", padx=16, pady=(16, 12))

        def add_stat_row(parent, icon, title, value_var, icon_color, icon_bg):
            row = tk.Frame(parent, bg="#ffffff")
            row.pack(fill="x", padx=16, pady=8)
            icon_lbl = tk.Label(row, text=icon, font=ui_font(12, bold=True), bg=icon_bg, fg=icon_color, width=2, height=1)
            icon_lbl.pack(side="left", padx=(0, 12))
            text_frame = tk.Frame(row, bg="#ffffff")
            text_frame.pack(side="left", fill="x", expand=True)
            tk.Label(text_frame, text=title, bg="#ffffff", fg=UI_MUTED, font=ui_font(9)).pack(anchor="w")
            tk.Label(text_frame, textvariable=value_var, bg="#ffffff", fg=UI_TEXT, font=ui_font(12, bold=True)).pack(anchor="w")

        stat_total_var = tk.StringVar(value="0")
        stat_active_var = tk.StringVar(value="0")
        stat_pending_var = tk.StringVar(value="0")
        stat_blocked_var = tk.StringVar(value="0")
        stat_time_var = tk.StringVar(value="Cập nhật: --:--:--")

        add_stat_row(right_card, "📄", "Máy đã duyệt", stat_total_var, "#2563eb", "#eff6ff")
        add_stat_row(right_card, "🛡", "Đang hoạt động", stat_active_var, "#16a34a", "#f0fdf4")
        add_stat_row(right_card, "⏳", "Chờ duyệt", stat_pending_var, "#d97706", "#fffbeb")
        add_stat_row(right_card, "🚫", "Đã chặn", stat_blocked_var, "#dc2626", "#fef2f2")

        time_frame = tk.Frame(right_card, bg="#ffffff")
        time_frame.pack(fill="x", padx=16, pady=(16, 16))
        tk.Label(time_frame, textvariable=stat_time_var, bg="#ffffff", fg=UI_MUTED, font=ui_font(8)).pack(side="left")
        refresh_icon = tk.Label(time_frame, text="↻", bg="#ffffff", fg=UI_PRIMARY, font=ui_font(12, bold=True), cursor="hand2")
        refresh_icon.pack(side="right")
        refresh_icon.bind("<Button-1>", lambda e: clear_search())

        main_content = tk.Frame(panel, bg=UI_BG)
        main_content.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        header_frame = tk.Frame(main_content, bg=UI_BG)
        header_frame.pack(fill="x", pady=(0, 20))
        
        title_row = tk.Frame(header_frame, bg=UI_BG)
        title_row.pack(fill="x")
        tk.Label(title_row, text="👥", bg=UI_BG, fg=UI_PRIMARY, font=ui_font(16)).pack(side="left", padx=(0, 8))
        tk.Label(title_row, text="Duyệt máy thành viên", bg=UI_BG, fg=UI_TEXT, font=ui_font(16, bold=True)).pack(side="left")
        
        tk.Label(header_frame, text="Nhập mã máy do thành viên cung cấp để tạo mã duyệt cho họ.", bg=UI_BG, fg=UI_MUTED, font=ui_font(10)).pack(anchor="w", pady=(4, 0), padx=(36, 0))

        form_box = tk.Frame(main_content, bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER)
        form_box.pack(fill="x", pady=(0, 20))
        
        tk.Label(form_box, text="🗝 Tạo mã duyệt mới", bg="#ffffff", fg=UI_PRIMARY, font=ui_font(11, bold=True)).pack(anchor="w", padx=20, pady=(16, 12))

        form_grid = tk.Frame(form_box, bg="#ffffff")
        form_grid.pack(fill="x", padx=20)

        col1 = tk.Frame(form_grid, bg="#ffffff")
        col1.pack(side="left", fill="x", expand=True, padx=(0, 16))
        tk.Label(col1, text="Mã máy", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
        machine_var = tk.StringVar()
        machine_entry = tk.Entry(col1, textvariable=machine_var, relief="solid", bd=1, font=ui_font(11), bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER, highlightcolor=UI_PRIMARY)
        machine_entry.pack(fill="x", pady=(6, 12))

        col2 = tk.Frame(form_grid, bg="#ffffff")
        col2.pack(side="left", fill="x", expand=True)
        tk.Label(col2, text="Tên người (tuỳ chọn)", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w")
        name_var = tk.StringVar()
        name_entry = tk.Entry(col2, textvariable=name_var, relief="solid", bd=1, font=ui_font(11), bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER, highlightcolor=UI_PRIMARY)
        name_entry.pack(fill="x", pady=(6, 12))

        tk.Label(form_box, text="Mã duyệt", bg="#ffffff", fg=UI_TEXT, font=ui_font(10, bold=True)).pack(anchor="w", padx=20)
        approval_var = tk.StringVar()
        approval_entry = tk.Entry(form_box, textvariable=approval_var, relief="flat", bd=0, font=ui_font(12, bold=True), fg="#0369a1", bg="#f0f9ff", highlightthickness=1, highlightbackground="#bae6fd")
        approval_entry.pack(fill="x", padx=20, pady=(6, 16), ipady=4)
        approval_entry.configure(state="readonly")

        last_generated_code = {"value": ""}

        actions = tk.Frame(form_box, bg="#ffffff")
        actions.pack(fill="x", padx=20, pady=(0, 20))

        list_box = tk.Frame(main_content, bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER)
        list_box.pack(fill="both", expand=True)

        list_header = tk.Frame(list_box, bg="#ffffff")
        list_header.pack(fill="x", padx=20, pady=(16, 12))

        list_title_box = tk.Frame(list_header, bg="#ffffff")
        list_title_box.pack(side="left")

        tk.Label(list_title_box, text="📄 Danh sách máy đã duyệt", bg="#ffffff", fg=UI_PRIMARY, font=ui_font(11, bold=True)).pack(anchor="w")
        summary_var = tk.StringVar(value="Đang tải...")
        tk.Label(list_title_box, textvariable=summary_var, bg="#ffffff", fg=UI_MUTED, font=ui_font(9)).pack(anchor="w", pady=(2, 0))

        search_bar = tk.Frame(list_box, bg="#ffffff")
        search_bar.pack(fill="x", padx=20, pady=(0, 12))
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_bar, textvariable=search_var, relief="solid", bd=1, font=ui_font(10), bg="#ffffff", highlightthickness=1, highlightbackground=UI_BORDER, highlightcolor=UI_PRIMARY)
        search_entry.pack(side="left", fill="x", expand=True, ipady=3)
        search_entry.insert(0, " Tìm kiếm theo mã máy, tên người hoặc mã duyệt...")
        search_entry.bind("<FocusIn>", lambda e: search_entry.delete(0, 'end') if search_var.get().startswith(" Tìm kiếm") else None)
        search_entry.bind("<FocusOut>", lambda e: search_entry.insert(0, " Tìm kiếm theo mã máy, tên người hoặc mã duyệt...") if not search_var.get() else None)

        form_inputs = {machine_entry, name_entry, approval_entry, search_entry}

        list_frame = tk.Frame(list_box, bg="#ffffff")
        list_frame.pack(fill="both", expand=True, padx=20)

        style = ttk.Style()
        style.configure("Custom.Treeview", background="#ffffff", fieldbackground="#ffffff", rowheight=36, borderwidth=0, font=ui_font(10))
        style.configure("Custom.Treeview.Heading", font=ui_font(9, bold=True), background="#f8fbff", foreground="#667085")
        style.layout("Custom.Treeview", [('Custom.Treeview.treearea', {'sticky': 'nswe'})])

        approved_tree = ttk.Treeview(list_frame, style="Custom.Treeview", columns=("machine", "user", "code", "time", "status", "action"), show="headings", height=8)
        approved_tree.heading("machine", text=" Mã máy")
        approved_tree.heading("user", text="Tên người")
        approved_tree.heading("code", text="Mã duyệt")
        approved_tree.heading("time", text="Thời gian duyệt")
        approved_tree.heading("status", text="Trạng thái")
        approved_tree.heading("action", text="Thao tác")

        approved_tree.column("machine", width=200, anchor="w")
        approved_tree.column("user", width=120, anchor="center")
        approved_tree.column("code", width=150, anchor="center")
        approved_tree.column("time", width=140, anchor="center")
        approved_tree.column("status", width=120, anchor="center")
        approved_tree.column("action", width=80, anchor="center")

        approved_tree.tag_configure("online", foreground="#16a34a")
        approved_tree.tag_configure("away", foreground="#d97706")
        approved_tree.tag_configure("old", foreground="#6b7280")
        approved_tree.tag_configure("stripe", background="#fafafa")

        approved_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=approved_tree.yview)
        approved_tree.configure(yscrollcommand=approved_scroll.set)
        approved_tree.grid(row=0, column=0, sticky="nsew")
        approved_scroll.grid(row=0, column=1, sticky="ns")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        list_actions = tk.Frame(list_box, bg="#ffffff")
        list_actions.pack(fill="x", padx=20, pady=(16, 20))

        def filter_rows(rows, query):
            query = str(query or "").strip()
            if not query or query.startswith(" Tìm kiếm"):
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

            import_local_approval_to_admin_list()
            server_rows = self._get_presence_machine_cache()
            sync_presence_machines_to_admin_list(server_rows)
            all_rows = load_admin_approved_machines()
            rows = filter_rows(all_rows, search_var.get())
            server_map = {}
            for item in server_rows:
                if not isinstance(item, dict):
                    continue
                machine_key = str(item.get("machine_code") or "").strip().upper()
                if machine_key:
                    server_map[machine_key] = item

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
                else:
                    dt = parse_machine_datetime(last_seen)
                    if dt is None:
                        status_tag = "old"
                    else:
                        age_min = max(0, int((datetime.now() - dt).total_seconds() // 60))
                        status_tag = "away" if age_min < 180 else "old"

                tags = (status_tag, "stripe" if i % 2 == 1 else "")
                values = ("  " + machine, str(row.get("user_name", "") or "").strip(), row.get("approval_code", ""), row.get("approved_at", ""), status_text, "⋮")

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
            
            # Update Right Panel stats
            stat_total_var.set(str(total_count))
            stat_active_var.set(str(online_count))
            stat_pending_var.set("0")
            stat_blocked_var.set("0")
            stat_time_var.set(f"Cập nhật: {datetime.now().strftime('%H:%M:%S %d/%m/%Y')}")

            target = str(select_machine or "").strip().upper()
            children = approved_tree.get_children()
            selected_id = target if target and approved_tree.exists(target) else (children[0] if children else "")

            if selected_id:
                approved_tree.focus(selected_id)
                approved_tree.see(selected_id)
                if fill_selection:
                    approved_tree.selection_set(selected_id)
                    fill_from_selected()
                else:
                    approved_tree.selection_remove(approved_tree.selection())

            try:
                if refresh_job["id"] is not None:
                    self.admin_approval_panel.after_cancel(refresh_job["id"])
            except Exception:
                pass
            try:
                if getattr(self, "admin_approval_panel", None) is not None and self.admin_approval_panel.winfo_exists():
                    refresh_job["id"] = self.admin_approval_panel.after(5000, lambda: refresh_list(select_machine=select_machine, fill_selection=False, auto=True))
            except Exception:
                refresh_job["id"] = None

        def search_rows():
            refresh_list()
            search_entry.focus_set()

        def clear_search():
            search_var.set("")
            search_entry.delete(0, 'end')
            search_entry.insert(0, " Tìm kiếm theo mã máy, tên người hoặc mã duyệt...")
            refresh_list()

        def fill_from_selected(_event=None):
            selected = approved_tree.selection()
            if not selected:
                return
            values = approved_tree.item(selected[0], "values")
            if values:
                machine_var.set(str(values[0]).strip())
                name_var.set(values[1] if len(values) > 1 else "")
                approval_entry.configure(state="normal")
                approval_var.set(values[2] if len(values) > 2 else "")
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
                fill_from_selected()
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Xóa máy này", command=delete_selected)
            menu.tk_popup(event.x_root, event.y_root)

        approved_tree.bind("<Delete>", lambda _e: delete_selected())
        approved_tree.bind("<Button-3>", open_list_menu)

        ui_button(actions, "+ Tạo mã duyệt", generate, width=13, variant="primary").pack(side="left", padx=(0, 8))
        ui_button(actions, "❐ Copy mã", copy_code, width=10, variant="soft").pack(side="left")
        ui_button(actions, "✕ Đóng", close_panel, width=9).pack(side="right")

        ui_button(search_bar, "Tìm", search_rows, width=8, variant="primary").pack(side="left", padx=(8, 0))
        ui_button(search_bar, "Bộ lọc", search_rows, width=8, variant="soft").pack(side="left", padx=(8, 0))
        ui_button(search_bar, "🗑 Xóa máy", delete_selected, width=10, variant="warn").pack(side="right")

        ui_button(list_actions, "🗑 Xóa máy đã chọn", delete_selected, width=16, variant="warn").pack(side="left")
        ui_button(list_actions, "↻ Tải lại danh sách", clear_search, width=16, variant="soft").pack(side="left", padx=(8, 0))

        refresh_list(fill_selection=False)
        clear_approval_form()
