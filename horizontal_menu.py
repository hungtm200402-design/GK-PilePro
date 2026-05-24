import re

content = open('app.py', 'r', encoding='utf-8').read()

start_marker = 'sidebar = tk.Frame('
end_marker = 'main = tk.Frame(shell, bg=UI_BG, padx=self.main_padx, pady=self.main_pady)'

start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx == -1 or end_idx == -1:
    print('Could not find markers')
    exit(1)

new_sidebar_code = '''        sidebar = tk.Frame(
            shell,
            bg="#f8fbff",
            padx=16,
            pady=8,
            highlightthickness=1,
            highlightbackground="#e7edf6",
        )
        sidebar.pack(side="top", fill="x")
        
        brand = tk.Frame(sidebar, bg="#f8fbff")
        brand.pack(side="left", fill="y", padx=(0, 24))
        logo_file = resource_path(*APP_LOGO_PNG.parts)
        try:
            if logo_file.exists():
                logo_source = Image.open(logo_file).convert("RGBA")
                bbox = logo_source.getchannel("A").getbbox()
                if bbox:
                    logo_source = logo_source.crop(bbox)
                logo_source.thumbnail((100, 32), Image.LANCZOS)
                self.app_logo_img = ImageTk.PhotoImage(logo_source)
                tk.Label(brand, image=self.app_logo_img, bg="#f8fbff").pack(side="left", anchor="center")
        except Exception:
            self.app_logo_img = None

        nav_items = [
            ("home", "⌂", "Trang chủ", True),
            ("excel", "▦", "Excel", False),
            ("history", "◷", "Lịch sử", False),
            ("mapping", "▤", "Mẫu mapping", False),
            ("settings", "⚙", "Cài đặt", False),
            ("help", "?", "Trợ giúp", False),
            ("about", "i", "Giới thiệu", False),
        ]
        
        nav_container = tk.Frame(sidebar, bg="#f8fbff")
        nav_container.pack(side="left", fill="y")
        
        for page_id, icon, text, active in nav_items:
            bg = "#dbeafe" if active else "#f8fbff"
            fg = UI_PRIMARY if active else "#667085"
            row = tk.Frame(nav_container, bg=bg, padx=0, pady=0, highlightthickness=1, highlightbackground=UI_PRIMARY if active else "#e7edf6", cursor="hand2")
            row.pack(side="left", fill="y", padx=4)
            inner = tk.Frame(row, bg=bg, padx=10, pady=4, cursor="hand2")
            inner.pack(fill="both", expand=True)
            
            accent = tk.Frame(row, bg=UI_PRIMARY if active else bg, height=3)
            accent.pack(side="bottom", fill="x")
            
            icon_label = tk.Label(inner, text=icon, width=2, bg=bg, fg=fg, font=("Segoe UI", 11), cursor="hand2")
            icon_label.pack(side="left")
            text_label = tk.Label(inner, text=text, bg=bg, fg=fg, font=("Segoe UI", 9, "bold" if active else "normal"), cursor="hand2")
            text_label.pack(side="left", padx=(4, 0))
            
            self.nav_widgets[page_id] = {"row": row, "inner": inner, "accent": accent, "icon": icon_label, "label": text_label}
            if page_id == "excel":
                for widget in (row, inner, accent, icon_label, text_label):
                    widget.bind("<Button-1>", self.show_excel_page)
            elif page_id == "home":
                for widget in (row, inner, accent, icon_label, text_label):
                    widget.bind("<Button-1>", self.show_home_page)
            elif page_id == "settings":
                for widget in (row, inner, accent, icon_label, text_label):
                    widget.bind("<Button-1>", self.show_settings_dialog)
            elif page_id == "help":
                for widget in (row, inner, accent, icon_label, text_label):
                    widget.bind("<Button-1>", self.show_help_dialog)
            elif page_id == "about":
                for widget in (row, inner, accent, icon_label, text_label):
                    widget.bind("<Button-1>", self.show_about_dialog)

        user_box = tk.Frame(sidebar, bg="#f8fbff")
        user_box.pack(side="right", fill="y")
        
        user_info = tk.Frame(user_box, bg="#f8fbff")
        user_info.pack(side="right", fill="y")
        
        tk.Label(user_info, text=self.user_name, font=("Segoe UI", 9, "bold"), bg="#f8fbff", fg=UI_TEXT).pack(side="top", anchor="e")
        tk.Label(user_info, textvariable=self.user_role_var, font=("Segoe UI", 8), bg="#f8fbff", fg=UI_MUTED).pack(side="top", anchor="e")
        
        status_frame = tk.Frame(user_box, bg="#f8fbff")
        status_frame.pack(side="right", fill="y", padx=(0, 16))
        
        self.status = tk.Label(
            status_frame,
            text="● Sẵn sàng",
            anchor="e",
            fg=UI_SUCCESS,
            bg="#f8fbff",
            font=("Segoe UI", 8, "bold")
        )
        self.status.pack(side="top", fill="x", pady=(6, 0))
        
        if is_admin_build():
            ui_button(status_frame, "Duyệt máy", self.open_admin_approval_panel, width=12, variant="warn").pack(side="top", pady=(2, 0))

        '''

content = content[:start_idx] + new_sidebar_code + content[end_idx:]

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Sidebar updated successfully!')
