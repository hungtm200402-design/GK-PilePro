import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

start_marker = 'toolbar = card(self.home_page,'
end_marker = 'filters = card(self.home_page, padx=14, pady=9)'

start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx == -1 or end_idx == -1:
    print('Could not find markers')
    exit(1)

new_toolbar_code = '''        toolbar = card(self.home_page, padx=12, pady=10)
        toolbar.pack(fill="x", pady=(12, 10))
        
        def make_group(parent, title, btns, bg_color):
            group = tk.Frame(parent, bg=bg_color)
            group.pack(side="left", padx=(0, 16))
            if title:
                tk.Label(group, text=title, font=("Segoe UI", 8, "bold"), fg=UI_MUTED, bg=bg_color).pack(side="top", anchor="w", pady=(0, 4))
            btn_row = tk.Frame(group, bg=bg_color)
            btn_row.pack(side="top", fill="x")
            for text, command, variant in btns:
                ui_button(btn_row, text, command, width=0, variant=variant).pack(side="left", padx=(0, 4))
            return group

        source_btns = [
            ("Chọn Excel", self.choose_excel, "primary"),
            ("Đọc workbook", self.scan_current_workbook, "default"),
            ("Đọc từng sheet", self.read_each_sheet_content, "default"),
            ("Đọc công thức", self.read_current_excel_formulas, "default"),
            ("Đọc lại", self.refresh_excel_header_info, "soft"),
        ]
        make_group(toolbar, "NGUỒN DỮ LIỆU", source_btns, UI_SURFACE)
        
        separator1 = tk.Frame(toolbar, bg="#e2e8f0", width=1)
        separator1.pack(side="left", fill="y", padx=(0, 16), pady=4)

        process_btns = [
            ("Đọc bảng", self.run_gemini, "soft"),
            ("Đọc phiếu cọc", self.run_gemini_phieu_coc, "soft"),
            ("Auto map", self.build_mapping, "warn"),
        ]
        make_group(toolbar, "XỬ LÝ DỮ LIỆU", process_btns, UI_SURFACE)
        
        separator2 = tk.Frame(toolbar, bg="#e2e8f0", width=1)
        separator2.pack(side="left", fill="y", padx=(0, 16), pady=4)

        export_btns = [
            ("Xem trước", self.preview_excel, "soft"),
            ("Xuất ra Excel", self.fill_excel, "success"),
        ]
        make_group(toolbar, "XEM & XUẤT", export_btns, UI_SURFACE)

        '''

content = content[:start_idx] + new_toolbar_code + content[end_idx:]

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Toolbar updated successfully!')
