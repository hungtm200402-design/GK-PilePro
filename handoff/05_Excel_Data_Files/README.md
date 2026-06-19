# 05 Excel/Data Files

## Pham vi

Doc workbook, doc sheet, doc cong thuc, mapping, preview, ghi Excel, backup, validate du lieu OCR truoc khi xuat.

## File duoc sua

Da copy san trong `source/`:

- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`
- `gk_pilepro/ui/gk_excel_ui.py`
- `gk_pilepro/ui/gk_editors.py`

## Khong tu y sua

- `presence_server.py`
- Admin approval UI.
- Build spec neu khong them dependency.

## Can phoi hop khi

- Doi cau truc OCR: hoi `Gemini OCR Service`.
- Doi giao dien preview/mapping: hoi `User Client`.
- Doi noi luu backup/mapping/history: hoi `Local App Data`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile app.py gk_pilepro\gk_excel.py gk_pilepro\gk_overrides.py
```
