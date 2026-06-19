# 01 User Client

## Pham vi

Giao dien user, OCR flow, mapping, preview, xuat Excel, lich su, thanh cap nhat, lock khi may chua duyet.

## File duoc sua

Da copy san trong `source/`:

- `app.py`
- `gk_pilepro/ui/gk_ocr_ui.py`
- `gk_pilepro/ui/gk_excel_ui.py`
- `gk_pilepro/ui/gk_history_ui.py`
- `gk_pilepro/ui/gk_editors.py`
- `gk_pilepro/ui/gk_settings_ui.py`
- `gk_pilepro/ui/gk_ui.py`
- `gk_pilepro/ui/gk_icons.py`

## Khong tu y sua

- `presence_server.py`
- `presence_server.spec`
- `Tool_KL.spec`, tru khi them/xoa asset/dependency cho user.
- `GK_PilePro_Admin.spec`

## Can phoi hop khi

- Sua heartbeat, approval, update endpoint: hoi `Presence/Update Server`.
- Sua mapping/export Excel: hoi `Excel/Data Files`.
- Them/xoa icon/logo/anh: hoi `Assets/UI Resources` va `Build System`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile app.py
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
```

Exe output:

```text
dist\GK PilePro.exe
```
