# GK PilePro Ownership Map

Tai lieu nay chia pham vi code theo tung ben/pham vi cong viec. Khi giao cho mot nguoi lam mot mang, nguoi do chi nen sua cac file trong pham vi cua mang do. Neu can sua file ngoai pham vi, phai bao lai nguoi phu trach mang lien quan.

## Nguyen tac chung

- Khong sua `dist/` bang tay. File exe trong `dist/` chi duoc tao tu spec build.
- Khong xoa hoac doi ten asset neu chua kiem tra `app.py`, `gk_pilepro/`, va cac file `.spec`.
- Khong sua server khi chi duoc giao client, tru khi thay doi API da duoc thong nhat.
- Khong sua build spec khi chi thay doi UI/logic, tru khi co them/bot asset hoac dependency.
- Moi thay doi lien quan update/user/admin nen build lai dung exe tuong ung de kiem tra.

## User Client

Pham vi: giao dien user, OCR flow, mapping, preview, xuat Excel, lich su, update popup, lock khi chua duyet.

File chinh:

- `app.py`
- `gk_pilepro/ui/gk_ocr_ui.py`
- `gk_pilepro/ui/gk_excel_ui.py`
- `gk_pilepro/ui/gk_history_ui.py`
- `gk_pilepro/ui/gk_editors.py`
- `gk_pilepro/ui/gk_settings_ui.py`
- `gk_pilepro/ui/gk_ui.py`
- `gk_pilepro/ui/gk_icons.py`

Can phoi hop voi:

- `Presence/Update Server` neu sua heartbeat, approval, update endpoint.
- `Excel/Data Files` neu sua mapping, preview, export Excel.
- `Assets/UI Resources` neu them logo/icon/anh giao dien.

Kiem tra/build:

```bat
.\.venv\Scripts\python.exe -m py_compile app.py
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
```

## Admin Client

Pham vi: duyet may, danh sach user online, log loi user gui len, bat/tat server, UI admin.

File chinh:

- `app.py`
- `gk_pilepro/ui/gk_admin_ui.py`
- `gk_pilepro/ui/gk_settings_ui.py`
- `gk_pilepro/gk_core.py`

Can phoi hop voi:

- `Presence/Update Server` neu sua API server, approved machines, error logs.
- `Build System` neu thay asset/dependency cua admin.

Kiem tra/build:

```bat
.\.venv\Scripts\python.exe -m py_compile app.py presence_server.py
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
```

## Presence/Update Server

Pham vi: server noi bo, heartbeat, approved machines, error logs, update info, tai exe user moi.

File chinh:

- `presence_server.py`
- `presence_server.spec`
- `gk_pilepro/gk_core.py`

API/endpoint lien quan:

- `/health`
- `/heartbeat`
- `/machines`
- `/approved-machines`
- `/error-logs`
- `/update-info`
- `/updates/user-exe`

Can phoi hop voi:

- `User Client` neu doi response hoac hanh vi update.
- `Admin Client` neu doi duyet may/log loi.
- `Build System` neu doi cach dong goi server.

Kiem tra/build:

```bat
.\.venv\Scripts\python.exe -m py_compile presence_server.py
.\.venv\Scripts\pyinstaller.exe presence_server.spec
```

## Gemini OCR Service

Pham vi: prompt Gemini, goi model, parse JSON, chuan hoa bang OCR, fallback model.

File chinh:

- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`
- `gk_pilepro/ui/gk_ocr_ui.py`

Can phoi hop voi:

- `User Client` neu thay nut/luong doc anh.
- `Excel/Data Files` neu thay cau truc bang OCR tra ve.
- `Local App Data` neu thay log/history OCR.

Kiem tra/build:

```bat
.\.venv\Scripts\python.exe -m py_compile app.py gk_pilepro\gk_excel.py gk_pilepro\gk_overrides.py
```

## Excel/Data Files

Pham vi: doc workbook, sheet, cong thuc, mapping, ghi Excel, backup, validate du lieu.

File chinh:

- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`
- `gk_pilepro/ui/gk_excel_ui.py`
- `gk_pilepro/ui/gk_editors.py`

Can phoi hop voi:

- `Gemini OCR Service` neu thay cau truc OCR.
- `User Client` neu thay preview/mapping UI.
- `Local App Data` neu thay noi luu mapping/history/backup.

Kiem tra:

```bat
.\.venv\Scripts\python.exe -m py_compile app.py gk_pilepro\gk_excel.py gk_pilepro\gk_overrides.py
```

## Local App Data

Pham vi: noi luu settings, history, mapping templates, approval cache, logs, backup path.

File chinh:

- `gk_pilepro/gk_core.py`
- Cac noi goi `app_data_path(...)`

Du lieu lien quan:

- `tool_kl_settings.json`
- `tool_kl_history.json`
- `tool_kl_mapping_templates.json`
- `gk_pilepro_approval.json`
- `gk_pilepro_approved_machines.json`
- `logs/`
- `backups/`

Can phoi hop voi:

- `User Client` neu thay settings/history UI.
- `Admin Client` neu thay approved machine data.
- `Excel/Data Files` neu thay backup/mapping templates.

## Build System

Pham vi: build exe user/admin/server, dependency, asset bundle, PyInstaller.

File chinh:

- `Tool_KL.spec`
- `GK_PilePro_Admin.spec`
- `presence_server.spec`
- `requirements.txt`
- `install.bat`
- `run.bat`
- `test.bat`
- `pyi_hooks/`

Can phoi hop voi:

- `Assets/UI Resources` neu them/xoa asset.
- Tat ca cac mang neu dependency moi can duoc bundle.

Build:

```bat
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
.\.venv\Scripts\pyinstaller.exe presence_server.spec
```

## Assets/UI Resources

Pham vi: logo, icon, splash, sidebar icon, button icon, anh trang tri.

File/thu muc chinh:

- `assets/`
- Cac hang asset trong `app.py`
- Cac muc `datas` trong `Tool_KL.spec` va `GK_PilePro_Admin.spec`

Can phoi hop voi:

- `User Client` neu thay UI user.
- `Admin Client` neu thay UI admin.
- `Build System` neu them/xoa file asset can dong goi.

Kiem tra:

```bat
rg -n "ten_file_asset.png" app.py gk_pilepro Tool_KL.spec GK_PilePro_Admin.spec
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
```

## Bang phoi hop nhanh

| Mang can sua | File nen sua | File can tranh sua neu khong can |
| --- | --- | --- |
| User Client | `app.py`, `gk_pilepro/ui/*` | `presence_server.py`, `*.spec` |
| Admin Client | `gk_admin_ui.py`, `gk_core.py`, `app.py` | OCR prompt, Excel export core neu khong lien quan |
| Presence/Update Server | `presence_server.py`, `gk_core.py` | UI files |
| Gemini OCR Service | `gk_excel.py`, `gk_overrides.py`, `gk_ocr_ui.py` | Server/build spec |
| Excel/Data Files | `gk_excel.py`, `gk_overrides.py`, `gk_excel_ui.py`, `gk_editors.py` | Server endpoint |
| Local App Data | `gk_core.py` | UI layout neu khong can |
| Build System | `*.spec`, `requirements.txt`, bat files | Business logic |
| Assets/UI Resources | `assets/`, asset constants/spec datas | Server/OCR logic |
