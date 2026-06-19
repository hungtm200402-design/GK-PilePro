# GK PilePro

Ung dung phuc hoi, doc OCR va quan ly du lieu coc.

## Chay ban da build

Ban User:

```bat
dist\GK PilePro.exe
```

Ban Admin:

```bat
dist\GK PilePro Admin.exe
```

## Chay tu source

May can co Python 3.12.

Lan dau:

```bat
install.bat
```

Nhung lan sau:

```bat
run.bat
```

## Build exe

Build ban User:

```bat
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
```

Build ban Admin:

```bat
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
```

## Cau hinh API

Tao file `.env` tu `.env.example`, roi dien API key/model neu can.

```bat
copy .env.example .env
```

File `.env` khong duoc dua len Git.

## Cau truc thu muc

```text
app.py                 Entry point va luong dieu phoi ung dung
gk_pilepro/            Package code chinh cua app
  gk_core.py           Runtime, cau hinh, approval va presence client
  gk_excel.py          Xu ly OCR, mapping, Excel va formula
  gk_overrides.py      Override logic workbook/preview/run Gemini dang ap dung cho App
  ui/                  Package giao dien
    gk_ui.py           UI constants, font scale va widget dung chung
    gk_icons.py        Helper icon, DPI va work-area Windows
    gk_editors.py      MappingEditor va TableEditor
    gk_settings_ui.py  Settings, help va about dialogs
    gk_admin_ui.py     Admin log va approval panels
    gk_history_ui.py   History views va history detail rendering
    gk_excel_ui.py     Excel page, mapping templates va workbook UI actions
    gk_ocr_ui.py       OCR workflow, image preview va clipboard actions
assets/                Logo, icon va tai nguyen giao dien
dist/                  File exe da build de giao cho khach
docs/                  Tai lieu va ghi chu phien ban
logs/                  Log va ket qua debug
tools/                 Script ho tro sua loi/kiem tra noi bo
backups/               File backup cu, khong dung de chay app
specs/legacy/          Spec build cu de tham khao
pyi_hooks/             Hook PyInstaller
runtime_data/          Du lieu tam/legacy da tach khoi root
```

## Phan chia pham vi lam viec

Xem `handoff/` de gui source cho tung nguoi theo mang. Moi folder trong `handoff/` co README rieng cho nguoi phu trach.

Tai lieu tong hop nam o `docs/OWNERSHIP.md`.

- User Client
- Admin Client
- Presence/Update Server
- Gemini OCR Service
- Excel/Data Files
- Local App Data
- Build System
- Assets/UI Resources
