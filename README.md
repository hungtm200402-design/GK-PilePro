# GK PilePro

Ung dung phuc hoi, doc OCR va quan ly du lieu coc.

## Chay ban da build

Ban User:

```bat
dist\Tool_KL_UI8.exe
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
app.py                 Source chinh cua ung dung
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
