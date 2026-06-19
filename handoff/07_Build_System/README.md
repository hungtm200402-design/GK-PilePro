# 07 Build System

## Pham vi

Build exe user/admin/server, dependency, asset bundle, PyInstaller hook, batch script.

## File duoc sua

Da copy san trong `source/`:

- `Tool_KL.spec`
- `GK_PilePro_Admin.spec`
- `presence_server.spec`
- `requirements.txt`
- `install.bat`
- `run.bat`
- `run_presence_server.bat`
- `test.bat`
- `pyi_hooks/`

## Khong tu y sua

- Business logic trong `gk_pilepro/`.
- UI logic trong `app.py`, tru khi build can doi path/asset constant.

## Can phoi hop khi

- Them/xoa asset: hoi `Assets/UI Resources`.
- Them dependency OCR/API/Excel: hoi mang lien quan.
- Build loi do file exe dang chay: dung process truoc khi build lai.

## Build

```bat
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
.\.venv\Scripts\pyinstaller.exe presence_server.spec
```

Output:

```text
dist\GK PilePro.exe
dist\GK PilePro Admin.exe
dist\presence_server.exe
```
