# 02 Admin Client

## Pham vi

Duyet may user, danh sach user online/offline, log loi user gui len, bat/tat server, UI admin.

## File duoc sua

Da copy san trong `source/`:

- `app.py`
- `gk_pilepro/ui/gk_admin_ui.py`
- `gk_pilepro/ui/gk_settings_ui.py`
- `gk_pilepro/gk_core.py`

## Khong tu y sua

- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`
- `presence_server.py`, tru khi doi API admin-server.
- `Tool_KL.spec`

## Can phoi hop khi

- Doi approved machines, error logs, heartbeat: hoi `Presence/Update Server`.
- Doi logo/icon/admin assets: hoi `Assets/UI Resources`.
- Doi spec/dependency admin: hoi `Build System`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile app.py presence_server.py
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
```

Exe output:

```text
dist\GK PilePro Admin.exe
```
