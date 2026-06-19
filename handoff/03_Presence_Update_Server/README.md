# 03 Presence/Update Server

## Pham vi

Server noi bo, heartbeat, approved machines, error logs, update info, tai exe user moi.

## File duoc sua

Da copy san trong `source/`:

- `presence_server.py`
- `presence_server.spec`
- `gk_pilepro/gk_core.py`

## Endpoint quan trong

- `/health`
- `/heartbeat`
- `/machines`
- `/approved-machines`
- `/error-logs`
- `/update-info`
- `/updates/user-exe`

## Khong tu y sua

- UI files trong `gk_pilepro/ui/`
- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`

## Can phoi hop khi

- Doi response update/approval/heartbeat: hoi `User Client` va `Admin Client`.
- Doi build server: hoi `Build System`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile presence_server.py
.\.venv\Scripts\pyinstaller.exe presence_server.spec
```

Exe output:

```text
dist\presence_server.exe
```
