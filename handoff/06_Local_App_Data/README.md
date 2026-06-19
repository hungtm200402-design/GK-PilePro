# 06 Local App Data

## Pham vi

Noi luu settings, history, mapping templates, approval cache, logs, backup path.

## File duoc sua

Da copy san trong `source/`:

- `gk_pilepro/gk_core.py`
- Cac noi goi `app_data_path(...)`

## Data lien quan

- `tool_kl_settings.json`
- `tool_kl_history.json`
- `tool_kl_mapping_templates.json`
- `gk_pilepro_approval.json`
- `gk_pilepro_approved_machines.json`
- `logs/`
- `backups/`

## Khong tu y sua

- UI layout neu chi doi path/data.
- Server endpoint neu khong doi sync remote.

## Can phoi hop khi

- Doi settings/history UI: hoi `User Client`.
- Doi approved machine data: hoi `Admin Client` va `Presence/Update Server`.
- Doi backup/mapping templates: hoi `Excel/Data Files`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile gk_pilepro\gk_core.py app.py
```
