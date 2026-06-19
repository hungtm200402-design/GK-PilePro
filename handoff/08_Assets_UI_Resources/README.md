# 08 Assets/UI Resources

## Pham vi

Logo, icon, splash, sidebar icon, button icon, anh trang tri.

## File/thu muc duoc sua

Trong `source/` co san cac file khai bao asset/build:

- `assets/`
- Cac hang asset trong `app.py`
- Cac muc `datas` trong `Tool_KL.spec`
- Cac muc `datas` trong `GK_PilePro_Admin.spec`

## Khong tu y sua

- OCR logic.
- Server logic.
- Excel export logic.

## Can phoi hop khi

- Them/xoa asset: hoi `Build System` de cap nhat spec.
- Doi UI user: hoi `User Client`.
- Doi UI admin: hoi `Admin Client`.

## Kiem tra reference truoc khi xoa asset

```bat
rg -n "ten_file_asset.png" app.py gk_pilepro Tool_KL.spec GK_PilePro_Admin.spec
```

## Build kiem tra sau khi doi asset

```bat
.\.venv\Scripts\pyinstaller.exe Tool_KL.spec
.\.venv\Scripts\pyinstaller.exe GK_PilePro_Admin.spec
```
